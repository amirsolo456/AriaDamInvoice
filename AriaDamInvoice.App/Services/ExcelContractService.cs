using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using AriaDamInvoice.Models;

namespace AriaDamInvoice.Services;

public sealed class ExcelContractService
{
    private const int FirstAnimalRow = 7;
    private const int TemplateAnimalRows = 10;
    private const int EarTagTitleRow = 19;
    private const int FirstEarTagRow = 20;
    private const int TemplateEarTagRows = 5;
    private const int ObligationsTitleRow = 26;
    private const int SignatureRow = 29;
    private const int LastTemplateRow = 32;
    private const double A4HeightPoints = 841.89;
    private const double DefaultRowHeight = 15d;

    private readonly string _templatePath;

    public ExcelContractService(string templatePath)
    {
        _templatePath = templatePath;
    }

    public (string ExcelPath, string PdfPath) Create(InvoiceData data, string outputDirectory)
    {
        ArgumentNullException.ThrowIfNull(data);

        if (!File.Exists(_templatePath))
            throw new FileNotFoundException(
                "فایل اکسل مرجع در مسیر ذخیره‌شده پیدا نشد. از داشبورد مسیر فایل مرجع را دوباره انتخاب کنید.",
                _templatePath);

        if (!Directory.Exists(outputDirectory))
            throw new DirectoryNotFoundException("پوشه خروجی در دسترس نیست.");

        var buyer = MakeSafeFileName(string.IsNullOrWhiteSpace(data.BuyerName)
            ? "بدون نام" : data.BuyerName);
        var baseName = $"{data.DocumentNumber}_{data.ContractDate.Replace('/', '-')}_{buyer}";
        var excelPath = GetUniquePath(outputDirectory, baseName, ".xlsx");
        var pdfPath = Path.ChangeExtension(excelPath, ".pdf");

        using (var templateMemory = CreateTemplateSnapshot())
        {
            templateMemory.Position = 0;
            using var workbook = new XLWorkbook(templateMemory);
            var sheet = workbook.Worksheet("قرارداد فروش دام");

            FillBuyerDetails(sheet, data);

            // Animal table: rows 7-16, total row 17 in the original template.
            var animalDelta = ResizeAnimalRows(sheet, data.Animals.Count);
            var animalCount = Math.Max(1, data.Animals.Count);
            var totalRow = FirstAnimalRow + animalCount;
            FillAnimalRows(sheet, data.Animals, totalRow);

            // The ear-tag block is independently sized. Eight tags fit per row.
            var earTagTitle = EarTagTitleRow + animalDelta;
            sheet.Cell(earTagTitle, 1).Value = "پلاک گوش دام ارسالی";
            var tagDelta = ResizeEarTagRows(sheet, data.EarTags.Count, animalDelta);
            var tagRows = Math.Max(1, (int)Math.Ceiling(data.EarTags.Count / 8d));
            var firstTagRow = FirstEarTagRow + animalDelta;
            FillEarTagRows(sheet, data.EarTags, firstTagRow, tagRows);

            var totalDelta = animalDelta + tagDelta;
            var obligationsTitle = ObligationsTitleRow + totalDelta;
            FillObligations(sheet, obligationsTitle);

            var spacerRow = SignatureRow - 1 + totalDelta;
            var signatureRow = SignatureRow + totalDelta;
            var lastRow = LastTemplateRow + totalDelta;

            // Preserve the signature and footer rows from the reference workbook.
            sheet.Cell(signatureRow, 1).Value = "امضاء خریدار";
            sheet.Cell(signatureRow, 5).Value = "امضاء فروشنده";
            ConfigureTypography(sheet, lastRow, totalRow, firstTagRow, tagRows,
                earTagTitle, obligationsTitle);
            ConfigurePrintLayout(sheet, lastRow, spacerRow);

            workbook.SaveAs(excelPath);
        }

        ExcelPdfExporter.ExportWorkbook(excelPath, pdfPath);

        Process.Start(new ProcessStartInfo
        {
            FileName = pdfPath,
            UseShellExecute = true
        });

        return (excelPath, pdfPath);
    }

    private MemoryStream CreateTemplateSnapshot()
    {
        try
        {
            var memory = new MemoryStream();
            using var stream = new FileStream(
                _templatePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            stream.CopyTo(memory);
            memory.Position = 0;
            return memory;
        }
        catch (IOException ex)
        {
            throw new IOException(
                "فایل اکسل مرجع در حال حاضر توسط برنامه دیگری قفل شده است. لطفاً فایل مرجع را در Excel ببندید و دوباره صدور قرارداد را انجام دهید.", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new UnauthorizedAccessException(
                "دسترسی خواندن فایل اکسل مرجع ممکن نیست. سطح دسترسی فایل را بررسی کنید.", ex);
        }
    }

    private static void FillBuyerDetails(IXLWorksheet sheet, InvoiceData data)
    {
        sheet.Cell("C3").Value = data.ContractDate;
        sheet.Cell("G3").Value = data.BuyerName;
        sheet.Cell("C4").Value = data.Phone;
        sheet.Cell("G4").Value = data.Address;

        foreach (var address in new[] { "C3", "G3", "C4", "G4" })
            sheet.Cell(address).Style.Font.FontColor = XLColor.Black;

        // Search the complete used range rather than assuming the label is inside A1:H5.
        // This also supports templates whose header columns or merged cells were rearranged.
        var usedRange = sheet.RangeUsed()
            ?? throw new InvalidDataException("فایل اکسل مرجع خالی است یا محدوده استفاده‌شده ندارد.");

        static string NormalizeLabel(string value)
        {
            return value
                .Replace("ي", "ی")
                .Replace("ك", "ک")
                .Replace("\u200c", "")
                .Replace(" ", "")
                .Replace(":", "")
                .Replace("：", "")
                .Trim();
        }

        var documentLabel = usedRange.Cells()
            .FirstOrDefault(cell =>
                NormalizeLabel(cell.GetString()).Contains("شمارهسند", StringComparison.Ordinal));

        if (documentLabel is null)
            throw new InvalidDataException(
                "عنوان «شماره سند» در هیچ‌یک از سلول‌های فایل اکسل مرجع پیدا نشد. " +
                "عبارت «شماره سند» را داخل یک سلول واقعی اکسل (نه Text Box) قرار دهید و فایل را ذخیره کنید.");

        var labelMerge = sheet.MergedRanges.FirstOrDefault(range =>
            range.RangeAddress.FirstAddress.RowNumber == documentLabel.Address.RowNumber &&
            range.RangeAddress.FirstAddress.ColumnNumber == documentLabel.Address.ColumnNumber);

        var labelFirstColumn = labelMerge?.RangeAddress.FirstAddress.ColumnNumber
                               ?? documentLabel.Address.ColumnNumber;
        var labelLastColumn = labelMerge?.RangeAddress.LastAddress.ColumnNumber
                              ?? documentLabel.Address.ColumnNumber;
        var lastUsedColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;

        // Prefer the cell immediately to the left of the label; if the label is
        // already at the left edge, use the first cell after its merged range.
        var targetColumn = labelFirstColumn > 1
            ? labelFirstColumn - 1
            : labelLastColumn + 1;

        if (targetColumn > sheet.ColumnCount())
            throw new InvalidDataException(
                "در کنار عنوان شماره سند فضای کافی برای درج شماره وجود ندارد. فایل مرجع را اصلاح کنید.");

        var numberCell = sheet.Cell(documentLabel.Address.RowNumber, targetColumn);
        var numberMerge = sheet.MergedRanges.FirstOrDefault(range =>
            range.Contains(numberCell));
        if (numberMerge is not null)
            numberCell = sheet.Cell(
                numberMerge.RangeAddress.FirstAddress.RowNumber,
                numberMerge.RangeAddress.FirstAddress.ColumnNumber);

        numberCell.Value = data.DocumentNumber.ToString(CultureInfo.InvariantCulture);
        numberCell.Style.Font.FontColor = XLColor.Black;
        numberCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        numberCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static int ResizeAnimalRows(IXLWorksheet sheet, int requestedCount)
    {
        var count = Math.Max(1, requestedCount);
        var delta = count - TemplateAnimalRows;
        var originalTotalRow = FirstAnimalRow + TemplateAnimalRows;

        if (delta < 0)
        {
            sheet.Rows(FirstAnimalRow + count, originalTotalRow - 1).Delete();
        }
        else if (delta > 0)
        {
            var sourceRow = originalTotalRow - 1;
            sheet.Row(originalTotalRow).InsertRowsAbove(delta);

            for (var row = originalTotalRow; row < originalTotalRow + delta; row++)
            {
                sheet.Row(sourceRow).CopyTo(sheet.Row(row));
                sheet.Row(row).Clear(XLClearOptions.Contents);
                EnsureTotalCellMerge(sheet, row);
                sheet.Row(row).Height = sheet.Row(sourceRow).Height;
            }
        }

        return delta;
    }

    private static void EnsureTotalCellMerge(IXLWorksheet sheet, int row)
    {
        var merged = sheet.MergedRanges.Any(range =>
            range.RangeAddress.FirstAddress.RowNumber == row &&
            range.RangeAddress.FirstAddress.ColumnNumber == 7 &&
            range.RangeAddress.LastAddress.ColumnNumber == 8);

        if (!merged)
            sheet.Range(row, 7, row, 8).Merge();
    }

    private static void FillAnimalRows(
        IXLWorksheet sheet, IReadOnlyList<AnimalLine> animals, int totalRow)
    {
        var count = Math.Max(1, animals.Count);

        for (var index = 0; index < count; index++)
        {
            var row = FirstAnimalRow + index;
            sheet.Cell(row, 1).Value = index + 1;

            if (index >= animals.Count)
            {
                for (var column = 2; column <= 6; column++)
                    sheet.Cell(row, column).Clear(XLClearOptions.Contents);
                sheet.Cell(row, 7).FormulaA1 = string.Empty;
                continue;
            }

            var animal = animals[index];
            sheet.Range(row, 1, row, 8).Style.Font.FontColor = XLColor.Black;
            sheet.Cell(row, 2).Value = animal.Count;
            sheet.Cell(row, 3).Value = animal.Weight;
            sheet.Cell(row, 4).Value = animal.AnimalType;
            sheet.Cell(row, 5).Value = animal.Breed;
            sheet.Cell(row, 6).Value = animal.UnitPrice;
            sheet.Cell(row, 7).FormulaA1 = $"=IF(OR(B{row}=\"\",F{row}=\"\"),\"\",B{row}*F{row})";
            sheet.Cell(row, 8).Value = animal.Description;
            sheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
        }

        sheet.Cell(totalRow, 1).Value = "جمع کل:";
        sheet.Cell(totalRow, 7).FormulaA1 =
            $"=SUM(G{FirstAnimalRow}:G{totalRow - 1})";
        sheet.Cell(totalRow, 7).Style.NumberFormat.Format = "#,##0";
    }

    // Ear-tag template rows 20-24 contain eight cells per row.
    private static int ResizeEarTagRows(IXLWorksheet sheet, int tagCount, int animalDelta)
    {
        var neededRows = Math.Max(1, (int)Math.Ceiling(tagCount / 8d));
        var firstRow = FirstEarTagRow + animalDelta;
        var currentLastRow = firstRow + TemplateEarTagRows - 1;
        var delta = neededRows - TemplateEarTagRows;

        if (delta < 0)
        {
            sheet.Rows(firstRow + neededRows, currentLastRow).Delete();
        }
        else if (delta > 0)
        {
            var insertAt = currentLastRow + 1;
            var sourceRow = currentLastRow;
            sheet.Row(insertAt).InsertRowsAbove(delta);

            for (var row = insertAt; row < insertAt + delta; row++)
            {
                sheet.Row(sourceRow).CopyTo(sheet.Row(row));
                sheet.Row(row).Clear(XLClearOptions.Contents);
                sheet.Row(row).Height = sheet.Row(sourceRow).Height;
            }
        }

        return delta;
    }

    private static void FillEarTagRows(
        IXLWorksheet sheet, IReadOnlyList<string> tags, int firstRow, int rowCount)
    {
        for (var offset = 0; offset < rowCount; offset++)
        {
            var row = firstRow + offset;
            for (var column = 1; column <= 8; column++)
                sheet.Cell(row, column).Clear(XLClearOptions.Contents);
        }

        for (var index = 0; index < tags.Count; index++)
        {
            var row = firstRow + index / 8;
            var column = 1 + index % 8;
            sheet.Cell(row, column).Value = tags[index];
            sheet.Cell(row, column).Style.Font.FontColor = XLColor.Black;
        }
    }

    private static void FillObligations(IXLWorksheet sheet, int titleRow)
    {
        var bodyRow = titleRow + 1;
        sheet.Cell(titleRow, 1).Value = ContractText.ObligationsTitle;
        sheet.Cell(bodyRow, 1).Value = ContractText.ObligationsBody;

        var body = sheet.Cell(bodyRow, 1);
        body.Style.Alignment.WrapText = true;
        body.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        body.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        body.Style.Font.FontName = "B Nazanin";
        body.Style.Font.FontSize = 14;
    }

    private static void ConfigureTypography(
        IXLWorksheet sheet, int lastRow, int totalRow, int firstTagRow, int tagRows,
        int tagTitleRow, int obligationsTitleRow)
    {
        var all = sheet.Range(1, 1, lastRow, 8);
        all.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        all.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        all.Style.Font.FontName = "B Nazanin";
        all.Style.Font.FontSize = 14;

        // Numeric content uses B Yekan and a readable size.
        sheet.Range(7, 1, totalRow - 1, 1).Style.Font.FontName = "B Yekan";
        sheet.Range(7, 1, totalRow - 1, 1).Style.Font.FontSize = 12;
        sheet.Range(7, 2, totalRow - 1, 3).Style.Font.FontName = "B Yekan";
        sheet.Range(7, 2, totalRow - 1, 3).Style.Font.FontSize = 12;
        sheet.Range(7, 6, totalRow, 7).Style.Font.FontName = "B Yekan";
        sheet.Range(7, 6, totalRow, 7).Style.Font.FontSize = 12;
        sheet.Cell("C3").Style.Font.FontName = "B Yekan";
        sheet.Cell("C4").Style.Font.FontName = "B Yekan";

        if (tagRows > 0)
        {
            var tags = sheet.Range(firstTagRow, 1, firstTagRow + tagRows - 1, 8);
            tags.Style.Font.FontName = "B Yekan";
            tags.Style.Font.FontSize = 12;
        }

        sheet.Row(6).Style.Font.FontName = "B Nazanin";
        sheet.Row(6).Style.Font.FontSize = 14;
        sheet.Row(6).Style.Font.Bold = true;
        sheet.Row(tagTitleRow).Style.Font.FontName = "B Nazanin";
        sheet.Row(tagTitleRow).Style.Font.FontSize = 16;
        sheet.Row(tagTitleRow).Style.Font.Bold = true;
        sheet.Row(obligationsTitleRow).Style.Font.FontName = "B Nazanin";
        sheet.Row(obligationsTitleRow).Style.Font.FontSize = 16;
        sheet.Row(obligationsTitleRow).Style.Font.Bold = true;
        sheet.Row(obligationsTitleRow + 1).Style.Font.FontName = "B Nazanin";
        sheet.Row(obligationsTitleRow + 1).Style.Font.FontSize = 14;
    }

    private static void ConfigurePrintLayout(IXLWorksheet sheet, int lastRow, int spacerRow)
    {
        const double marginInPoints = 18d; // 0.25 inch
        const double safetySpace = 18d;
        const double minimumSpacerHeight = 15d;

        // Stretch the blank row immediately above the signature block. This anchors
        // signatures and the address/footer at the bottom of A4 when content is short.
        double occupiedHeight = 0;
        for (var row = 1; row <= lastRow; row++)
        {
            if (row == spacerRow)
                continue;

            var height = sheet.Row(row).Height;
            occupiedHeight += height > 0 ? height : DefaultRowHeight;
        }

        var printableHeight = A4HeightPoints - (2 * marginInPoints) - safetySpace;
        sheet.Row(spacerRow).Height =
            Math.Max(minimumSpacerHeight, printableHeight - occupiedHeight);

        sheet.PageSetup.PrintAreas.Clear();
        sheet.PageSetup.PrintAreas.Add($"A1:H{lastRow}");
        sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
        sheet.PageSetup.PagesWide = 1;
        sheet.PageSetup.PagesTall = 1;
        sheet.PageSetup.Margins.Top = 0.25;
        sheet.PageSetup.Margins.Bottom = 0.25;
        sheet.PageSetup.Margins.Left = 0.25;
        sheet.PageSetup.Margins.Right = 0.25;
        sheet.PageSetup.CenterHorizontally = true;
        sheet.PageSetup.CenterVertically = false;
    }

    private static string MakeSafeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(value.Select(ch => invalid.Contains(ch) ? '_' : ch)).Trim();
    }

    private static string GetUniquePath(string directory, string baseName, string extension)
    {
        var candidate = Path.Combine(directory, baseName + extension);
        var index = 2;
        while (File.Exists(candidate) || File.Exists(Path.ChangeExtension(candidate, ".pdf")))
            candidate = Path.Combine(directory, $"{baseName}_{index++}{extension}");
        return candidate;
    }
}

internal static class ExcelPdfExporter
{
    public static void ExportWorkbook(string excelPath, string pdfPath)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType is null)
            throw new InvalidOperationException(
                "Microsoft Excel روی این رایانه نصب نیست یا ثبت ویندوزی آن کامل نیست.");

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? contractSheet = null;
        dynamic? pageSetup = null;

        try
        {
            excel = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("اجرای Microsoft Excel ممکن نشد.");
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            workbook = excel.Workbooks.Open(excelPath, ReadOnly: true);
            workbook.ForceFullCalculation = true;
            excel.CalculateFullRebuild();

            contractSheet = workbook.Worksheets("قرارداد فروش دام");
            pageSetup = contractSheet.PageSetup;
            pageSetup.PaperSize = 9; // xlPaperA4
            pageSetup.Orientation = 1; // xlPortrait
            pageSetup.Zoom = false;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 1;
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = false;
            contractSheet.ExportAsFixedFormat(0, pdfPath);

            workbook.Close(false);
            workbook = null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"فایل اکسل ساخته شد، اما تولید PDF با خطا مواجه شد. فایل اکسل در دسترس است.\\n{ex.Message}", ex);
        }
        finally
        {
            if (pageSetup is not null)
                try { Marshal.FinalReleaseComObject(pageSetup); } catch { }
            if (contractSheet is not null)
                try { Marshal.FinalReleaseComObject(contractSheet); } catch { }
            if (workbook is not null)
            {
                try { workbook.Close(false); } catch { }
                try { Marshal.FinalReleaseComObject(workbook); } catch { }
            }
            if (excel is not null)
            {
                try { excel.Quit(); } catch { }
                try { Marshal.FinalReleaseComObject(excel); } catch { }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}