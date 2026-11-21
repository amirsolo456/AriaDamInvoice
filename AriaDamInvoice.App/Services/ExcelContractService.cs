using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using AriaDamInvoice.Models;

namespace AriaDamInvoice.Services;

public sealed class ExcelContractService
{
    private const int FirstAnimalRow = 8;
    private const int TemplateAnimalRows = 10;
    private const int EarTagTitleRow = 20;
    private const int FirstEarTagRow = 21;
    private const int TemplateEarTagRows = 5;
    private const int ObligationsTitleRow = 27;
    private const int SignatureRow = 30;
    private const int LastTemplateRow = 33;
    private const double A4HeightPoints = 841.89;
    // The supplied reference workbook is configured for 64% print zoom.
    // Keeping this fixed makes the flexible spacer predictable on every PC.
    private const double PrintScale = 0.64d;
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

            // Reference workbook: header row 7, animal rows 8-17, total row 18.
            var animalDelta = ResizeAnimalRows(sheet, data.Animals.Count);
            var animalCount = Math.Max(1, data.Animals.Count);
            var totalRow = FirstAnimalRow + animalCount;
            FillAnimalRows(sheet, data.Animals, totalRow);

            // Reference workbook: title row 20, five tag rows (21-25). Eight tags fit per row.
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

        // Opening the PDF is a convenience step only. The invoice is already
        // successfully created at this point, so a missing PDF viewer must not
        // make the whole operation look like a failed invoice generation.
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = pdfPath,
                UseShellExecute = true
            });
        }
        catch (Exception)
        {
            // If Windows has no default PDF application (or its association is
            // broken), open Explorer and select the generated PDF instead.
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{pdfPath}\"",
                    UseShellExecute = true
                });
            }
            catch (Exception)
            {
                // Keep the successfully generated Excel/PDF files; do not
                // report a generation failure just because auto-open failed.
            }
        }

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
        sheet.Cell("C4").Value = data.ContractDate;
        sheet.Cell("G4").Value = data.BuyerName;
        sheet.Cell("C5").Value = data.Phone;
        sheet.Cell("G5").Value = data.Address;

        foreach (var address in new[] { "C4", "G4", "C5", "G5" })
            sheet.Cell(address).Style.Font.FontColor = XLColor.Black;

        // The approved reference workbook places the document-number label in G1.
        // H1 is the fixed value cell; keep this mapping explicit.
        var numberCell = sheet.Cell("H1");
        numberCell.Value = data.DocumentNumber.ToString(CultureInfo.InvariantCulture);
        numberCell.Style.Font.FontColor = XLColor.Black;
        numberCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
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
            sheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
        }

        sheet.Cell(totalRow, 1).Value = "جمع کل:";
        sheet.Cell(totalRow, 7).FormulaA1 =
            $"=SUM(G{FirstAnimalRow}:G{totalRow - 1})";
        sheet.Cell(totalRow, 7).Style.NumberFormat.Format = "#,##0";
    }

    // Ear-tag template rows 21-25 contain eight cells per row.
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
        sheet.Range(FirstAnimalRow, 1, totalRow - 1, 1).Style.Font.FontName = "B Yekan";
        sheet.Range(FirstAnimalRow, 1, totalRow - 1, 1).Style.Font.FontSize = 12;
        sheet.Range(FirstAnimalRow, 2, totalRow - 1, 3).Style.Font.FontName = "B Yekan";
        sheet.Range(FirstAnimalRow, 2, totalRow - 1, 3).Style.Font.FontSize = 12;
        sheet.Range(FirstAnimalRow, 6, totalRow, 7).Style.Font.FontName = "B Yekan";
        sheet.Range(FirstAnimalRow, 6, totalRow, 7).Style.Font.FontSize = 12;
        sheet.Cell("C5").Style.Font.FontName = "B Yekan";
        sheet.Cell("C4").Style.Font.FontName = "B Yekan";

        if (tagRows > 0)
        {
            var tags = sheet.Range(firstTagRow, 1, firstTagRow + tagRows - 1, 8);
            tags.Style.Font.FontName = "B Yekan";
            tags.Style.Font.FontSize = 12;
        }

        sheet.Row(7).Style.Font.FontName = "B Nazanin";
        sheet.Row(7).Style.Font.FontSize = 14;
        sheet.Row(7).Style.Font.Bold = true;
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
        // Keep the flexible spacer at its normal minimum height. Excel's one-page
        // scaling expands short contracts and scales down longer ones proportionally.
        sheet.Row(spacerRow).Height = DefaultRowHeight;

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
        sheet.PageSetup.CenterVertically = true;
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
            // Fit the complete print area to exactly one A4 page.
            // Excel scales short contracts up and longer contracts down.
            pageSetup.Zoom = false;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 1;
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = true;
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