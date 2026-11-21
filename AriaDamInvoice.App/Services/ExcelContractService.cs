using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using AriaDamInvoice.Models;

namespace AriaDamInvoice.Services;

public sealed class ExcelContractService
{
    private const int FirstAnimalRow = 7;
    private const int TemplateAnimalRows = 10;
    private const int TemplateEarTagTitleRow = 19;
    private const int FirstEarTagRow = 20;
    private const int TemplateEarTagRows = 5;
    private const int TemplateObligationsTitleRow = 26;
    private const int TemplateSignatureRow = 29;
    private const int TemplateLastRow = 32;

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

        var safeBuyer = MakeSafeFileName(string.IsNullOrWhiteSpace(data.BuyerName)
            ? "بدون نام" : data.BuyerName);
        var baseName = $"{data.ContractDate.Replace('/', '-')}_{safeBuyer}";
        var excelPath = GetUniquePath(outputDirectory, baseName, ".xlsx");
        var pdfPath = Path.ChangeExtension(excelPath, ".pdf");

        using (var templateStream = new FileStream(
                   _templatePath,
                   FileMode.Open,
                   FileAccess.Read,
                   FileShare.ReadWrite | FileShare.Delete))
        using (var templateMemory = new MemoryStream())
        {
            templateStream.CopyTo(templateMemory);
            templateMemory.Position = 0;

            using var workbook = new XLWorkbook(templateMemory);
            var sheet = workbook.Worksheet("قرارداد فروش دام");

            FillBuyerDetails(sheet, data);

            // Resize each variable section independently. Inserting/deleting rows at
            // the section boundary moves every following template section together.
            var animalDelta = ResizeAnimalRows(sheet, data.Animals.Count);
            var animalCount = Math.Max(1, data.Animals.Count);
            var summaryRow = FirstAnimalRow + animalCount;
            FillAnimalRows(sheet, data.Animals, summaryRow);

            var tagDelta = ResizeEarTagRows(sheet, data.EarTags, animalDelta);
            FillEarTagRows(sheet, data.EarTags, animalDelta);

            var obligationsTitleRow = TemplateObligationsTitleRow + animalDelta + tagDelta;
            FillObligations(sheet, obligationsTitleRow);

            var signatureRow = TemplateSignatureRow + animalDelta + tagDelta;
            // Preserve the signature cells and their merged layout from the template.
            sheet.Cell(signatureRow, 1).Value = "امضاء خریدار";
            sheet.Cell(signatureRow, 5).Value = "امضاء فروشنده";

            var lastRow = TemplateLastRow + animalDelta + tagDelta;
            ConfigurePrintLayout(sheet, lastRow);

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

    private static void FillBuyerDetails(IXLWorksheet sheet, InvoiceData data)
    {
        sheet.Cell("C3").Value = data.ContractDate;
        sheet.Cell("G3").Value = data.BuyerName;
        sheet.Cell("C4").Value = data.Phone;
        sheet.Cell("G4").Value = data.Address;
    }

    // Returns the number of rows inserted (positive) or deleted (negative).
    private static int ResizeAnimalRows(IXLWorksheet sheet, int requestedCount)
    {
        var count = Math.Max(1, requestedCount);
        var delta = count - TemplateAnimalRows;
        var summaryRow = FirstAnimalRow + TemplateAnimalRows;

        if (delta < 0)
        {
            sheet.Rows(FirstAnimalRow + count, summaryRow - 1).Delete();
        }
        else if (delta > 0)
        {
            sheet.Row(summaryRow).InsertRowsAbove(delta);

            // The original last animal row is the formatting prototype.
            for (var row = summaryRow; row < summaryRow + delta; row++)
            {
                sheet.Row(FirstAnimalRow + TemplateAnimalRows - 1).CopyTo(sheet.Row(row));
                sheet.Row(row).Clear(XLClearOptions.Contents);
                sheet.Cell(row, 1).Style = sheet.Cell(FirstAnimalRow + TemplateAnimalRows - 1, 1).Style;
                EnsureAnimalTotalMerge(sheet, row);
                sheet.Row(row).Height = sheet.Row(FirstAnimalRow + TemplateAnimalRows - 1).Height;
            }
        }

        return delta;
    }

    private static void EnsureAnimalTotalMerge(IXLWorksheet sheet, int row)
    {
        var exists = sheet.MergedRanges.Any(r =>
            r.RangeAddress.FirstAddress.RowNumber == row &&
            r.RangeAddress.FirstAddress.ColumnNumber == 7 &&
            r.RangeAddress.LastAddress.ColumnNumber == 8);
        if (!exists)
            sheet.Range(row, 7, row, 8).Merge();
    }

    private static void FillAnimalRows(
        IXLWorksheet sheet,
        IReadOnlyList<AnimalLine> animals,
        int summaryRow)
    {
        var count = Math.Max(1, animals.Count);

        for (var i = 0; i < count; i++)
        {
            var row = FirstAnimalRow + i;
            sheet.Cell(row, 1).Value = i + 1;

            if (i >= animals.Count)
            {
                for (var col = 2; col <= 6; col++)
                    sheet.Cell(row, col).Clear(XLClearOptions.Contents);
                continue;
            }

            var animal = animals[i];
            sheet.Cell(row, 2).Value = animal.Count;
            sheet.Cell(row, 3).Value = animal.Weight;
            sheet.Cell(row, 4).Value = animal.AnimalType;
            sheet.Cell(row, 5).Value = animal.Breed;
            sheet.Cell(row, 6).Value = animal.UnitPrice;
            sheet.Cell(row, 7).FormulaA1 = $"=B{row}*F{row}";
            sheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
        }

        sheet.Cell(summaryRow, 1).Value = "جمع کل:";
        sheet.Cell(summaryRow, 7).FormulaA1 =
            $"=SUM(G{FirstAnimalRow}:G{summaryRow - 1})";
        sheet.Cell(summaryRow, 7).Style.NumberFormat.Format = "#,##0";
    }

    // Ear tags use the template's 8-column grid: 8 tags per row.
    // Returns the number of rows inserted (positive) or deleted (negative).
    private static int ResizeEarTagRows(
        IXLWorksheet sheet,
        IReadOnlyList<string> tags,
        int animalDelta)
    {
        var neededRows = Math.Max(1, (int)Math.Ceiling(tags.Count / 8d));
        var firstTagRow = FirstEarTagRow + animalDelta;
        var currentRowsEnd = firstTagRow + TemplateEarTagRows - 1;
        var delta = neededRows - TemplateEarTagRows;

        if (delta < 0)
        {
            sheet.Rows(firstTagRow + neededRows, currentRowsEnd).Delete();
        }
        else if (delta > 0)
        {
            var insertAt = currentRowsEnd + 1;
            var prototypeRow = currentRowsEnd;
            sheet.Row(insertAt).InsertRowsAbove(delta);

            for (var row = insertAt; row < insertAt + delta; row++)
            {
                sheet.Row(prototypeRow).CopyTo(sheet.Row(row));
                sheet.Row(row).Clear(XLClearOptions.Contents);
                sheet.Row(row).Height = sheet.Row(prototypeRow).Height;
            }
        }

        return delta;
    }

    private static void FillEarTagRows(
        IXLWorksheet sheet,
        IReadOnlyList<string> tags,
        int animalDelta)
    {
        var neededRows = Math.Max(1, (int)Math.Ceiling(tags.Count / 8d));
        var firstRow = FirstEarTagRow + animalDelta;

        for (var rowOffset = 0; rowOffset < neededRows; rowOffset++)
        {
            var row = firstRow + rowOffset;
            for (var col = 1; col <= 8; col++)
                sheet.Cell(row, col).Clear(XLClearOptions.Contents);
        }

        for (var i = 0; i < tags.Count; i++)
        {
            var row = firstRow + i / 8;
            var col = 1 + i % 8;
            sheet.Cell(row, col).Value = tags[i];
            sheet.Cell(row, col).Style.Font.FontName = "B Yekan";
        }
    }

    // Keep the obligations area as the template's original single merged body row.
    // Do not merge several worksheet rows: doing so was a source of PDF layout drift.
    private static void FillObligations(IXLWorksheet sheet, int titleRow)
    {
        var bodyRow = titleRow + 1;

        sheet.Cell(titleRow, 1).Value = ContractText.ObligationsTitle;
        sheet.Cell(titleRow, 1).Style.Font.Bold = true;
        sheet.Cell(titleRow, 1).Style.Font.FontName = "B Nazanin";
        sheet.Cell(titleRow, 1).Style.Alignment.Horizontal =
            XLAlignmentHorizontalValues.Right;

        sheet.Cell(bodyRow, 1).Value = ContractText.ObligationsBody;
        var body = sheet.Cell(bodyRow, 1);
        body.Style.Font.FontName = "B Nazanin";
        body.Style.Font.FontSize = 11;
        body.Style.Alignment.WrapText = true;
        body.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        body.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        // Keep the original template row height, avoiding an oversized merged block.
    }

    private static void ConfigurePrintLayout(IXLWorksheet sheet, int lastRow)
    {
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
        while (File.Exists(candidate))
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

            // Export only the contract sheet and force a single portrait A4 page.
            contractSheet = workbook.Worksheets("قرارداد فروش دام");
            pageSetup = contractSheet.PageSetup;
            pageSetup.PaperSize = 9; // xlPaperA4
            pageSetup.Orientation = 1; // xlPortrait
            pageSetup.Zoom = false;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 1;
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