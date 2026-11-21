using System.Runtime.InteropServices;
using ClosedXML.Excel;
using AriaDamInvoice.Models;
using System.IO;

namespace AriaDamInvoice.Services;

public sealed class ExcelContractService
{
    private const int FirstAnimalRow = 7;
    private const int InitialAnimalRows = 10;
    private const int FirstEarTagRow = 27;
    private const int InitialEarTagRows = 5;
    private readonly string _templatePath = Path.Combine(
        AppContext.BaseDirectory, "Templates", "InvoiceTemplate.xlsx");

    public (string ExcelPath, string PdfPath) Create(InvoiceData data, string outputDirectory)
    {
        ArgumentNullException.ThrowIfNull(data);
        if (!File.Exists(_templatePath))
            throw new FileNotFoundException(
                "قالب اکسل همراه برنامه پیدا نشد. برنامه را دوباره از مخزن دریافت و Build کنید.",
                _templatePath);
        if (!Directory.Exists(outputDirectory))
            throw new DirectoryNotFoundException("پوشه خروجی در دسترس نیست.");

        var safeBuyer = MakeSafeFileName(string.IsNullOrWhiteSpace(data.BuyerName)
            ? "بدون نام" : data.BuyerName);
        var baseName = $"{data.ContractDate.Replace('/', '-')}_{safeBuyer}";
        var excelPath = GetUniquePath(outputDirectory, baseName, ".xlsx");
        var pdfPath = Path.ChangeExtension(excelPath, ".pdf");

        using (var workbook = new XLWorkbook(_templatePath))
        {
            var sheet = workbook.Worksheet("قرارداد فروش دام");
            FillBuyerDetails(sheet, data);
            var summaryRow = ResizeAnimalRows(sheet, data.Animals.Count);
            FillAnimalRows(sheet, data.Animals, summaryRow);
            var signatureRow = ResizeEarTagRows(sheet, data.EarTags, data.Animals.Count);
            FillObligationsOnFront(sheet, data.Animals.Count);
            FillBackPage(workbook.Worksheet("پشت فاکتور"));

            var lastRow = signatureRow + 4;
            sheet.PageSetup.PrintAreas.Clear();
            sheet.PageSetup.PrintAreas.Add($"A1:H{lastRow}");
            sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            sheet.PageSetup.PaperSize = XLPaperSize.A4;
            sheet.PageSetup.PagesWide = 1;
            sheet.PageSetup.PagesTall = 0;
            sheet.PageSetup.Margins.Top = 0.35;
            sheet.PageSetup.Margins.Bottom = 0.35;
            sheet.PageSetup.Margins.Left = 0.35;
            sheet.PageSetup.Margins.Right = 0.35;

            var back = workbook.Worksheet("پشت فاکتور");
            back.PageSetup.PrintAreas.Clear();
            back.PageSetup.PrintAreas.Add("A1:H19");
            back.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            back.PageSetup.PaperSize = XLPaperSize.A4;
            back.PageSetup.PagesWide = 1;
            back.PageSetup.PagesTall = 1;

            workbook.SaveAs(excelPath);
        }

        ExcelPdfExporter.ExportWorkbook(excelPath, pdfPath);
        return (excelPath, pdfPath);
    }

    private static void FillBuyerDetails(IXLWorksheet sheet, InvoiceData data)
    {
        sheet.Cell("C3").Value = data.ContractDate;
        sheet.Cell("G3").Value = data.BuyerName;
        sheet.Cell("C4").Value = data.Phone;
        sheet.Cell("G4").Value = data.Address;
    }

    private static int ResizeAnimalRows(IXLWorksheet sheet, int count)
    {
        var rowCount = Math.Max(1, count);
        if (rowCount < InitialAnimalRows)
        {
            sheet.Rows(FirstAnimalRow + rowCount, FirstAnimalRow + InitialAnimalRows - 1).Delete();
        }
        else if (rowCount > InitialAnimalRows)
        {
            var extra = rowCount - InitialAnimalRows;
            sheet.Row(17).InsertRowsAbove(extra);
            for (var row = 17; row < 17 + extra; row++)
            {
                sheet.Row(15).CopyTo(sheet.Row(row));
                sheet.Range($"G{row}:H{row}").Merge();
                sheet.Row(row).Height = sheet.Row(15).Height;
            }
        }

        return FirstAnimalRow + rowCount;
    }

    private static void FillAnimalRows(IXLWorksheet sheet, IReadOnlyList<AnimalLine> animals, int summaryRow)
    {
        for (var i = 0; i < Math.Max(1, animals.Count); i++)
        {
            var rowNumber = FirstAnimalRow + i;
            sheet.Cell(rowNumber, 1).Value = i + 1;

            if (i >= animals.Count)
            {
                for (var col = 2; col <= 6; col++)
                    sheet.Cell(rowNumber, col).Clear(XLClearOptions.Contents);
                sheet.Cell(rowNumber, 7).Clear(XLClearOptions.Contents);
                continue;
            }

            var animal = animals[i];
            sheet.Cell(rowNumber, 2).Value = animal.Count;
            sheet.Cell(rowNumber, 3).Value = animal.Weight;
            sheet.Cell(rowNumber, 4).Value = animal.AnimalType;
            sheet.Cell(rowNumber, 5).Value = animal.Breed;
            sheet.Cell(rowNumber, 6).Value = animal.UnitPrice;
            sheet.Cell(rowNumber, 7).FormulaA1 = $"=B{rowNumber}*F{rowNumber}";
            sheet.Cell(rowNumber, 6).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(rowNumber, 7).Style.NumberFormat.Format = "#,##0";
        }

        sheet.Cell(summaryRow, 1).Value = "جمع کل:";
        sheet.Cell(summaryRow, 7).FormulaA1 = $"=SUM(G{FirstAnimalRow}:G{summaryRow - 1})";
        sheet.Cell(summaryRow, 7).Style.NumberFormat.Format = "#,##0";
    }

    private static int ResizeEarTagRows(IXLWorksheet sheet, IReadOnlyList<string> tags, int animalCount)
    {
        var neededRows = Math.Max(1, (int)Math.Ceiling(tags.Count / 8d));
        var titleRow = 16 + Math.Max(1, animalCount);
        var firstTagRow = titleRow + 1;
        var originalEnd = firstTagRow + InitialEarTagRows - 1;

        if (neededRows < InitialEarTagRows)
        {
            sheet.Rows(firstTagRow + neededRows, originalEnd).Delete();
        }
        else if (neededRows > InitialEarTagRows)
        {
            var extra = neededRows - InitialEarTagRows;
            var insertAt = originalEnd + 1;
            sheet.Row(insertAt).InsertRowsAbove(extra);
            for (var row = insertAt; row < insertAt + extra; row++)
            {
                sheet.Row(originalEnd).CopyTo(sheet.Row(row));
                sheet.Row(row).Height = sheet.Row(originalEnd).Height;
            }
        }

        for (var i = 0; i < tags.Count; i++)
        {
            var row = firstTagRow + i / 8;
            var column = 1 + i % 8;
            sheet.Cell(row, column).Value = tags[i];
            sheet.Cell(row, column).Style.Font.FontName = "B Yekan";
        }

        var signatureRow = firstTagRow + neededRows + 1;
        sheet.Cell(signatureRow, 1).Value = "امضاء خریدار";
        sheet.Cell(signatureRow, 5).Value = "امضاء فروشنده";
        return signatureRow;
    }

    private static void FillObligationsOnFront(IXLWorksheet sheet, int animalCount)
    {
        var titleRow = 9 + Math.Max(1, animalCount);
        var bodyStart = titleRow + 1;
        var bodyEnd = bodyStart + 4;

        sheet.Cell(titleRow, 1).Value = ContractText.ObligationsTitle;
        sheet.Range(titleRow, 1, titleRow, 8).Merge();
        sheet.Cell(titleRow, 1).Style.Font.Bold = true;
        sheet.Cell(titleRow, 1).Style.Font.FontName = "B Nazanin";
        sheet.Cell(titleRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

        foreach (var merged in sheet.MergedRanges
                     .Where(r => r.RangeAddress.FirstAddress.RowNumber >= bodyStart &&
                                 r.RangeAddress.LastAddress.RowNumber <= bodyEnd)
                     .ToList())
            merged.Unmerge();
        sheet.Range(bodyStart, 1, bodyEnd, 8).Merge();
        var body = sheet.Cell(bodyStart, 1);
        body.Value = ContractText.ObligationsBody;
        body.Style.Font.FontName = "B Nazanin";
        body.Style.Font.FontSize = 11;
        body.Style.Alignment.WrapText = true;
        body.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        body.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        for (var row = bodyStart; row <= bodyEnd; row++)
            sheet.Row(row).Height = 34;
    }

    private static void FillBackPage(IXLWorksheet sheet)
    {
        sheet.Cell("A2").Value = ContractText.ObligationsTitle;
        sheet.Cell("A3").Value =
            "کلیه نژاد دام های قید شده در ستون نژاد دام طبق شماره پلاک گوش دارای گواهی ژنتیک از آزمایشگاه می باشند.";
        sheet.Cell("A5").Value =
            "بند اول:دام در هنگام خروج از مجموعه سالم بوده و مسئولیت ان بعد از خروج بر عهده خریدار میباشد.";
        sheet.Cell("A8").Value =
            "بند دوم:خریدار موظف بوده برای حمل دام مجوزات لازم را اخذ نماید و مسئولیت ان بر عهده فروشنده نمیباشد.";
        sheet.Cell("A11").Value =
            "بند سوم:برای رعایت مسائل بهداشتی پس از خروج دام از مجموعه،به هیچ عنوان مرجوع یا تعویض نمیگردد.";
        sheet.Cell("A14").Value =
            "بند چهارم:ابستن بودن دام بعد از رویت و سونوگرافی مشتری ارسال میشود و به هر دلیل بعد از ارسال مسئولیت ان برعهده فروشنده نمیباشد.";
        sheet.Cell("A18").Value = "انشا الله خداوند به طرفین خیر و برکت عنایت فرماید.";

        foreach (var cell in new[] { "A2", "A3", "A5", "A8", "A11", "A14", "A18" })
        {
            sheet.Cell(cell).Style.Font.FontName = "B Nazanin";
            sheet.Cell(cell).Style.Alignment.WrapText = true;
            sheet.Cell(cell).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        sheet.Row(3).Height = 38;
        sheet.Row(5).Height = 30;
        sheet.Row(6).Height = 30;
        sheet.Row(8).Height = 30;
        sheet.Row(9).Height = 30;
        sheet.Row(11).Height = 30;
        sheet.Row(12).Height = 30;
        sheet.Row(14).Height = 30;
        sheet.Row(15).Height = 30;
        sheet.Row(16).Height = 30;
        sheet.Row(18).Height = 30;
        sheet.Row(19).Height = 30;
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
        try
        {
            excel = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("اجرای Microsoft Excel ممکن نشد.");
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            workbook = excel.Workbooks.Open(excelPath, ReadOnly: true);
            workbook.ForceFullCalculation = true;
            workbook.CalculateFullRebuild();
            workbook.ExportAsFixedFormat(0, pdfPath);
            workbook.Close(false);
            workbook = null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"فایل اکسل ساخته شد، اما تولید PDF با خطا مواجه شد. فایل اکسل در دسترس است.\n{ex.Message}", ex);
        }
        finally
        {
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