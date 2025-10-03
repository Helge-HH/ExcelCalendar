using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelCalendar;

static class CalendarExport
{
    public static void CreateCalendar(FileInfo workbookFile)
    {
        //var myDocumentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        //var workbookFile = Path.Combine(myDocumentsFolder, "MyWorkbook.xlsx");
        
        if (workbookFile.Exists)
            workbookFile.Delete();
        
        // Open the workbook (or create it if it doesn't exist)
        using var p = new ExcelPackage(workbookFile);
        // Get the Worksheet created in the previous code sample. 
        var ws= p.Workbook.Worksheets.Add("MySheet");

        ws.PrinterSettings.PaperSize = ePaperSize.A4;
        ws.PrinterSettings.Orientation = eOrientation.Landscape;
        //ws.PrinterSettings.FitToPage = true;
        
        var date = DateOnly.FromDateTime(DateTime.Now);
        for (var i = 1; i <= 10; i++)
        {
            CreateDay(ws, date.AddDays(i - 1), i);
        }

        // Save and close the package.
        p.Save();
    }

    private static void CreateDay(ExcelWorksheet ws, DateOnly date, int index)
    {
        const string dateFormatDayLong = "[$-407]dddd";
        const string dateFormatDayOfMonth = "[$-407]d.";
        const string dateFormatMonthName = "[$-407]mmmm";
        const string dateFormatYear = "[$-407]yyyy";

        const ExcelBorderStyle borderStyle = ExcelBorderStyle.Medium;
        const ExcelBorderStyle borderStyleDouble = ExcelBorderStyle.Double;
        
        var dayRange = ws.Cells[1, index, 27, index];
        dayRange.Style.Border.BorderAround(borderStyle, false);

        if (date.DayOfWeek == DayOfWeek.Sunday)
        {
            dayRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //dayRange.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Background1); 
            // dayRange.Style.Fill.BackgroundColor.SetColor(ExcelIndexedColor.Indexed61); 
            dayRange.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#EEECE1"));

            ws.Column(index).PageBreak = true;
        }   

        var cell = ws.Cells[1, index];
        cell.Style.Font.Size = 14;
            
        cell = ws.Cells[2, index];
        var address = cell.Address;
        cell.Value = date;
        cell.Style.Numberformat.Format = dateFormatDayLong;
        cell.Style.Font.Size = 12;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.Border.Top.Style = borderStyle;
        
        cell = ws.Cells[3, index];
        cell.Formula = "=" + address;
        cell.Style.Numberformat.Format = dateFormatDayOfMonth;
        cell.Style.Font.Size = 24;
        cell.Style.Font.Bold = true;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.Border.Top.Style = borderStyle;

        cell = ws.Cells[4, index];
        cell.Style.Border.Top.Style = borderStyleDouble;
        
        cell = ws.Cells[6, index];
        cell.Style.Border.Top.Style = borderStyleDouble;
        
        cell = ws.Cells[7, index];
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        cell.Style.Border.Top.Style = borderStyleDouble;
        cell.Style.Border.Bottom.Style = borderStyle;


        cell = ws.Cells[26, index];
        if (date.DayOfWeek == DayOfWeek.Tuesday || date.Day == 2)
        {
            cell.Formula = "=" + address;
            cell.Style.Numberformat.Format = dateFormatYear;
            cell.Style.Font.Size = 18;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Border.Left.Style = ExcelBorderStyle.None;
        }

        if (date.DayOfWeek == DayOfWeek.Monday ||  date.Day == 1)
        {
            cell.Formula = "=" + address;
            cell.Style.Numberformat.Format = dateFormatMonthName;
            cell.Style.Font.Size = 16;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Border.Right.Style = ExcelBorderStyle.None;
        }
        cell.Style.Border.Top.Style = borderStyle;
        
        cell = ws.Cells[27, index];
        if (date.DayOfWeek == DayOfWeek.Monday)
        {
            cell.Formula = $"CONCATENATE(ISOWEEKNUM({address}), \". Woche\")";
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        cell.Style.Border.Top.Style = borderStyle;
        
        ws.Columns[index].Width = 16.0;
        
    }
}