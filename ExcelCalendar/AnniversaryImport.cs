using System.Diagnostics;
using System.Text.Json;
using OfficeOpenXml;

namespace ExcelCalendar;

internal static class AnniversaryImport
{
    public static void Import(FileInfo extractFile, FileInfo outputFile)
    {
        using var package = new ExcelPackage(extractFile);
        var ws = package.Workbook.Worksheets.FirstOrDefault();
        var list = new List<CalendarDay>();

        Console.WriteLine($"Dimension address: {ws.Dimension.Address}");
        Console.WriteLine($"Dimension from row: {ws.Dimension.Start.Row}");
        Console.WriteLine($"Dimension to row: {ws.Dimension.End.Row}");
        Console.WriteLine($"Dimension from column: {ws.Dimension.Start.Column}");
        Console.WriteLine($"Dimension to column: {ws.Dimension.End.Column}");        

        for(var i = 2; i <= ws.Dimension.End.Row; i++)
        {
            var name = ws.Cells[i, 1].Text;
            var type = ws.Cells[i, 2].Text;
            var date = ws.Cells[i, 3].GetCellValue<DateTime>();
            var anniverrary = new Anniversary(date.Day, date.Month);

            EventType eventType = type switch
            {
                "Geburtstag" => EventType.Birthday,
                "Feiertag" => EventType.Holyday,
                "Jahreszeit" => EventType.Season,
                "Hochzeitstag" => EventType.WeddingDay,
                "Todestag" => EventType.DayOfDeath,
                _ => EventType.Birthday
            };

            list.Add(new CalendarDay(name, eventType, anniverrary));
        }     
        var json = JsonSerializer.Serialize(list);
        File.WriteAllText(outputFile.FullName, json);
    }
}