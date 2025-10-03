using System.CommandLine;
using System.CommandLine.Parsing;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelCalendar;

static class Program
{
    public static int Main(string[] args)
    {
        // If you use EPPlus for Noncommercial personal use.
        ExcelPackage.License
            .SetNonCommercialPersonal(
                "Helge Meyer"); //This will also set the Author property to the name provided in the argument.

        Console.WriteLine("Hello, World!");

        Option<FileInfo> fileOption = new("--output", ["-o"])
        {
            Description = "The Excel (xlsx) file to create the calendar.",
            Required = true
        };
        Option<FileInfo> extractOption = new("--extract")
        {
            Description = "The Excel (xlsx) file to create the calendar.",
            Required = true
        };
        RootCommand rootCommand = new("Sample app for System.CommandLine");
        rootCommand.Options.Add(fileOption);
        rootCommand.Options.Add(extractOption);

        var parseResult = rootCommand.Parse(args);

        if (parseResult.Errors.Count == 0 && parseResult.GetValue(extractOption) is { } extractFile
                                          && parseResult.GetValue(fileOption) is { } outputFile)
        {
            AnniversaryImport.Import(extractFile, outputFile);
        }
        else
        if (parseResult.Errors.Count == 0 && parseResult.GetValue(fileOption) is { } parsedFile)
        {
            CalendarExport.CreateCalendar(parsedFile);
            return 0;
        }
        foreach (ParseError parseError in parseResult.Errors)
        {
            Console.Error.WriteLine(parseError.Message);
        }
        return 1;
    }
}