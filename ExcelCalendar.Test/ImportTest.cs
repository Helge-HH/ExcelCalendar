
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelCalendar.Test;

public class Tests
{
    private static readonly VerifySettings Settings;
    
    static Tests()
    {
        Settings = new VerifySettings();
        Settings.UseDirectory("snapshots");
    }
}

public class ImportTest
{
    private readonly VerifySettings Settings;

    public ImportTest()
    {
        Settings = new VerifySettings();
        Settings.UseDirectory("snapshots");

        Environment.SetEnvironmentVariable("EPPlusLicense", "NonCommercialPersonal:Helge Meyer");    
    }
    
    [Fact]
    public Task ImportShouldMatchTest()
    {
        var extractFile = new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "SampleAnniversaries.xlsx"));
        var outputFile = new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "ExcelCalendar.json"));
        AnniversaryImport.Import(extractFile, outputFile);
        
        var json = File.ReadAllText(outputFile.FullName);
        var jsonIndent = JsonNode.Parse(json)?.ToJsonString(new JsonSerializerOptions(){WriteIndented = true});

        return Verify(jsonIndent, Settings);
    }
}