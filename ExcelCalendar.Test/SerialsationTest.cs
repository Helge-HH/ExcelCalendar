using System.Text.Json;

namespace ExcelCalendar.Test;

public class AnniversaryConverterTest
{
    [Fact]
    public void ConvertsToJsonTest()
    {
        var anniversaries = new []
        {
            new Anniversary(29, 3),
            new Anniversary(20, 10),
            new Anniversary(15, 11),
        };
        
        string json1 = JsonSerializer.Serialize(anniversaries, new JsonSerializerOptions { WriteIndented = true });

        Assert.Equal("[\n  \"29.3\",\n  \"20.10\",\n  \"15.11\"\n]", json1);
    }

    [Fact]
    public void ConvertsFromJson()
    {
        var anniversaryExpected = new []
        {
            new Anniversary(29, 3),
            new Anniversary(20, 10),
            new Anniversary(15, 11),
        };

        string json1 = "[\n  \"29.3\",\n  \"20.10\",\n  \"15.11\"\n]";
        
        var anniversaries = JsonSerializer.Deserialize<Anniversary[]>(json1);

        Assert.Equal(anniversaryExpected, anniversaries);
    }

    [Fact]
    public void CalendarDayCanConvertToJsonTest()
    {
        CalendarDay[] calendarDays =
        [
            new(name: "Helge", type: EventType.Birthday, date: new Anniversary(15, 11)),
            new(name: "Tag dt. Einheit", type: EventType.Holyday, date: new Anniversary(3, 10))
        ];
        
        string json1 = JsonSerializer.Serialize(calendarDays, new JsonSerializerOptions { WriteIndented = false });

        Assert.Equal("[{\"Name\":\"Helge\",\"Date\":\"15.11\",\"Event\":\"Birthday\"},{\"Name\":\"Tag dt. Einheit\",\"Date\":\"3.10\",\"Event\":\"Holyday\"}]", json1);
    }
}