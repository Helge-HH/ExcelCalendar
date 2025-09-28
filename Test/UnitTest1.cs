using System.Text.Json;
using epPlus1;

namespace Test;

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
}