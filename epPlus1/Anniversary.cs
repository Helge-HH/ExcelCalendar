using System.Text.Json;
using System.Text.Json.Serialization;

namespace epPlus1;

[JsonConverter(typeof(AnniversaryConverter))]
public struct Anniversary
{
    [JsonPropertyName("day")]
    public int Day { get; set; }   
    [JsonPropertyName("month")]
    public int Month { get; set; }

    public Anniversary(int day, int month)
    {
        if(month < 1 || month > 12)
            throw new ArgumentOutOfRangeException(nameof(month), "Month must be between 1 and 12.");
        if(day < 1 || day > DateTime.DaysInMonth(2000, month))
            throw new ArgumentOutOfRangeException(nameof(day), "Day must be a valid date.");
        Day = day;
        Month = month;
    }
    
    public override string ToString() => $"{Day}.{Month}";
    
    public DateOnly ToDate(int year) => new(year, Month, Day);
}

public class AnniversaryConverter : JsonConverter<Anniversary>
{
    public override Anniversary Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        var str = reader.GetString() ?? throw new JsonException("Null string for anniversary.");
        var parts = str.Split('.');
        if(parts.Length != 2)
            throw new JsonException($"Invalid format {str}. expected 'day.month'");
        if(!int.TryParse(parts[0], out var day) || !int.TryParse(parts[1], out var month))
            throw new JsonException($"Invalid numbers in {str}");
        return new Anniversary(day, month);
    }

    public override void Write(Utf8JsonWriter writer, Anniversary value, JsonSerializerOptions options)
    {
        writer.WriteStringValue($"{value.Day}.{value.Month}");
    }
}