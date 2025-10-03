using System.Text.Json.Serialization;

namespace ExcelCalendar;

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum EventType {Holyday, Birthday, DayOfDeath, Season, WeddingDay}

public struct CalendarDay
{
    public CalendarDay(string name, EventType type, Anniversary date)
    {
        Name = name;
        Type = type;
        Date = date;
    }

    [JsonPropertyName("Name")]
    public string Name {  get; set; }
    
    [JsonPropertyName("Date")]
    public Anniversary Date {  get; set; }
    
    [JsonPropertyName("Event")]
    public EventType Type {  get; set; }
}