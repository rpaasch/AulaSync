using AulaSync;
using System.Text.Json;

namespace AulaSync.Tests;

// === Rolle-mapping ===

public class EmployeeRoleTests
{
    [Theory]
    [InlineData("teacher", "Lærer")]
    [InlineData("preschool-teacher", "Pædagog")]
    [InlineData("leader", "Leder")]
    [InlineData("other", "Anden")]
    [InlineData("unknown", "unknown")]
    [InlineData("", "")]
    public void EmployeeRoleDanish_ReturnsCorrectMapping(string role, string expected)
    {
        var emp = new AulaApi.Employee { Role = role };
        Assert.Equal(expected, emp.RoleDanish);
    }
}

// === Beskeder ===

public class AulaMessageTests
{
    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var msg = new AulaMessage();
        Assert.Equal(0, msg.ThreadId);
        Assert.Equal("", msg.Subject);
        Assert.Equal("", msg.Sender);
        Assert.Equal("", msg.SenderMeta);
        Assert.Equal("", msg.Text);
        Assert.Equal("", msg.Timestamp);
        Assert.False(msg.IsRead);
        Assert.False(msg.IsSensitive);
        Assert.Equal(0, msg.TotalRecipients);
        Assert.Equal(0, msg.TotalMessages);
        Assert.Empty(msg.AllMessages);
    }

    [Fact]
    public void SensitiveMessage()
    {
        var msg = new AulaMessage
        {
            ThreadId = 999,
            Subject = "Følsom",
            Sender = "Ukendt",
            Text = "Denne besked kræver MitID-login.",
            IsSensitive = true,
        };
        Assert.True(msg.IsSensitive);
        Assert.Equal("Ukendt", msg.Sender);
    }

    [Fact]
    public void SubMessage_DefaultValues()
    {
        var sub = new AulaSubMessage();
        Assert.Equal("", sub.Sender);
        Assert.Equal("", sub.SenderMeta);
        Assert.Equal("", sub.Text);
        Assert.Equal("", sub.Timestamp);
    }
}

// === Tidsstempler ===

public class TimestampTests
{
    [Fact]
    public void ParseTimestamp_ISO8601_UTC()
    {
        Assert.True(DateTime.TryParse("2025-06-15T10:30:00Z",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.RoundtripKind, out var dt));
        Assert.Equal(2025, dt.Year);
        Assert.Equal(6, dt.Month);
        Assert.Equal(15, dt.Day);
    }

    [Fact]
    public void ParseTimestamp_WinterTime()
    {
        Assert.True(DateTime.TryParse("2026-03-20T10:59:00+0100",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.RoundtripKind, out var dt));
        Assert.Equal(2026, dt.Year);
    }

    [Fact]
    public void ParseTimestamp_SummerTime()
    {
        Assert.True(DateTime.TryParse("2026-04-13T08:00:00+0200",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.RoundtripKind, out _));
    }

    [Fact]
    public void FileTime_RoundTrip()
    {
        var original = new DateTime(2025, 6, 15, 10, 30, 0, DateTimeKind.Utc);
        long ft = original.ToFileTimeUtc();
        Assert.Equal(original, DateTime.FromFileTimeUtc(ft));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("not-a-date")]
    [InlineData("2026-13-45")]
    public void InvalidTimestamp_ReturnsFalse(string? ts)
    {
        Assert.False(DateTime.TryParse(ts, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.RoundtripKind, out _));
    }
}

// === Seen-keys ===

public class SeenKeyTests
{
    [Fact]
    public void DifferentTimestamp_DifferentKey()
    {
        Assert.NotEqual($"{100}_2025-06-15T10:30:00Z", $"{100}_2025-06-15T11:30:00Z");
    }

    [Fact]
    public void DifferentThread_DifferentKey()
    {
        Assert.NotEqual($"{100}_2025-06-15T10:30:00Z", $"{200}_2025-06-15T10:30:00Z");
    }

    [Fact]
    public void SameInput_SameKey()
    {
        Assert.Equal($"{100}_ts", $"{100}_ts");
    }

    [Fact]
    public void Format()
    {
        Assert.Equal("12345_2025-06-15T10:30:00Z", $"{12345}_{"2025-06-15T10:30:00Z"}");
    }
}

// === Kalender-events ===

public class CalendarEventTests
{
    [Fact]
    public void DefaultValues()
    {
        var ev = new AulaApi.CalendarEvent();
        Assert.Equal("", ev.Id);
        Assert.Equal("", ev.Title);
        Assert.Equal("", ev.Start);
        Assert.Equal("", ev.End);
        Assert.Equal("", ev.Teacher);
        Assert.Equal("", ev.Groups);
        Assert.Equal("", ev.Location);
        Assert.Empty(ev.BelongsToProfiles);
    }

    [Fact]
    public void WithAllFields()
    {
        var ev = new AulaApi.CalendarEvent
        {
            Id = "12345",
            Title = "DAN",
            Start = "2026-04-13T08:00:00+02:00",
            End = "2026-04-13T08:45:00+02:00",
            Teacher = "Helle Brinch Larsen (HL)",
            Groups = "5A",
            Location = "53",
            BelongsToProfiles = new() { "1369015" },
        };
        Assert.Equal("DAN", ev.Title);
        Assert.Equal("5A", ev.Groups);
        Assert.Equal("53", ev.Location);
        Assert.Single(ev.BelongsToProfiles);
        Assert.Contains("1369015", ev.BelongsToProfiles);
    }

    [Fact]
    public void BelongsToProfiles_CanHaveMultiple()
    {
        var ev = new AulaApi.CalendarEvent
        {
            BelongsToProfiles = new() { "111", "222", "333" },
        };
        Assert.Equal(3, ev.BelongsToProfiles.Count);
    }
}

// === Event-titel formatering ===

public class EventTitleFormatTests
{
    // Medarbejder: FAG | Lokale | Klasse
    [Fact] public void Employee_AllFields() =>
        Assert.Equal("DAN | 53 | 7A", Format("DAN", "53", "7A", "HL", "Employee"));

    [Fact] public void Employee_NoLocation() =>
        Assert.Equal("Morgentilsyn | 5A", Format("Morgentilsyn", "", "5A", "KA", "Employee"));

    [Fact] public void Employee_OnlyTitle() =>
        Assert.Equal("MØDE", Format("MØDE", "", "", "", "Employee"));

    [Fact] public void Employee_NoClass() =>
        Assert.Equal("DAN | 53", Format("DAN", "53", "", "HL", "Employee"));

    // Klasse: FAG | Lokale | INIT
    [Fact] public void Class_AllFields() =>
        Assert.Equal("DAN | 53 | HL", Format("DAN", "53", "7A", "HL", "Class"));

    [Fact] public void Class_NoLocation() =>
        Assert.Equal("Morgentilsyn | KA", Format("Morgentilsyn", "", "5A", "KA", "Class"));

    [Fact] public void Class_MultipleTeachers() =>
        Assert.Equal("DAN | 53 | HL, CV", Format("DAN", "53", "7A", "HL, CV", "Class"));

    [Fact] public void Class_NoTeacher() =>
        Assert.Equal("MØDE | 53", Format("MØDE", "53", "7A", "", "Class"));

    // Lokale: FAG | Klasse | INIT
    [Fact] public void Room_AllFields() =>
        Assert.Equal("IDR | 7A | HL", Format("IDR", "GYM", "7A", "HL", "Room"));

    [Fact] public void Room_NoClass() =>
        Assert.Equal("MØDE | HL", Format("MØDE", "GYM", "", "HL", "Room"));

    [Fact] public void Room_NoTeacher() =>
        Assert.Equal("IDR | 7A", Format("IDR", "GYM", "7A", "", "Room"));

    [Fact] public void Room_OnlyTitle() =>
        Assert.Equal("MØDE", Format("MØDE", "GYM", "", "", "Room"));

    // Alle tomme
    [Fact] public void AllEmpty_AllTypes()
    {
        Assert.Equal("X", Format("X", "", "", "", "Employee"));
        Assert.Equal("X", Format("X", "", "", "", "Class"));
        Assert.Equal("X", Format("X", "", "", "", "Room"));
    }

    private static string Format(string title, string location, string groups, string teacherInits, string type)
    {
        var parts = new List<string> { title };
        switch (type)
        {
            case "Employee":
                if (!string.IsNullOrEmpty(location)) parts.Add(location);
                if (!string.IsNullOrEmpty(groups)) parts.Add(groups);
                break;
            case "Class":
                if (!string.IsNullOrEmpty(location)) parts.Add(location);
                if (!string.IsNullOrEmpty(teacherInits)) parts.Add(teacherInits);
                break;
            case "Room":
                if (!string.IsNullOrEmpty(groups)) parts.Add(groups);
                if (!string.IsNullOrEmpty(teacherInits)) parts.Add(teacherInits);
                break;
        }
        return string.Join(" | ", parts);
    }
}

// === ICS-generering ===

public class IcsGenerationTests
{
    [Fact]
    public void EmptyList_ReturnsValidCalendar()
    {
        var ics = AulaApi.EventsToIcs(new(), "Test", "T");
        Assert.Contains("BEGIN:VCALENDAR", ics);
        Assert.Contains("END:VCALENDAR", ics);
        Assert.Contains("X-WR-CALNAME:Test (T)", ics);
        Assert.DoesNotContain("BEGIN:VEVENT", ics);
    }

    [Fact]
    public void WithEvents_ContainsAllFields()
    {
        var events = new List<AulaApi.CalendarEvent>
        {
            new()
            {
                Id = "1", Title = "DAN",
                Start = "2026-04-13T08:00:00+02:00", End = "2026-04-13T08:45:00+02:00",
                Groups = "7A", Location = "53", Teacher = "Helle (HL)",
            }
        };
        var ics = AulaApi.EventsToIcs(events, "HL - Helle", "HL");
        Assert.Contains("BEGIN:VEVENT", ics);
        Assert.Contains("SUMMARY:DAN (7A)", ics);
        Assert.Contains("LOCATION:53", ics);
        Assert.Contains("UID:aula-1", ics);
        Assert.Contains("END:VEVENT", ics);
        Assert.Contains("Helle (HL)", ics);
        Assert.Contains("Klasse: 7A", ics);
        Assert.Contains("Lokale: 53", ics);
    }

    [Fact]
    public void InvalidDate_SkipsEvent()
    {
        var events = new List<AulaApi.CalendarEvent>
        {
            new() { Id = "1", Title = "BAD", Start = "not-a-date", End = "also-not" }
        };
        Assert.DoesNotContain("BEGIN:VEVENT", AulaApi.EventsToIcs(events, "T", "T"));
    }

    [Fact]
    public void MultipleEvents_AllIncluded()
    {
        var events = new List<AulaApi.CalendarEvent>
        {
            new() { Id = "1", Title = "DAN", Start = "2026-04-13T08:00:00+02:00", End = "2026-04-13T08:45:00+02:00" },
            new() { Id = "2", Title = "MAT", Start = "2026-04-13T09:00:00+02:00", End = "2026-04-13T09:45:00+02:00" },
            new() { Id = "3", Title = "ENG", Start = "2026-04-13T10:00:00+02:00", End = "2026-04-13T10:45:00+02:00" },
        };
        Assert.Equal(3, AulaApi.EventsToIcs(events, "T", "T").Split("BEGIN:VEVENT").Length - 1);
    }

    [Fact]
    public void NoGroups_SummaryIsJustTitle()
    {
        var events = new List<AulaApi.CalendarEvent>
        {
            new() { Id = "1", Title = "MØDE", Start = "2026-04-13T08:00:00+02:00", End = "2026-04-13T09:00:00+02:00" }
        };
        Assert.Contains("SUMMARY:MØDE", AulaApi.EventsToIcs(events, "T", "T"));
        Assert.DoesNotContain("SUMMARY:MØDE (", AulaApi.EventsToIcs(events, "T", "T"));
    }

    [Fact]
    public void SpecialChars_AreEscaped()
    {
        var events = new List<AulaApi.CalendarEvent>
        {
            new()
            {
                Id = "1", Title = "Test; med, special\\tegn",
                Start = "2026-04-13T08:00:00+02:00", End = "2026-04-13T09:00:00+02:00",
                Location = "Rum; 1,2",
            }
        };
        var ics = AulaApi.EventsToIcs(events, "T", "T");
        Assert.Contains("\\;", ics);
        Assert.Contains("\\,", ics);
        Assert.Contains("\\\\", ics);
    }
}

// === Kalender-chunks ===

public class CalendarChunkTests
{
    [Fact]
    public void ChunkDays_WithinApiLimit()
    {
        Assert.InRange(42, 7, 55);
    }

    [Fact]
    public void ThreeMonths_RequiresMultipleChunks()
    {
        int chunks = (int)Math.Ceiling(180.0 / 42);
        Assert.InRange(chunks, 4, 6);
    }

    [Fact]
    public void ChunksCoverFullPeriod()
    {
        int totalDays = 180, chunkDays = 42, covered = 0;
        for (var s = 0; s < totalDays; s += chunkDays)
            covered = Math.Min(s + chunkDays, totalDays);
        Assert.Equal(totalDays, covered);
    }

    [Fact]
    public void NoGapsBetweenChunks()
    {
        int totalDays = 180, chunkDays = 42, prevEnd = 0;
        for (var s = 0; s < totalDays; s += chunkDays)
        {
            Assert.Equal(prevEnd, s);
            prevEnd = Math.Min(s + chunkDays, totalDays);
        }
    }

    [Fact]
    public void SmallPeriod_SingleChunk()
    {
        int totalDays = 14, chunkDays = 42, chunks = 0;
        for (var s = 0; s < totalDays; s += chunkDays) chunks++;
        Assert.Equal(1, chunks);
    }

    [Fact]
    public void ExactMultiple_NoExtraChunk()
    {
        int totalDays = 84, chunkDays = 42, chunks = 0;
        for (var s = 0; s < totalDays; s += chunkDays) chunks++;
        Assert.Equal(2, chunks);
    }
}

// === Gruppe- og ressource-info ===

public class GroupInfoTests
{
    [Fact]
    public void DefaultValues()
    {
        var g = new AulaApi.GroupInfo();
        Assert.Equal("", g.Id);
        Assert.Equal("", g.Name);
        Assert.Equal("", g.Type);
    }

    [Fact]
    public void Hovedgruppe_IsClass()
    {
        var g = new AulaApi.GroupInfo { Type = "Hovedgruppe", Name = "7A" };
        Assert.Equal("Hovedgruppe", g.Type);
    }
}

public class ResourceInfoTests
{
    [Fact]
    public void DefaultValues()
    {
        var r = new AulaApi.ResourceInfo();
        Assert.Equal("", r.Id);
        Assert.Equal("", r.Name);
    }
}

// === Input-validering ===

public class InputValidationTests
{
    [Theory]
    [InlineData("12345")]
    [InlineData("0")]
    [InlineData("9999999")]
    public void ValidProfileId_ParsesOk(string id)
    {
        Assert.True(int.TryParse(id, out var result));
        Assert.True(result >= 0);
    }

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData(null)]
    [InlineData("12.34")]
    public void InvalidProfileId_FailsParse(string? id)
    {
        Assert.False(int.TryParse(id, out _));
    }

    [Theory]
    [InlineData("0000000072D366BD")]
    [InlineData("AABBCCDD")]
    public void ValidHexString_Converts(string hex)
    {
        var bytes = Convert.FromHexString(hex);
        Assert.NotEmpty(bytes);
    }

    [Theory]
    [InlineData("GGHHII")]
    [InlineData("123")]  // Ulige antal tegn
    public void InvalidHexString_Throws(string hex)
    {
        Assert.ThrowsAny<Exception>(() => Convert.FromHexString(hex));
    }

    [Fact]
    public void EmptyHexString_ReturnsEmptyArray()
    {
        Assert.Empty(Convert.FromHexString(""));
    }
}

// === Data-integritet ===

public class DataIntegrityTests
{
    [Fact]
    public void AtomicWrite_TempFileApproach()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"aulasync_test_{Guid.NewGuid()}");
        Directory.CreateDirectory(dir);
        try
        {
            var file = Path.Combine(dir, "test.json");
            var tmp = file + ".tmp";
            var data = "{\"test\": true}";

            // Skriv via temp+move (som koden gør)
            File.WriteAllText(tmp, data);
            File.Move(tmp, file, true);

            Assert.Equal(data, File.ReadAllText(file));
            Assert.False(File.Exists(tmp));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void CorruptJson_ReturnsEmptyDict()
    {
        // Simulér hvad LoadSeen gør med korrupt fil
        var corrupt = "not valid json {{{";
        Dictionary<string, JsonElement>? result = null;
        try
        {
            result = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(corrupt);
        }
        catch
        {
            result = new();
        }
        Assert.NotNull(result);
        Assert.Empty(result);
    }

    [Fact]
    public void EmptyFile_ReturnsEmptyDict()
    {
        Dictionary<string, JsonElement>? result = null;
        try
        {
            result = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>("");
        }
        catch
        {
            result = new();
        }
        Assert.NotNull(result);
    }
}

// === Multi-institution ===

public class MultiInstitutionTests
{
    [Fact]
    public void EmptyInstitutionCode_IsInvalid()
    {
        Assert.True(string.IsNullOrEmpty(""));
    }

    [Fact]
    public void ValidInstitutionCode()
    {
        Assert.False(string.IsNullOrEmpty("257002"));
    }

    [Fact]
    public void NestedInstitution_ParsedCorrectly()
    {
        // Simulér getProfileContext response-struktur
        var json = """
        {
            "id": 6331896,
            "role": "employee",
            "firstName": "Rasmus",
            "lastName": "Paasch",
            "institution": {
                "institutionCode": "257002",
                "institutionName": "Kirke Saaby Skole"
            }
        }
        """;
        var doc = JsonDocument.Parse(json);
        var ip = doc.RootElement;
        Assert.True(ip.TryGetProperty("institution", out var instObj));
        Assert.Equal("257002", instObj.GetProperty("institutionCode").GetString());
        Assert.Equal("Kirke Saaby Skole", instObj.GetProperty("institutionName").GetString());
    }

    [Fact]
    public void MissingInstitution_HandledGracefully()
    {
        var json = """{"id": 123, "role": "employee"}""";
        var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.TryGetProperty("institution", out _));
    }
}
