using AulaSync.Core;

namespace AulaSync.Tests;

public class AulaParsersTests
{
    [Fact]
    public void ProfileContext()
    {
        var p = AulaParsers.ParseProfileContext(Fixture.Read("profileContext.json"));
        Assert.Equal(new Profile("9000001", "Test Bruger", "TB", "leader", "101001", "Testskolen"), p);
    }

    // Rollen er institutionsrollen (data.institutions[].institutionRole for egen profil), ikke portalrollen "employee".
    [Fact]
    public void ProfileContext_role_comes_from_own_institution()
    {
        var json = """
            {"status":{"code":0,"message":"OK"},"data":{"institutionProfile":{"id":9000001,"role":"employee","firstName":"Test","lastName":"Bruger",
             "institution":{"institutionCode":"101001","institutionName":"Testskolen"}},
             "institutions":[{"institutionCode":"202002","institutionProfileId":9000002,"institutionRole":"teacher"},
                             {"institutionCode":"101001","institutionProfileId":9000001,"institutionRole":"preschool-teacher"}]}}
            """;
        Assert.Equal("preschool-teacher", AulaParsers.ParseProfileContext(json).Role);
        Assert.Equal("", AulaParsers.ParseProfileContext(json.Replace("\"institutions\"", "\"x\"")).Role);
    }

    [Fact]
    public void FirstProfile_fallback()
    {
        var p = AulaParsers.ParseFirstProfile(Fixture.Read("profilesByLogin.json"));
        Assert.Equal(new Profile("9000001", "Test Bruger", "", "", "101001", "Testskolen"), p);
    }

    // Webappen logger ud ved status.code 20 eller subCode 23 ("sessionExpired") og sender til login ved code 448.
    [Theory]
    [InlineData("""{"status":{"code":448,"subCode":0,"message":"x"},"data":null}""")]
    [InlineData("""{"status":{"code":20,"message":"x"},"data":null}""")]
    [InlineData("""{"status":{"code":0,"subCode":23,"message":"x"},"data":null}""")]
    public void Session_expiry_codes(string json)
    {
        Assert.Throws<SessionExpiredException>(() => AulaParsers.EnsureOk(json));
        Assert.Throws<SessionExpiredException>(() => AulaParsers.ParseEvents(json));
    }

    // 401 i svaret betyder "ingen tilladelse" i webappen (fx til et lokale), ikke udløb.
    [Fact]
    public void Status_401_means_no_access() =>
        Assert.Throws<ForbiddenException>(() => AulaParsers.ParseEvents("""{"status":{"code":401,"subCode":1,"message":"x"},"data":null}"""));

    [Fact]
    public void Deactivated_profile_is_no_access_with_its_own_text()
    {
        var ex = Assert.Throws<ForbiddenException>(() => AulaParsers.ParseEvents("""{"status":{"code":0,"subCode":33,"message":"x"},"data":null}"""));
        Assert.Equal("Profilen er deaktiveret i Aula", ex.Message);
    }

    [Theory]
    [InlineData(451, "Du har ikke fået adgang til Aula endnu")]
    [InlineData(452, "Din Aula-bruger er deaktiveret")]
    public void Account_access_codes(int code, string text)
    {
        var ex = Assert.Throws<AccessDeniedException>(() => AulaParsers.EnsureOk($$"""{"status":{"code":{{code}},"message":"x"},"data":null}"""));
        Assert.Equal(text, ex.Message);
    }

    // 403 betyder "ingen adgang" (fx ikke medlem af gruppen), ikke udløbet session.
    [Fact]
    public void Status_403_means_no_access_not_expired()
    {
        Assert.Throws<ForbiddenException>(() => AulaParsers.EnsureOk(Fixture.Read("forbidden.json")));
        Assert.Throws<ForbiddenException>(() => AulaParsers.ParseEvents(Fixture.Read("forbidden.json")));
    }

    [Fact]
    public void Error_status_throws() =>
        Assert.Throws<AulaException>(() => AulaParsers.ParseEvents(Fixture.Read("errorStatus.json")));

    [Theory]
    [InlineData("")]
    [InlineData("<html>Vedligeholdelse</html>")]
    [InlineData("{\"data\":[]}")]
    [InlineData("[]")]
    public void Invalid_responses_throw_AulaException(string json)
    {
        var ex = Assert.ThrowsAny<AulaException>(() => AulaParsers.ParseEvents(json));
        Assert.IsNotType<SessionExpiredException>(ex);
    }

    [Fact]
    public void EmployeeSearch_keeps_only_employees_with_valid_ids()
    {
        var list = AulaParsers.ParseEmployeeSearch(Fixture.Read("searchEmployees.json"));
        Assert.Equal(["1001", "2001", "5"], list.Select(e => e.Id));
        Assert.Equal(new Employee("1001", "Anna Eksempel", "AE", "teacher"), list[0]);
    }

    [Fact]
    public void Search_with_null_data_is_empty()
    {
        Assert.Empty(AulaParsers.ParseEmployeeSearch(Fixture.Read("searchNullData.json")));
        Assert.Empty(AulaParsers.ParseGroupSearch(Fixture.Read("searchNullData.json")));
    }

    [Fact]
    public void MasterData()
    {
        var list = AulaParsers.ParseProfileMasterData(Fixture.Read("masterData.json"));
        Assert.Equal(3, list.Count);
        Assert.Equal(new Employee("1001", "Anna Bech Eksempel", "AE", "teacher"), list[0]);
    }

    [Fact]
    public void GroupSearch_keeps_only_main_groups() =>
        Assert.Equal([new NamedItem("88231", "7A")], AulaParsers.ParseGroupSearch(Fixture.Read("searchGroups.json")));

    // findGroups kalder feltet "type" (findProfilesAndGroups: "groupType"); værdien sammenlignes uden store/små bogstaver.
    [Fact]
    public void Group_list_keeps_main_groups_from_type_field() =>
        Assert.Equal([new NamedItem("88231", "7A"), new NamedItem("88233", "0B")], AulaParsers.ParseGroupSearch(Fixture.Read("findGroups.json")));

    // Lokaler: resourceType "location" eller "extra-location" (som i webappens lokalevælger), uanset store og små bogstaver,
    // og ikke inaktive. Uden kategori beholdes ressourcen.
    [Fact]
    public void Resources_keep_active_rooms_only() =>
        Assert.Equal([new NamedItem("412", "53"), new NamedItem("413", "Gymnastiksal"), new NamedItem("416", "Hal 2"), new NamedItem("417", "Musiklokale")],
            AulaParsers.ParseResources(Fixture.Read("resourcesAll.json")));

    [Fact]
    public void Resources_numeric_ids_and_html_decoded_names() =>
        Assert.Equal([new NamedItem("412", "53"), new NamedItem("413", "Gymnastiksal & scene")], AulaParsers.ParseResources(Fixture.Read("resources.json")));

    [Fact]
    public void Events_normal_lesson()
    {
        var r = AulaParsers.ParseEvents(Fixture.Read("events.json"));
        var e = r.Events.Single(x => x.Id == "5001");
        Assert.Equal("DAN", e.Title);
        Assert.Equal(DateTimeOffset.Parse("2026-04-13T06:00:00Z"), e.Start);
        Assert.Equal("53", e.Location);
        Assert.Equal(["7A"], e.Groups);
        Assert.Equal([new Participant("Anna Eksempel", "AE"), new Participant("Carl Testesen", "CT")], e.Teachers);
        Assert.False(e.IsSubstitute);
        Assert.Null(e.SubstituteFor);
    }

    [Fact]
    public void Events_substitute_lesson()
    {
        var e = AulaParsers.ParseEvents(Fixture.Read("events.json")).Events.Single(x => x.Id == "5002");
        Assert.True(e.IsSubstitute);
        Assert.Equal(new Participant("Anna Eksempel", "AE"), e.SubstituteFor);
        Assert.Equal([new Participant("Vera Vikar", "VV")], e.Teachers);
        Assert.Equal("54", e.Location);
    }

    [Fact]
    public void Events_meeting_without_lesson_and_html_title()
    {
        var e = AulaParsers.ParseEvents(Fixture.Read("events.json")).Events.Single(x => x.Id == "5003");
        Assert.Equal("Personalemøde", e.Title);
        Assert.Empty(e.Teachers);
        Assert.Equal("", e.Location);
    }

    [Fact]
    public void Events_with_invalid_dates_are_skipped_and_counted()
    {
        var r = AulaParsers.ParseEvents(Fixture.Read("events.json"));
        Assert.Equal(3, r.Events.Count);
        Assert.Equal(1, r.Skipped);
    }

    [Fact]
    public void ParseEvents_empty_list_is_ok() =>
        Assert.Empty(AulaParsers.ParseEvents(Fixture.Read("eventsEmpty.json")).Events);

    [Fact]
    public void ParseEvents_null_data_throws() =>
        Assert.Throws<AulaException>(() => AulaParsers.ParseEvents(Fixture.Read("eventsNullData.json")));
}
