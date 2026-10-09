using System.Runtime.CompilerServices;
using System.Web;

using AIStudio.Tools.ToolCallingSystem.ToolCallingImplementations.SharePoint;

namespace AIStudio.Tests.Tools.ToolCalling;

/// <summary>
/// Checks the parts of the SharePoint search which decide what is asked and what the model gets.
/// </summary>
/// <remarks>
/// The model supplies words, a site, and a page number, and everything around them comes from
/// here: the query the words end up in, the request to the search API, and the few fields which
/// are read from its answer. A mistake in the first two lets a model change what is searched; one
/// in the last hands it an address it cannot open, or text SharePoint never showed a user. The
/// answers are made up, in the form a SharePoint Server 2019 writes them. The request itself needs
/// a real SharePoint and is left to a manual test.
/// </remarks>
[TestFixture]
public sealed class SharePointSearchToolTests
{
    private static readonly Uri SITE = new("https://intranet.example.org/");

    [Test]
    public void PlainWordsAreSearchedAsTheyAre()
    {
        Assert.That(QueryText("  travel   expenses "), Is.EqualTo("travel expenses"));
    }

    [TestCase("author:\"Jane Doe\" budget", "author Jane Doe budget")]
    [TestCase("contentclass=urn:content-class:SPSPeople", "contentclass urn content-class SPSPeople")]
    [TestCase("size>1000 write<2020-01-01", "size 1000 write 2020-01-01")]
    [TestCase("(budget OR plan) {User.Name}", "budget OR plan User.Name")]
    public void TheWordsCannotBecomeAnotherKindOfQuery(string query, string expected)
    {
        //
        // In SharePoint's query language, a colon or a comparison after a word asks for a property
        // instead of text. A model could search by author that way, or ask for people instead of
        // documents, which the organization never switched on:
        //
        Assert.That(QueryText(query), Is.EqualTo(expected), "Without the characters of the query language, every part is a word to find.");
    }

    [TestCase(":")]
    [TestCase("\"\" ()")]
    public void AQueryOfNothingButSyntaxIsRefused(string query)
    {
        Assert.That(SharePointSearchTool.TryBuildQueryText(query, null, out var queryText), Is.False, "An empty query would match everything the user may read.");
        Assert.That(queryText, Is.Null);
    }

    [Test]
    public void ASiteRestrictsTheSearchToItsPath()
    {
        Assert.That(QueryText("budget OR plan", new Uri("https://teams.example.org/sites/Standards")), Is.EqualTo("(budget OR plan) path:\"https://teams.example.org/sites/Standards\""), "The words are grouped, so an OR among them cannot reach past the restriction.");
    }

    [Test]
    public void ASiteCannotLeaveItsRestriction()
    {
        var site = new Uri("https://teams.example.org/sites/a\" OR author:\"jane");

        Assert.That(QueryText("budget", site), Is.EqualTo("(budget) path:\"https://teams.example.org/sites/a OR author:jane\""), "The site comes from the model as well. Without its quotes, nothing in it can end the restriction, so the rest stays part of a path which matches nothing.");
    }

    [Test]
    public void ASiteKeepsTheSpacesOfItsPath()
    {
        var site = new Uri("https://teams.example.org/sites/Standards/Shared%20Documents");

        Assert.That(QueryText("visitors", site), Is.EqualTo("(visitors) path:\"https://teams.example.org/sites/Standards/Shared Documents\""), "SharePoint keeps the path of a hit with its spaces. Asked for %20, it would find nothing in a library such as Shared Documents.");
    }

    [Test]
    public void TheSearchGoesToTheSearchApiOfTheSite()
    {
        var searchUrl = SharePointSearchTool.BuildSearchUrl(new Uri("https://intranet.example.org/sites/portal/"), "travel expenses", 1);
        var parameters = HttpUtility.ParseQueryString(searchUrl.Query);

        Assert.That(searchUrl.GetLeftPart(UriPartial.Path), Is.EqualTo("https://intranet.example.org/sites/portal/_api/search/query"), "The API lies below the configured site.");
        Assert.That(parameters["querytext"], Is.EqualTo("'travel expenses'"), "The API takes its texts in single quotes.");
        Assert.That(parameters["rowlimit"], Is.EqualTo("10"));
        Assert.That(parameters["startrow"], Is.EqualTo("0"));
        Assert.That(parameters["selectproperties"], Does.Contain("HitHighlightedSummary"), "Only the properties the tool hands on are asked for.");
    }

    [TestCase(1, "0")]
    [TestCase(2, "10")]
    [TestCase(5, "40")]
    public void APageStartsWhereTheOneBeforeEnded(int page, string expectedStartRow)
    {
        var parameters = HttpUtility.ParseQueryString(SharePointSearchTool.BuildSearchUrl(SITE, "travel", page).Query);

        Assert.That(parameters["startrow"], Is.EqualTo(expectedStartRow));
    }

    [Test]
    public void AnApostropheCannotEndTheQuotedText()
    {
        var parameters = HttpUtility.ParseQueryString(SharePointSearchTool.BuildSearchUrl(SITE, "the manager's guide", 1).Query);

        Assert.That(parameters["querytext"], Is.EqualTo("'the manager''s guide'"), "Written twice, a single quote is part of the text. Written once, it would end the text and leave the rest to the API.");
    }

    [Test]
    public void TheQueryStaysInsideItsParameter()
    {
        // Characters which mean something in a URL must neither end a parameter, start a fragment, nor change the path:
        const string QUERY = "R&D #1 ../../admin?x=1&rowlimit=500";
        var searchUrl = SharePointSearchTool.BuildSearchUrl(SITE, QUERY, 1);
        var parameters = HttpUtility.ParseQueryString(searchUrl.Query);

        Assert.That(searchUrl.AbsolutePath, Is.EqualTo("/_api/search/query"));
        Assert.That(searchUrl.Fragment, Is.Empty);
        Assert.That(parameters.AllKeys, Is.EquivalentTo(new[] { "querytext", "selectproperties", "rowlimit", "startrow", "trimduplicates" }), "No word of the query may become a parameter of its own.");
        Assert.That(parameters["querytext"], Is.EqualTo($"'{QUERY}'"));
        Assert.That(parameters["rowlimit"], Is.EqualTo("10"));
    }

    [Test]
    public void TheHitsAreReadFromTheAnswer()
    {
        var response = Parse("sharepoint-search.xml");

        Assert.Multiple(() =>
        {
            Assert.That(response.TotalRows, Is.EqualTo(37));
            Assert.That(response.Hits.Select(hit => hit.Title), Is.EqualTo(new[]
            {
                "New travel expense rules from January",
                "Travel Expense Guideline",
                "Finance Team Site",
                "Visitor Registration.pdf",
                "Travel budget 2026",
                "Canteen menu",
            }), "The hits keep the order SharePoint ranked them in. A hit without a title is named after its file, and line breaks in a title go.");

            var page = response.Hits[0];
            Assert.That(page.Url.AbsoluteUri, Is.EqualTo("https://intranet.example.org/News/Pages/travel-expense-rules.aspx"));
            Assert.That(page.Excerpt, Is.Empty, "SharePoint has no excerpt for every hit.");
            Assert.That(page.Modified, Is.EqualTo("2026-01-12"));
            Assert.That(page.FileType, Is.EqualTo("html"));
            Assert.That(page.Site, Is.EqualTo("https://intranet.example.org/News"));
        });
    }

    [Test]
    public void AHitWhichIsNoWebAddressIsLeftOut()
    {
        var response = Parse("sharepoint-search.xml");

        Assert.That(response.Hits.Select(hit => hit.Url.Scheme), Is.All.EqualTo("https"), "SharePoint also indexes file shares. Neither the model nor a link in the answer could open such a hit.");
        Assert.That(response.Hits, Has.Count.EqualTo(6));
    }

    [Test]
    public void TheMarksOfSharePointLeaveTheExcerpt()
    {
        var response = Parse("sharepoint-search.xml");

        Assert.Multiple(() =>
        {
            Assert.That(response.Hits[1].Excerpt, Is.EqualTo("Receipts for travel expenses are handed in within four weeks … a rail journey is preferred over a flight for every travel below 600 km"), "The marks around the words found go, and a place where text was left out becomes an ellipsis.");
            Assert.That(response.Hits[2].Excerpt, Is.EqualTo("Template for travel requests 07.08.2025 09:37 Finance … Expenses_Overview_2025 09.05.2025 14:01 Finance"), "The line breaks of a list view go as well.");
        });
    }

    [TestCase(null, "")]
    [TestCase("   ", "")]
    [TestCase("<ddd/> plain <C0>words</C0> <ddd/>", "plain words")]
    [TestCase("a <b>bold</b> claim", "a <b>bold</b> claim")]
    public void OnlyTheMarksOfSharePointAreRemoved(string? summary, string expected)
    {
        Assert.That(SharePointSearchResponse.CleanExcerpt(summary), Is.EqualTo(expected), "Anything else is text of the document and stays for the prompt injection filter to judge.");
    }

    [Test]
    public void AnAnswerWithoutHitsIsStillAnAnswer()
    {
        var response = Parse("sharepoint-search-empty.xml");

        Assert.That(response.Hits, Is.Empty);
        Assert.That(response.TotalRows, Is.Zero);
        Assert.That(SharePointSearchTool.HasMore(response, 1), Is.False);
    }

    [TestCase("")]
    [TestCase("<html><body>Sign in</body></html>")]
    [TestCase("{\"d\":{\"query\":{}}}")]
    [TestCase("<?xml version=\"1.0\"?><feed xmlns=\"http://www.w3.org/2005/Atom\"/>")]
    [TestCase("<?xml version=\"1.0\"?><!DOCTYPE d [<!ENTITY e \"x\">]><d:query xmlns:d=\"http://schemas.microsoft.com/ado/2007/08/dataservices\"><d:RelevantResults/></d:query>")]
    public void SomethingElseIsNoAnswerOfTheSearch(string body)
    {
        Assert.That(SharePointSearchResponse.TryParse(body, out var response), Is.False, "A sign-in page, an error, or another document must not reach the model as a search without hits.");
        Assert.That(response, Is.Null);
    }

    [Test]
    public void TheModelGetsAnAddressItCanPassOn()
    {
        var response = Parse("sharepoint-search.xml");
        var content = SharePointSearchTool.BuildModelContent(response.Hits, 1, true);
        var document = content["results"]![1]!;

        Assert.Multiple(() =>
        {
            Assert.That(document["url"]!.GetValue<string>(), Is.EqualTo("https://docs.example.org/Documents/Travel%20Expense%20Guideline.pdf"), "A space would end the address where the model reads it, and Read Web Page compares the escaped form.");
            Assert.That(document["file_type"]!.GetValue<string>(), Is.EqualTo("pdf"));
            Assert.That(content["results"]![4]!["file_type"]!.GetValue<string>(), Is.EqualTo("xlsx"), "The instructions name the file types in lower case.");
            Assert.That(content["page"]!.GetValue<int>(), Is.EqualTo(1));
            Assert.That(content["has_more"]!.GetValue<bool>(), Is.True);
        });
    }

    [Test]
    public void AnEmptyFieldIsLeftOut()
    {
        var response = Parse("sharepoint-search.xml");
        var content = SharePointSearchTool.BuildModelContent(response.Hits, 1, false);

        Assert.Multiple(() =>
        {
            Assert.That(content["results"]![0]!.AsObject().ContainsKey("excerpt"), Is.False);
            Assert.That(content["results"]![4]!.AsObject().ContainsKey("modified"), Is.False);
            Assert.That(content["results"]![4]!.AsObject().ContainsKey("site"), Is.True);
        });
    }

    [TestCase(1, true)]
    [TestCase(3, true)]
    [TestCase(4, false)]
    [TestCase(5, false)]
    public void MoreIsOfferedWhileSharePointHasMoreAndTheModelMayPage(int page, bool expected)
    {
        // The answer names 37 hits, so the fourth page is the last one with any:
        Assert.That(SharePointSearchTool.HasMore(Parse("sharepoint-search.xml"), page), Is.EqualTo(expected));
    }

    [Test]
    public void TheLastAllowedPageOffersNoMore()
    {
        var response = Parse("sharepoint-search.xml") with { TotalRows = 10_000 };

        Assert.That(SharePointSearchTool.HasMore(response, SharePointSearchTool.MAX_PAGE - 1), Is.True);
        Assert.That(SharePointSearchTool.HasMore(response, SharePointSearchTool.MAX_PAGE), Is.False, "Offering a page the tool then refuses would cost the model a call.");
    }

    private static string QueryText(string query, Uri? site = null) => SharePointSearchTool.TryBuildQueryText(query, site, out var queryText)
        ? queryText
        : throw new AssertionException($"'{query}' should hold a word to search for.");

    private static SharePointSearchResponse Parse(string fixtureName) => SharePointSearchResponse.TryParse(ReadFixture(fixtureName), out var response)
        ? response
        : throw new AssertionException($"'{fixtureName}' should be an answer of the search API.");

    /// <summary>
    /// Reads a file from the fixtures next to this test.
    /// </summary>
    /// <remarks>
    /// Read from the source tree, so the fixture needs no entry in the project file.
    /// </remarks>
    private static string ReadFixture(string fileName, [CallerFilePath] string sourceFilePath = "") =>
        File.ReadAllText(Path.Combine(Path.GetDirectoryName(sourceFilePath)!, "Fixtures", fileName));
}
