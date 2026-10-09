using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace AIStudio.Tools.ToolCallingSystem.ToolCallingImplementations.SharePoint;

/// <summary>
/// One hit of a SharePoint search, cut down to what a model needs to judge and open it.
/// </summary>
/// <param name="Title">The title of the page or document.</param>
/// <param name="Url">Where the hit lies, in the form a request is sent in.</param>
/// <param name="Excerpt">The passage SharePoint found the words in, possibly empty.</param>
/// <param name="Modified">The day of the last change as yyyy-MM-dd, or empty when SharePoint named none.</param>
/// <param name="FileType">The file type as SharePoint names it, such as html or pdf, possibly empty.</param>
/// <param name="Site">The site the hit belongs to, or empty when SharePoint named none.</param>
public sealed record SharePointSearchHit(string Title, Uri Url, string Excerpt, string Modified, string FileType, string Site);

/// <summary>
/// The answer of SharePoint's search API, read from its XML.
/// </summary>
/// <remarks>
/// SharePoint answers /_api/search/query with one table: a row per hit, and in each row one cell
/// per property, more than fifty of them. A model needs six. Reading them here, instead of handing
/// the XML on, keeps a search of ten hits at a few hundred tokens instead of tens of thousands,
/// and it keeps everything else SharePoint knows about a document away from the model.
/// </remarks>
/// <param name="Hits">The hits in the order SharePoint ranked them.</param>
/// <param name="TotalRows">How many hits SharePoint has for the query, all pages together.</param>
public sealed partial record SharePointSearchResponse(IReadOnlyList<SharePointSearchHit> Hits, int TotalRows)
{
    private static readonly XNamespace DATA = "http://schemas.microsoft.com/ado/2007/08/dataservices";

    // SharePoint marks the words it found as <c0>word</c0>, with one number per search term:
    [GeneratedRegex(@"</?c\d+>", RegexOptions.IgnoreCase)]
    private static partial Regex HitHighlightPattern();

    // ... and the places where it left text out as <ddd/>:
    [GeneratedRegex(@"\s*<ddd\s*/>\s*", RegexOptions.IgnoreCase)]
    private static partial Regex OmissionPattern();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespacePattern();

    /// <summary>
    /// Reads the answer of the search API.
    /// </summary>
    /// <remarks>
    /// A hit without an HTTP or HTTPS address is left out: the model could not open it, and the
    /// user could not follow it. A document type definition is refused, because the answer of the
    /// API never has one and a document which brings one is not that answer.
    /// </remarks>
    /// <param name="xml">The body of the response.</param>
    /// <param name="response">The hits, when the body is an answer of the search API.</param>
    /// <returns>True when the body is an answer of the search API, also one without hits.</returns>
    public static bool TryParse(string xml, [NotNullWhen(true)] out SharePointSearchResponse? response)
    {
        response = null;
        XElement? relevantResults;
        try
        {
            using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            relevantResults = XDocument.Load(reader).Descendants(DATA + "RelevantResults").FirstOrDefault();
        }
        catch (XmlException)
        {
            return false;
        }

        if (relevantResults is null)
            return false;

        var hits = new List<SharePointSearchHit>();
        var rows = relevantResults.Element(DATA + "Table")?.Element(DATA + "Rows")?.Elements(DATA + "element") ?? [];
        foreach (var row in rows)
        {
            var cells = ReadCells(row);
            if (!TryReadAddress(cells.GetValueOrDefault("Path"), out var url))
                continue;

            var title = CollapseWhitespace(cells.GetValueOrDefault("Title"));
            hits.Add(new SharePointSearchHit(
                title.Length > 0 ? title : ReadNameFromAddress(url),
                url,
                CleanExcerpt(cells.GetValueOrDefault("HitHighlightedSummary")),
                ReadDay(cells.GetValueOrDefault("LastModifiedTime")),
                CollapseWhitespace(cells.GetValueOrDefault("FileType")).ToLowerInvariant(),
                TryReadAddress(cells.GetValueOrDefault("SPWebUrl"), out var site) ? site.AbsoluteUri : string.Empty));
        }

        _ = int.TryParse(relevantResults.Element(DATA + "TotalRows")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var totalRows);
        response = new SharePointSearchResponse(hits, Math.Max(totalRows, hits.Count));
        return true;
    }

    /// <summary>
    /// Turns the excerpt of a hit into plain text.
    /// </summary>
    /// <remarks>
    /// The marks around the words found go, the words stay. A place where SharePoint left text out
    /// becomes an ellipsis, so the model does not read two unrelated passages as one sentence.
    /// </remarks>
    internal static string CleanExcerpt(string? summary)
    {
        if (string.IsNullOrWhiteSpace(summary))
            return string.Empty;

        var text = HitHighlightPattern().Replace(summary, string.Empty);
        text = OmissionPattern().Replace(text, " … ");
        return CollapseWhitespace(text).TrimEnd('…', ' ').TrimStart('…', ' ');
    }

    private static Dictionary<string, string> ReadCells(XElement row)
    {
        // Property names differ in case between SharePoint versions, contentclass for instance:
        var cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in row.Element(DATA + "Cells")?.Elements(DATA + "element") ?? [])
        {
            var key = cell.Element(DATA + "Key")?.Value;
            if (!string.IsNullOrWhiteSpace(key))
                cells[key] = cell.Element(DATA + "Value")?.Value ?? string.Empty;
        }

        return cells;
    }

    private static bool TryReadAddress(string? value, [NotNullWhen(true)] out Uri? url) =>
        Uri.TryCreate(value?.Trim(), UriKind.Absolute, out url) && url.Scheme is "http" or "https";

    // For a hit without a title: the file name, or the host for the root of a site.
    private static string ReadNameFromAddress(Uri url)
    {
        var name = Uri.UnescapeDataString(url.Segments[^1].Trim('/'));
        return name.Length > 0 ? name : url.Host;
    }

    private static string ReadDay(string? value) =>
        DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var pointInTime)
            ? pointInTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : string.Empty;

    private static string CollapseWhitespace(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : WhitespacePattern().Replace(value, " ").Trim();
}
