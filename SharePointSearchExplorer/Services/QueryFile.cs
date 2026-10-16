using System.Xml;
using System.Xml.Linq;

namespace SharePointSearchExplorer.Services;

/// <summary>
/// Reads and writes saved queries as XML, so they can be backed up or shared:
/// <code>
/// &lt;SavedQueries version="1" exported="2026-09-30T12:00:00Z"&gt;
///   &lt;Query&gt;&lt;QueryText&gt;contenttype:document&lt;/QueryText&gt;&lt;/Query&gt;
/// &lt;/SavedQueries&gt;
/// </code>
/// Each query is an element of its own so more settings can be added to it later without breaking old files.
/// </summary>
public static class QueryFile
{
    public const string Extension = ".xml";

    private const string RootName = "SavedQueries";
    private const string QueryName = "Query";
    private const string TextName = "QueryText";

    public static void Save(string path, IEnumerable<string> queries)
    {
        var doc = new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(RootName,
                new XAttribute("version", 1),
                new XAttribute("exported", DateTime.UtcNow.ToString("o")),
                queries.Select(q => new XElement(QueryName, new XElement(TextName, q)))));
        doc.Save(path);
    }

    /// <summary>The non-empty queries in the file, in file order, without duplicates.</summary>
    /// <exception cref="InvalidDataException">The file isn't a saved queries file.</exception>
    public static List<string> Load(string path)
    {
        XDocument doc;
        try
        {
            doc = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException ex)
        {
            throw new InvalidDataException($"The file isn't valid XML: {ex.Message}", ex);
        }
        if (doc.Root?.Name.LocalName != RootName)
            throw new InvalidDataException($"The file isn't a saved queries file (expected a <{RootName}> root element).");

        return doc.Root.Elements()
            .Where(e => e.Name.LocalName == QueryName)
            // <Query><QueryText>…</QueryText></Query>, or simply <Query>…</Query>.
            .Select(e => (e.Elements().FirstOrDefault(c => c.Name.LocalName == TextName) ?? e).Value.Trim())
            .Where(q => q.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>A file name based on the query text, e.g. "contenttype document.xml".</summary>
    public static string SuggestedName(string query)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var name = new string(query.Select(c => invalid.Contains(c) || char.IsControl(c) ? ' ' : c).ToArray());
        name = string.Join(' ', name.Split(' ', StringSplitOptions.RemoveEmptyEntries));
        if (name.Length > 60) name = name[..60].TrimEnd();
        return name.Length == 0 ? "Query" : name;
    }
}
