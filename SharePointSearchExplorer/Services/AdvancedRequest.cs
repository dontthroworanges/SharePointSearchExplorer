using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Services;

/// <summary>The two REST endpoints the Advanced Query Rules page can call.</summary>
public enum AdvancedMode { Query, Suggest }

public enum ParamKind { Text, Number, Flag }

/// <summary>An entry in a field's suggestion list: what is shown, and the value it puts in the field.</summary>
public sealed record ParamSuggestion(string Label, string Value)
{
    public override string ToString() => Label;
}

/// <summary>One REST parameter that has a field on the Advanced Query Rules page.</summary>
/// <param name="Key">The query-string name, in the casing used when the URL is built.</param>
/// <param name="Quoted">Whether the value is sent in single quotes (string parameters are; numbers and flags are not).</param>
public sealed record ParamDef(
    string Key, string Label, ParamKind Kind = ParamKind.Text, bool Quoted = true, string Tip = "",
    IReadOnlyList<ParamSuggestion>? Suggestions = null);

/// <summary>
/// The parameters of an advanced query or suggest request. Converts both ways between the field values and
/// the request URL, so that the fields and the editable URL on the Advanced Query Rules page stay in step.
/// </summary>
public sealed class AdvancedRequest
{
    public const string QueryTextKey = "querytext";
    public const string SelectPropertiesKey = "selectproperties";

    private readonly Dictionary<string, string> _values = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Query-string segments with no field (unknown or half-typed parameters), kept exactly as typed.</summary>
    private readonly List<string> _extras = new();

    public AdvancedRequest(AdvancedMode mode) => Mode = mode;

    public AdvancedMode Mode { get; }

    public IReadOnlyList<ParamDef> Params => ParamsFor(Mode);

    public static IReadOnlyList<ParamDef> ParamsFor(AdvancedMode mode) => mode == AdvancedMode.Query ? QueryParams : SuggestParams;

    /// <summary>A request with the values the page starts from (and returns to on Reset).</summary>
    public static AdvancedRequest CreateDefault(AdvancedMode mode)
    {
        var request = new AdvancedRequest(mode);
        if (mode == AdvancedMode.Query) request.Set("clienttype", "ContentSearchRegular");
        return request;
    }

    public string Get(string key) => _values.TryGetValue(key, out var value) ? value : "";

    public void Set(string key, string value)
    {
        if (string.IsNullOrEmpty(value)) _values.Remove(key);
        else _values[key] = value;
    }

    #region Saved requests

    private const string EndpointPrefix = "_api/search/";

    /// <summary>
    /// A request as it is kept in the query history: the URL without the site ("_api/search/query?…"), so it
    /// can be told apart from a plain query text and run against whichever profile is current.
    /// </summary>
    public static string ToSaved(string url)
    {
        int start = url.IndexOf(EndpointPrefix, StringComparison.OrdinalIgnoreCase);
        return start < 0 ? url : url[start..];
    }

    public static bool IsSaved(string historyEntry) => historyEntry.StartsWith(EndpointPrefix, StringComparison.OrdinalIgnoreCase);

    /// <summary>The URL of a saved request on the given site.</summary>
    public static string FromSaved(string saved, string siteUrl) =>
        string.IsNullOrWhiteSpace(siteUrl) ? saved : siteUrl.TrimEnd('/') + "/" + saved;

    #endregion

    #region URL

    // Characters left readable in the URL. They are safe inside a query-string value here (HttpClient encodes
    // the ones that need it when sending); &, #, + and % stay encoded so the URL can be split up again.
    private static readonly (string Encoded, string Plain)[] s_readable =
    [
        ("%20", " "), ("%22", "\""), ("%27", "'"), ("%28", "("), ("%29", ")"), ("%2A", "*"),
        ("%2C", ","), ("%2F", "/"), ("%3A", ":"), ("%3D", "="), ("%7C", "|"),
    ];

    /// <summary>The request URL: the query text first, then every other parameter that has a value.</summary>
    public string Build(string siteUrl)
    {
        // An empty query matches everything, as on the Search Results page.
        var queryText = Get(QueryTextKey) is { Length: > 0 } text ? text : Mode == AdvancedMode.Query ? "*" : "";
        var parts = new List<string> { Format(Params.First(p => p.Key == QueryTextKey), queryText) };
        foreach (var def in Params)
        {
            if (def.Key != QueryTextKey && Get(def.Key) is { Length: > 0 } value) parts.Add(Format(def, value));
        }
        parts.AddRange(_extras);

        var endpoint = Mode == AdvancedMode.Query ? "_api/search/query" : "_api/search/suggest";
        var site = string.IsNullOrWhiteSpace(siteUrl) ? "" : siteUrl.TrimEnd('/') + "/";
        return $"{site}{endpoint}?{string.Join("&", parts)}";

        static string Format(ParamDef def, string value) => def.Quoted
            ? $"{def.Key}='{Escape(value.Replace("'", "''"))}'"
            : $"{def.Key}={Escape(value)}";
    }

    /// <summary>
    /// Reads a request back from a (possibly hand-edited) URL. The endpoint decides the mode; the site part
    /// is ignored, since requests always go to the current profile's site.
    /// </summary>
    public static AdvancedRequest Parse(string url)
    {
        int q = url.IndexOf('?');
        var path = q < 0 ? url : url[..q];
        var request = new AdvancedRequest(path.Contains("/suggest", StringComparison.OrdinalIgnoreCase) ? AdvancedMode.Suggest : AdvancedMode.Query);
        if (q < 0) return request;

        foreach (var segment in url[(q + 1)..].Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            int eq = segment.IndexOf('=');
            var def = eq < 0 ? null : request.Params.FirstOrDefault(p => p.Key.Equals(segment[..eq].Trim(), StringComparison.OrdinalIgnoreCase));
            if (def == null)
            {
                request._extras.Add(segment);
                continue;
            }

            var value = Uri.UnescapeDataString(segment[(eq + 1)..]);
            if (value.Length >= 2 && value[0] == '\'' && value[^1] == '\'') value = value[1..^1].Replace("''", "'");
            else if (value.StartsWith('\'')) value = value[1..]; // closing quote not typed yet
            request.Set(def.Key, value);
        }
        return request;
    }

    private static string Escape(string value)
    {
        var escaped = Uri.EscapeDataString(value);
        foreach (var (encoded, plain) in s_readable) escaped = escaped.Replace(encoded, plain);
        return escaped;
    }

    #endregion

    /// <summary>The parts of a query request the rest of the app uses (property lookups, the Debug tab).</summary>
    public SearchQuery ToSearchQuery() => new()
    {
        QueryText = Get(QueryTextKey),
        StartRow = int.TryParse(Get("startrow"), out int start) ? start : 0,
        RowLimit = int.TryParse(Get("rowlimit"), out int limit) ? limit : 10,
        SelectProperties = Get(SelectPropertiesKey).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
        Refiners = NullIfEmpty(Get("refiners")),
        RefinementFilters = NullIfEmpty(Get("refinementfilters")),
        HiddenConstraints = NullIfEmpty(Get("hiddenconstraints")),
        SortList = NullIfEmpty(Get("sortlist")),
        QueryTemplate = NullIfEmpty(Get("querytemplate")),
        SourceId = NullIfEmpty(Get("sourceid")),
        TrimDuplicates = !Get("trimduplicates").Equals("false", StringComparison.OrdinalIgnoreCase),
    };

    private static string? NullIfEmpty(string value) => value.Length == 0 ? null : value;

    #region Rank detail

    private static readonly string[] s_rankDetailProperties = ["RankDetail", "Title", "Path", "Language", "WorkId"];

    /// <summary>"Include rank detail" is not a parameter of its own: it is on while RankDetail is a select property.</summary>
    public static bool HasRankDetail(string selectProperties) =>
        selectProperties.Split(',', StringSplitOptions.TrimEntries).Contains("RankDetail", StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Adds RankDetail (and the properties the original tool's rank detail view needs) to the select
    /// properties, or removes RankDetail again.
    /// </summary>
    public static string SetRankDetail(string selectProperties, bool include)
    {
        if (HasRankDetail(selectProperties) == include) return selectProperties;
        var properties = selectProperties.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
        if (include)
        {
            foreach (var property in s_rankDetailProperties)
            {
                if (!properties.Contains(property, StringComparer.OrdinalIgnoreCase)) properties.Add(property);
            }
        }
        else
        {
            properties.RemoveAll(p => p.Equals("RankDetail", StringComparison.OrdinalIgnoreCase));
        }
        return string.Join(",", properties);
    }

    #endregion

    #region Parameter definitions

    private static ParamSuggestion[] Example(string value) => [new($"Example: {value}", value)];

    // Languages SharePoint search knows, as in the original tool: "Name [code | LCID]"; the LCID is the value.
    private static readonly ParamSuggestion[] s_cultures = new (string Name, string Code, int Lcid)[]
    {
        ("Afrikaans", "af", 54), ("Albanian", "sq", 28), ("Arabic", "ar", 1), ("Armenian", "hy", 43), ("Azerbaijani", "az", 44),
        ("Bangla", "bn", 69), ("Basque", "eu", 45), ("Belarusian", "be", 35), ("Bosnian", "bs", 30746), ("Breton", "br", 126),
        ("Bulgarian", "bg", 2), ("Catalan", "ca", 3), ("Chinese Simplified", "zh-cn", 2052), ("Chinese Traditional", "zh-tw", 1028),
        ("Croatian", "hr", 26), ("Czech", "cs", 5), ("Danish", "da", 6), ("Dutch", "nl", 19), ("English", "en", 9),
        ("English (United States)", "en-US", 1033), ("Estonian", "et", 37), ("Faeroese", "fo", 56), ("Farsi", "fa", 41),
        ("Filipino", "fil", 100), ("Finnish", "fi", 11), ("French", "fr", 12), ("Frisian", "fy", 98), ("Galician", "gl", 86),
        ("Georgian", "ka", 55), ("German", "de", 7), ("Greek", "el", 8), ("Gujarati", "gu", 71), ("Hausa", "ha", 104),
        ("Hebrew", "he", 13), ("Hindi", "hi", 57), ("Hungarian", "hu", 14), ("Icelandic", "is", 15), ("Indonesian", "id", 33),
        ("Irish Gaelic", "ga", 60), ("Italian", "it", 16), ("Japanese", "ja", 17), ("Kannada", "kn", 75), ("Kazakh", "kk", 63),
        ("Kirghiz", "ky", 64), ("Korean", "ko", 18), ("Kurdish", "ku", 146), ("Latvian", "lv", 38), ("Lithuanian", "lt", 39),
        ("Luxembourgish", "lb", 110), ("Macedonian", "mk", 47), ("Malay", "ms", 62), ("Malayalam", "ml", 76), ("Maltese", "mt", 58),
        ("Maori", "mi", 129), ("Marathi", "mr", 78), ("Mongolian", "mn", 80), ("Norwegian Bokmål", "nb", 31764),
        ("Norwegian Nynorsk", "nn", 30740), ("Pashto", "ps", 99), ("Polish", "pl", 21), ("Portuguese", "pt", 22), ("Punjabi", "pa", 70),
        ("Romanian", "ro", 24), ("Romansh", "rm", 23), ("Russian", "ru", 25), ("Sami Northern", "se", 59), ("Serbian", "sr", 31770),
        ("Slovak", "sk", 27), ("Slovenian", "sl", 36), ("Spanish", "es", 10), ("Swahili", "sw", 65), ("Swedish", "sv", 29),
        ("Tamil", "ta", 73), ("Telugu", "te", 74), ("Thai", "th", 30), ("Turkish", "tr", 31), ("Ukrainian", "uk", 34),
        ("Urdu", "ur", 32), ("Uzbek", "uz", 67), ("Vietnamese", "vi", 42), ("Welsh", "cy", 82), ("Zulu", "zu", 53),
    }.Select(c => new ParamSuggestion($"{c.Name} [{c.Code} | {c.Lcid}]", c.Lcid.ToString())).ToArray();

    // The built-in result sources.
    private static readonly ParamSuggestion[] s_resultSources = new (string Name, string Id)[]
    {
        ("Local SharePoint Results", "8413cd39-2156-4e00-b54d-11efd9abdb89"),
        ("Local People Results", "b09a7990-05ea-4af9-81ef-edfab16c4e31"),
        ("Local Reports And Data Results", "203fba36-2763-4060-9931-911ac8c0583b"),
        ("Local Video Results", "78b793ce-7956-4669-aa3b-451fc5defebf"),
        ("Documents", "e7ec8cee-ded8-43c9-beb5-436b54b31e84"),
        ("Items matching a content type", "5dc9f503-801e-4ced-8a2c-5d1237132419"),
        ("Items matching a tag", "e1327b9c-2b8c-4b23-99c9-3730cb29c3f7"),
        ("Items related to current user", "48fec42e-4a92-48ce-8363-c2703a40e67d"),
        ("Items with same keyword as this item", "5c069288-1d17-454a-8ac6-9c642a065f48"),
        ("Pages", "5e34578e-4d08-4edc-8bf3-002acf3cdbcc"),
        ("Pictures", "38403c8c-3975-41a8-826e-717f2d41568a"),
        ("Popular", "97c71db1-58ce-4891-8b64-585bc2326c12"),
        ("Recently changed items", "ba63bbae-fa9c-42c0-b027-9a878f16557c"),
        ("Recommended Items", "ec675252-14fa-4fbe-84dd-8d098ed74181"),
        ("Wiki", "9479bf85-e257-4318-b5a8-81a180f5faa1"),
    }.Select(s => new ParamSuggestion($"{s.Name} [{s.Id}]", s.Id)).ToArray();

    private static readonly ParamSuggestion[] s_clientTypes = new[]
    {
        "Alerts", "AllResultsQuery", "CatalogItemReuseQuery", "ContentSearchHigh", "ContentSearchLow", "ContentSearchRegular",
        "CSOM", "Custom", "DiscoveryDownloadManager", "DiscoverySearch", "DocsSharedWithMe", "Monitoring", "MySiteSecurityTrimmer",
        "MyTaskSync", "ObjectModel", "ObjectModelBackwardsCompatible", "PeopleResultsQuery", "ReportsAndDataResultsQuery", "RSS",
        "SearchWebPartConfiguration", "SEOSiteMapQuery", "SiteResultsQuery_All", "SiteResultsQuery_Docs", "SiteResultsQuery_Sites",
        "UI", "Unknown", "VideoResultsQuery", "WebService",
    }.Select(t => new ParamSuggestion(t, t)).ToArray();

    /// <summary>The Query tab of the original tool, in its order. Flags are shown together in one drop-down.</summary>
    public static readonly IReadOnlyList<ParamDef> QueryParams =
    [
        new(QueryTextKey, "Query text", Tip: "The search query, in KQL (or FQL when Enable FQL is on)."),
        new("startrow", "Start row", ParamKind.Number, Quoted: false, Tip: "A zero-based index of the first search result to be returned."),
        new("rowlimit", "Row limit", ParamKind.Number, Quoted: false, Tip: "The maximum number of search results to return, starting at the index specified in Start row."),
        new("rowsperpage", "Rows per page", ParamKind.Number, Quoted: false, Tip: "The number of results to return per page."),

        new("enablestemming", "Enable stemming", ParamKind.Flag, Quoted: false,
            Tip: "Whether inflectional forms of the query's tokens are used to find results. Example: with stemming, \"car\" also returns items containing \"cars\"."),
        new("enablequeryrules", "Enable query rules", ParamKind.Flag, Quoted: false, Tip: "Whether query rules are applied to the query."),
        new("enableinterleaving", "Enable interleaving", ParamKind.Flag, Quoted: false, Tip: "Whether query rule result blocks can be interleaved in the response."),
        new("enablenicknames", "Enable nicknames", ParamKind.Flag, Quoted: false, Tip: "Whether nicknames of the query's tokens are considered as well as the exact tokens."),
        new("processbestbets", "Process best bets", ParamKind.Flag, Quoted: false, Tip: "Whether promoted results are returned, if available, as a result set."),
        new("processpersonalfavorites", "Process personal favorites", ParamKind.Flag, Quoted: false, Tip: "Whether personal favorites are returned with the results."),
        new("trimduplicates", "Trim duplicates", ParamKind.Flag, Quoted: false, Tip: "Whether duplicates are removed before sorting, selecting and returning result items."),
        new("enablefql", "Enable FQL", ParamKind.Flag, Quoted: false, Tip: "Whether the query text is a FAST Query Language (FQL) query."),
        new("enablephonetic", "Enable phonetic", ParamKind.Flag, Quoted: false, Tip: "Whether the phonetic forms of tokens are used to find results."),
        new("bypassresulttypes", "Bypass result types", ParamKind.Flag, Quoted: false, Tip: "Whether the search result item type is returned for the query results."),
        new("generateblockranklog", "Generate block rank log", ParamKind.Flag, Quoted: false, Tip: "Whether a block rank log is generated."),

        new(SelectPropertiesKey, "Select properties",
            Tip: "The managed properties to return for each result. A property must be Retrievable in the search schema to be returned.",
            Suggestions: Example("Path,Url,Title,Author,Write,ServerRedirectedEmbedURL,WorkId")),
        new("hiddenconstraints", "Hidden constraints", Tip: "Additional query terms that are appended to the query.",
            Suggestions: Example("-ContentClass=urn:content-class:SPSPeople")),
        new("refiners", "Refiners", Tip: "The refiners to return, as a comma-separated list.",
            Suggestions: Example("companies,contentclass,ContentType(filter=7/0/*),FileExtension(filter=7/0/*),FileType(filter=6/0/*)")),
        new("refinementfilters", "Refinement filters",
            Tip: "Refinement tokens for drilling down into the results. Tokens are returned in the RefinementResults table of a query.",
            Suggestions: Example("filetype:equals(\"pptx\")")),
        new("sortlist", "Sort list", Tip: "The properties to sort the search results by.",
            Suggestions: Example("Rank:descending,modifiedby:ascending")),
        new("culture", "Culture", Quoted: false,
            Tip: "The language culture of the query, as an LCID. Example: 1033 for en-US.", Suggestions: s_cultures),
        new("sourceid", "Result source ID",
            Tip: "The result source to run the query against: one of the built-in sources, or the ID of a custom one.", Suggestions: s_resultSources),
        new("rankingmodelid", "Ranking model ID",
            Tip: "The GUID of the ranking model to use for the query. 9399df62-f089-4033-bdc5-a7ea22936e8e is the default model of SharePoint 2010.",
            Suggestions: Example("9399df62-f089-4033-bdc5-a7ea22936e8e")),
        new("trimduplicatesincludeid", "Trim duplicates include ID", ParamKind.Number, Quoted: false, Tip: "ID of the document to include duplicates for."),
        new("hithighlightedproperties", "Hit highlighted properties", Tip: "The properties to include in HitHighlightedProperties for each result."),
        new("collapsespecification", "Collapse specification",
            Tip: "Managed properties used to collapse individual results: results are collapsed into one (or a given number) when they match on these properties."),
        new("properties", "Appended query properties", Tip: "Extra query properties, comma separated.",
            Suggestions: Example("EnableMultiGeoSearch:true,ContentSetting:3")),
        new("personalizationdata", "Personalization data", Tip: "The user profile GUID of a person. Used for social distance in people search."),
        new("querytemplate", "Query template", Tip: "A query template to apply to the query text, e.g. with {searchTerms}."),
        new("resultsurl", "Results URL", Tip: "The URL of the page where the search results are going to be displayed."),
        new("querytag", "Query tag", Tip: "Custom tags that identify the query. Separate multiple tags with a semicolon."),
        new("clienttype", "Query log client type", Tip: "The type of client the query is logged as.", Suggestions: s_clientTypes),
    ];

    /// <summary>The Suggest tab of the original tool.</summary>
    public static readonly IReadOnlyList<ParamDef> SuggestParams =
    [
        new(QueryTextKey, "Query text", Tip: "The text to get suggestions for."),
        new("fprequerysuggestions", "Pre-query suggestions", ParamKind.Flag, Quoted: false, Tip: "Whether to retrieve pre-query suggestions (suggestions for a partly typed query)."),
        new("showpeoplenamesuggestions", "Show people name suggestions", ParamKind.Flag, Quoted: false, Tip: "Whether people names are included in the suggestions."),
        new("fhithighlighting", "Hit highlighting", ParamKind.Flag, Quoted: false, Tip: "Whether the matching part of each suggestion is highlighted."),
        new("fcapitalizefirstletters", "Capitalize first letters", ParamKind.Flag, Quoted: false, Tip: "Whether the first letter of each word in a suggestion is capitalized."),
        new("inumberofquerysuggestions", "Number of query suggestions", ParamKind.Number, Quoted: false, Tip: "The number of query suggestions to retrieve."),
        new("inumberofresultsuggestions", "Number of result suggestions", ParamKind.Number, Quoted: false, Tip: "The number of personal results to retrieve."),
        new("culture", "Culture", Quoted: false,
            Tip: "The language culture of the suggestion query, as an LCID. Example: 1033 for en-US.", Suggestions: s_cultures),
    ];

    #endregion
}
