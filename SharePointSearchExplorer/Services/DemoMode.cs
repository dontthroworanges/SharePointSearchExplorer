using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Services;

/// <summary>
/// TEMPORARY: an offline demo of a fictional Contoso tenant, for taking README screenshots. Start the app with
/// <c>--demo</c> (optionally <c>--demo-theme=Dark</c>) or set <c>SPSE_DEMO=1</c>. Settings are seeded fresh in a
/// temp folder on every launch, sign-in is skipped, and every SharePoint request is answered by
/// <see cref="DemoHandler"/> from the sample dataset below, so the real request and response code still runs.
/// </summary>
public static class DemoMode
{
    public static bool IsEnabled { get; } =
        Environment.GetCommandLineArgs().Any(a => a.Equals("--demo", StringComparison.OrdinalIgnoreCase))
        || Environment.GetEnvironmentVariable("SPSE_DEMO") == "1";

    public static string DataFolder { get; } = Path.Combine(Path.GetTempPath(), "SharePointSearchExplorer.Demo");

    private static string Theme =>
        Environment.GetCommandLineArgs()
            .FirstOrDefault(a => a.StartsWith("--demo-theme=", StringComparison.OrdinalIgnoreCase))?["--demo-theme=".Length..]
        ?? "Default";

    public static HttpMessageHandler CreateHandler() => new DemoHandler();

    /// <summary>Fresh settings for every demo launch: three Contoso profiles and a recent query history.</summary>
    public static AppSettings CreateSettings()
    {
        try
        {
            if (Directory.Exists(DataFolder)) Directory.Delete(DataFolder, recursive: true);
        }
        catch (IOException)
        {
        }

        var intranet = new ConnectionProfile
        {
            Name = "Contoso Intranet",
            SiteUrl = "https://contoso.sharepoint.com/sites/intranet",
            UserName = "megan.bowen@contoso.com",
            AccountId = "demo",
        };
        var data = new AppSettings
        {
            Profiles =
            [
                intranet,
                new ConnectionProfile
                {
                    Name = "Mark 8 Project Team",
                    SiteUrl = "https://contoso.sharepoint.com/sites/Mark8ProjectTeam",
                    UserName = "megan.bowen@contoso.com",
                    AccountId = "demo",
                },
                new ConnectionProfile
                {
                    Name = "Contoso Archive (on-premises)",
                    SiteUrl = "https://sharepoint.contoso.local/sites/archive",
                    AuthMode = AuthMode.Windows,
                },
            ],
            CurrentProfileId = intranet.Id,
            ShowAdvancedPages = true,
            UseFileTypeIcons = true,
            Theme = Theme,
            History =
            [
                "Mark 8",
                "budget FileType:xlsx",
                "Author:\"Megan Bowen\"",
                "onboarding",
                "_api/search/query?querytext='budget'&rowlimit=25&refiners='FileType,Author'&sortlist='LastModifiedTime:descending'&clienttype='ContentSearchRegular'",
                "ContentClass:STS_Site",
                "LastModifiedTime>2026-09-01 FileType:pptx",
                "brand guidelines",
                "Path:\"https://contoso.sharepoint.com/sites/Legal\"",
                "expense policy",
                "IsDocument:false",
            ],
        };
        return data;
    }

    #region Fake SharePoint endpoints

    private sealed class DemoHandler : HttpMessageHandler
    {
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            await Task.Delay(Random.Shared.Next(150, 380), ct);

            var uri = request.RequestUri!;
            var path = Uri.UnescapeDataString(uri.AbsolutePath).TrimEnd('/');
            bool verbose = request.Headers.Accept.ToString().Contains("verbose", StringComparison.OrdinalIgnoreCase);

            JsonNode body;
            if (path.EndsWith("/_api/contextinfo", StringComparison.OrdinalIgnoreCase))
            {
                body = new JsonObject
                {
                    ["d"] = new JsonObject
                    {
                        ["GetContextWebInformation"] = new JsonObject
                        {
                            ["FormDigestValue"] = $"0x{Convert.ToHexString(RandomNumberGenerator.GetBytes(32))},{DateTime.UtcNow:R}",
                            ["FormDigestTimeoutSeconds"] = 1800,
                            ["LibraryVersion"] = "16.0.10417.20017",
                        },
                    },
                };
                verbose = true;
            }
            else if (path.EndsWith("/_api/search/postquery", StringComparison.OrdinalIgnoreCase))
            {
                var json = request.Content == null ? "{}" : await request.Content.ReadAsStringAsync(ct);
                body = Search(QueryRequest.FromPostBody(json), verbose, "postquery");
            }
            else if (path.EndsWith("/_api/search/query", StringComparison.OrdinalIgnoreCase))
            {
                body = Search(QueryRequest.FromQueryString(uri.Query), verbose, "query");
            }
            else if (path.EndsWith("/_api/search/suggest", StringComparison.OrdinalIgnoreCase))
            {
                body = Suggest(QueryRequest.FromQueryString(uri.Query).QueryText, verbose);
            }
            else
            {
                var error = new JsonObject
                {
                    ["odata.error"] = new JsonObject
                    {
                        ["code"] = "-1, Microsoft.SharePoint.Client.InvalidClientQueryException",
                        ["message"] = new JsonObject { ["lang"] = "en-US", ["value"] = $"The endpoint '{path}' is not supported." },
                    },
                };
                return Response(HttpStatusCode.NotFound, error, verbose: false);
            }

            return Response(HttpStatusCode.OK, body, verbose);
        }

        private static HttpResponseMessage Response(HttpStatusCode status, JsonNode body, bool verbose)
        {
            var response = new HttpResponseMessage(status)
            {
                Version = new Version(2, 0),
                Content = new StringContent(body.ToJsonString(new JsonSerializerOptions { WriteIndented = false }), Encoding.UTF8),
            };
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json")
            {
                CharSet = "utf-8",
                Parameters = { new("odata", verbose ? "verbose" : "nometadata"), new("streaming", "true") },
            };
            var requestId = Guid.NewGuid().ToString();
            response.Headers.CacheControl = new System.Net.Http.Headers.CacheControlHeaderValue { Private = true, MaxAge = TimeSpan.Zero };
            response.Headers.TryAddWithoutValidation("X-SharePointHealthScore", "1");
            response.Headers.TryAddWithoutValidation("X-SP-SERVERSTATE", "ReadOnly=0");
            response.Headers.TryAddWithoutValidation("SPRequestDuration", Random.Shared.Next(60, 240).ToString(CultureInfo.InvariantCulture));
            response.Headers.TryAddWithoutValidation("SPIisLatency", Random.Shared.Next(0, 3).ToString(CultureInfo.InvariantCulture));
            response.Headers.TryAddWithoutValidation("X-MS-Correlation-Id", requestId);
            response.Headers.TryAddWithoutValidation("X-MSEdge-Ref", $"Ref A: {Convert.ToHexString(RandomNumberGenerator.GetBytes(16))} Ref B: SEA30EDGE0815 Ref C: {DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}");
            response.Headers.TryAddWithoutValidation("SPRequestGuid", requestId);
            response.Headers.TryAddWithoutValidation("request-id", requestId);
            response.Headers.TryAddWithoutValidation("MS-CV", Convert.ToBase64String(RandomNumberGenerator.GetBytes(12)) + ".0");
            response.Headers.TryAddWithoutValidation("Strict-Transport-Security", "max-age=31536000");
            response.Headers.TryAddWithoutValidation("X-FRAME-OPTIONS", "SAMEORIGIN");
            response.Headers.TryAddWithoutValidation("X-Content-Type-Options", "nosniff");
            response.Headers.TryAddWithoutValidation("MicrosoftSharePointTeamServices", "16.0.0.26511");
            response.Headers.TryAddWithoutValidation("X-MS-InvokeApp", "1; RequireReadOnly");
            response.Headers.TryAddWithoutValidation("X-Cache", "CONFIG_NOCACHE");
            response.Headers.Date = DateTimeOffset.UtcNow;
            return response;
        }
    }

    /// <summary>The parameters of a query or postquery request.</summary>
    private sealed class QueryRequest
    {
        public string QueryText { get; set; } = "*";
        public string? QueryTemplate { get; set; }
        public int StartRow { get; set; }
        public int RowLimit { get; set; } = 10;
        public List<string> SelectProperties { get; set; } = [];
        public string? Refiners { get; set; }
        public List<string> RefinementFilters { get; set; } = [];
        public string? HiddenConstraints { get; set; }
        public string? SortList { get; set; }

        public static QueryRequest FromQueryString(string query)
        {
            var r = new QueryRequest();
            foreach (var pair in query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
            {
                int eq = pair.IndexOf('=');
                if (eq < 0) continue;
                var key = pair[..eq].ToLowerInvariant();
                var value = Uri.UnescapeDataString(pair[(eq + 1)..].Replace('+', ' ')).Trim();
                if (value.Length >= 2 && value[0] == '\'' && value[^1] == '\'') value = value[1..^1].Replace("''", "'");
                switch (key)
                {
                    case "querytext": r.QueryText = value; break;
                    case "querytemplate": r.QueryTemplate = value; break;
                    case "startrow": r.StartRow = ParseInt(value, 0); break;
                    case "rowlimit": r.RowLimit = ParseInt(value, 10); break;
                    case "selectproperties": r.SelectProperties = SplitList(value); break;
                    case "refiners": r.Refiners = value; break;
                    case "refinementfilters": if (value.Length > 0) r.RefinementFilters.Add(value); break;
                    case "hiddenconstraints": r.HiddenConstraints = value; break;
                    case "sortlist": r.SortList = value; break;
                }
            }
            return r;
        }

        public static QueryRequest FromPostBody(string json)
        {
            var r = new QueryRequest();
            using var doc = JsonDocument.Parse(json);
            var root = SearchService.Prop(doc.RootElement, "request") ?? doc.RootElement;
            string? Str(string name) => SearchService.Prop(root, name) is { ValueKind: JsonValueKind.String } v ? v.GetString() : null;
            int Int(string name, int fallback) => SearchService.Prop(root, name) is { ValueKind: JsonValueKind.Number } v ? v.GetInt32() : fallback;
            List<string> List(string name) => SearchService.Prop(root, name) is { ValueKind: JsonValueKind.Array } a
                ? a.EnumerateArray().Select(e => e.GetString() ?? "").Where(s => s.Length > 0).ToList()
                : [];

            r.QueryText = Str("Querytext") ?? "*";
            r.QueryTemplate = Str("QueryTemplate");
            r.StartRow = Int("StartRow", 0);
            r.RowLimit = Int("RowLimit", 10);
            r.SelectProperties = List("SelectProperties");
            r.Refiners = Str("Refiners");
            r.RefinementFilters = List("RefinementFilters");
            r.HiddenConstraints = Str("HiddenConstraints");
            r.SortList = Str("SortList");
            return r;
        }

        private static int ParseInt(string value, int fallback) =>
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : fallback;

        private static List<string> SplitList(string value) =>
            value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
    }

    private static readonly string[] s_alwaysReturned =
        ["Rank", "DocId", "WorkId", "PartitionId", "UrlZone", "Culture", "ResultTypeId", "RenderTemplateId"];

    private static JsonNode Search(QueryRequest q, bool verbose, string wrapper)
    {
        var text = string.IsNullOrWhiteSpace(q.QueryText) ? "*" : q.QueryText;
        var kql = string.IsNullOrWhiteSpace(q.QueryTemplate) ? text : q.QueryTemplate.Replace("{searchTerms}", text, StringComparison.OrdinalIgnoreCase);
        var query = Kql.Parse(kql);
        var hidden = string.IsNullOrWhiteSpace(q.HiddenConstraints) ? null : Kql.Parse(q.HiddenConstraints);
        var filters = q.RefinementFilters.Select(RefinementFilter.Parse).ToList();

        var matches = Dataset.Items
            .Where(i => query.Matches(i) && (hidden?.Matches(i) ?? true) && filters.All(f => f(i)))
            .Select(i => (Item: i, Rank: query.Score(i)))
            .ToList();
        matches = Sort(matches, q.SortList);

        var page = matches.Skip(Math.Max(0, q.StartRow)).Take(Math.Max(0, q.RowLimit)).ToList();
        var select = q.SelectProperties.Count > 0 ? q.SelectProperties : SearchService.DefaultSelectProperties.ToList();
        var terms = query.FreeTextTerms;

        var rows = new JsonArray();
        foreach (var (item, rank) in page)
        {
            var cells = new JsonArray();
            foreach (var name in select.Concat(s_alwaysReturned).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                string? value = name.Equals("Rank", StringComparison.OrdinalIgnoreCase)
                    ? rank.ToString("0.0000000", CultureInfo.InvariantCulture)
                    : name.Equals("HitHighlightedSummary", StringComparison.OrdinalIgnoreCase)
                        ? Highlight(item.Summary, terms)
                        : item.Managed.TryGetValue(name, out var v) ? v : null;
                var cell = new JsonObject { ["Key"] = name, ["Value"] = value, ["ValueType"] = value == null ? "Null" : ValueType(name) };
                if (verbose) cell["__metadata"] = new JsonObject { ["type"] = "SP.KeyValue" };
                cells.Add(cell);
            }
            rows.Add(new JsonObject { ["Cells"] = List(cells, verbose) });
        }

        var refiners = new JsonArray();
        foreach (var name in SplitRefiners(q.Refiners))
        {
            var entries = new JsonArray();
            foreach (var (value, count) in RefinerValues(matches.Select(m => m.Item), name))
            {
                entries.Add(new JsonObject
                {
                    ["RefinementCount"] = count.ToString(CultureInfo.InvariantCulture),
                    ["RefinementName"] = value,
                    ["RefinementToken"] = Token(value),
                    ["RefinementValue"] = value,
                });
            }
            refiners.Add(new JsonObject { ["Name"] = name, ["Entries"] = List(entries, verbose) });
        }

        var correlationId = Guid.NewGuid().ToString();
        var properties = new JsonArray
        {
            KeyValue("RowLimit", q.RowLimit.ToString(CultureInfo.InvariantCulture), "Edm.Int32", verbose),
            KeyValue("SourceId", "8413cd39-2156-4e00-b54d-11efd9abdb89", "Edm.Guid", verbose),
            KeyValue("CorrelationId", correlationId, "Edm.Guid", verbose),
            KeyValue("WasGroupRestricted", "False", "Edm.Boolean", verbose),
            KeyValue("IsPartial", "False", "Edm.Boolean", verbose),
            KeyValue("InternalRequestId", Guid.NewGuid().ToString(), "Edm.Guid", verbose),
            KeyValue("IsMissingUnifiedGroups", "False", "Edm.Boolean", verbose),
            KeyValue("MultiGeoSearchStatus", "Full", "Edm.String", verbose),
            KeyValue("SerializedQuery", $"<Query Culture=\"en-US\" EnableStemming=\"True\" EnablePhonetic=\"False\" EnableNicknames=\"False\" IgnoreAllNoiseQuery=\"True\" SummaryLength=\"180\" MaxSnippetLength=\"180\" DesiredSnippetLength=\"90\" KeywordInclusion=\"0\" QueryText=\"{WebUtility.HtmlEncode(text)}\" QueryTemplate=\"{WebUtility.HtmlEncode(q.QueryTemplate ?? "")}\" TrimDuplicates=\"True\" Site=\"{Dataset.SiteCollectionId}\" Web=\"{Dataset.WebId}\" KeywordType=\"True\" HiddenConstraints=\"{WebUtility.HtmlEncode(q.HiddenConstraints ?? "")}\" />", "Edm.String", verbose),
        };

        var primary = new JsonObject
        {
            ["CustomResults"] = List(new JsonArray(), verbose),
            ["QueryId"] = Guid.NewGuid().ToString(),
            ["QueryRuleId"] = "00000000-0000-0000-0000-000000000000",
            ["RefinementResults"] = refiners.Count == 0 ? null : new JsonObject
            {
                ["GroupTemplateId"] = null,
                ["ItemTemplateId"] = null,
                ["Properties"] = List(new JsonArray(), verbose),
                ["Refiners"] = List(refiners, verbose),
                ["ResultTitle"] = null,
                ["ResultTitleUrl"] = null,
            },
            ["RelevantResults"] = new JsonObject
            {
                ["GroupTemplateId"] = null,
                ["ItemTemplateId"] = null,
                ["Properties"] = List(new JsonArray
                {
                    KeyValue("GenerationId", "9223372036854775806", "Edm.Int64", verbose),
                    KeyValue("indexSystem", "", "Edm.String", verbose),
                    KeyValue("ExecutionTimeMs", Random.Shared.Next(20, 90).ToString(CultureInfo.InvariantCulture), "Edm.Int32", verbose),
                    KeyValue("QueryModification", $"{kql} -ContentClass=urn:content-class:SPSPeople", "Edm.String", verbose),
                    KeyValue("RenderTemplateId", "~sitecollection/_catalogs/masterpage/Display Templates/Search/Group_Default.js", "Edm.String", verbose),
                    KeyValue("StartRecord", q.StartRow.ToString(CultureInfo.InvariantCulture), "Edm.Int32", verbose),
                    KeyValue("IsLastBlockInSubstrate", "True", "Edm.Boolean", verbose),
                    KeyValue("IsFirstBlockInSubstrate", "True", "Edm.Boolean", verbose),
                }, verbose),
                ["ResultTitle"] = null,
                ["ResultTitleUrl"] = null,
                ["RowCount"] = page.Count,
                ["Table"] = new JsonObject { ["Rows"] = List(rows, verbose) },
                ["TotalRows"] = matches.Count,
                ["TotalRowsIncludingDuplicates"] = matches.Count,
            },
            ["SpecialTermResults"] = null,
        };

        var result = new JsonObject
        {
            ["ElapsedTime"] = Random.Shared.Next(70, 260),
            ["PrimaryQueryResult"] = primary,
            ["Properties"] = List(properties, verbose),
            ["SecondaryQueryResults"] = List(new JsonArray(), verbose),
            ["SpellingSuggestion"] = "",
            ["TriggeredRules"] = List(new JsonArray(), verbose),
        };
        if (!verbose) return result;
        result["__metadata"] = new JsonObject { ["type"] = "Microsoft.Office.Server.Search.REST.SearchResult" };
        return new JsonObject { ["d"] = new JsonObject { [wrapper] = result } };
    }

    private static JsonNode Suggest(string text, bool verbose)
    {
        text = text.Trim();
        bool Hit(string s) => text.Length == 0 || s.Contains(text, StringComparison.OrdinalIgnoreCase);

        var queries = new JsonArray();
        foreach (var q in Dataset.PopularQueries.Where(Hit).Take(5))
            queries.Add(new JsonObject { ["IsPersonal"] = false, ["Query"] = q, ["QueryHighlight"] = Highlight(q, [text]) });

        var people = new JsonArray();
        foreach (var p in Dataset.People.Select(p => p.Name).Where(Hit).Take(4)) people.Add(p);

        var personal = new JsonArray();
        foreach (var item in Dataset.Items.Where(i => i.Get("IsDocument") == "true" && Hit(i.Get("Title"))).Take(4))
        {
            personal.Add(new JsonObject
            {
                ["HighlightedTitle"] = Highlight(item.Get("Title"), [text]),
                ["Id"] = null,
                ["IsBestBet"] = false,
                ["Title"] = item.Get("Title"),
                ["TypeId"] = null,
                ["Url"] = item.Get("Path"),
            });
        }

        var result = new JsonObject
        {
            ["PeopleNames"] = List(people, verbose),
            ["PersonalResults"] = List(personal, verbose),
            ["PopularQueries"] = List(new JsonArray(), verbose),
            ["Queries"] = List(queries, verbose),
        };
        return verbose ? new JsonObject { ["d"] = new JsonObject { ["suggest"] = result } } : result;
    }

    private static List<(DemoItem Item, double Rank)> Sort(List<(DemoItem Item, double Rank)> matches, string? sortList)
    {
        var keys = (sortList ?? "")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => s.Split(':'))
            .Where(p => p.Length == 2)
            .ToList();
        if (keys.Count == 0) return matches.OrderByDescending(m => m.Rank).ToList();

        IOrderedEnumerable<(DemoItem Item, double Rank)>? ordered = null;
        foreach (var key in keys)
        {
            var name = key[0];
            bool desc = key[1].StartsWith("desc", StringComparison.OrdinalIgnoreCase);
            Func<(DemoItem Item, double Rank), IComparable> selector = name.Equals("Rank", StringComparison.OrdinalIgnoreCase)
                ? m => m.Rank
                : m => SortKey(m.Item.Get(name));
            ordered = ordered == null
                ? desc ? matches.OrderByDescending(selector) : matches.OrderBy(selector)
                : desc ? ordered.ThenByDescending(selector) : ordered.ThenBy(selector);
        }
        return ordered!.ToList();

        static IComparable SortKey(string value) =>
            double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d
            : DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dt) ? dt.Ticks
            : (IComparable)value.ToLowerInvariant().GetHashCode();
    }

    private static IEnumerable<string> SplitRefiners(string? refiners)
    {
        if (string.IsNullOrWhiteSpace(refiners)) yield break;
        int depth = 0, start = 0;
        for (int i = 0; i <= refiners.Length; i++)
        {
            if (i < refiners.Length && refiners[i] == '(') depth++;
            else if (i < refiners.Length && refiners[i] == ')') depth--;
            else if (i == refiners.Length || (refiners[i] == ',' && depth == 0))
            {
                var part = refiners[start..i].Trim();
                int paren = part.IndexOf('(');
                if (paren >= 0) part = part[..paren];
                if (part.Length > 0) yield return part;
                start = i + 1;
            }
        }
    }

    private static IEnumerable<(string Value, int Count)> RefinerValues(IEnumerable<DemoItem> items, string refiner)
    {
        IEnumerable<string> Values(DemoItem item)
        {
            if (refiner.Equals("ManagedProperties", StringComparison.OrdinalIgnoreCase)) return item.Managed.Keys;
            if (refiner.Equals("CrawledProperties", StringComparison.OrdinalIgnoreCase)) return item.Crawled;
            var value = item.Get(refiner);
            return value.Length == 0 ? [] : value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }

        return items
            .SelectMany(Values)
            .GroupBy(v => v, StringComparer.OrdinalIgnoreCase)
            .Select(g => (g.Key, g.Count()))
            .OrderByDescending(e => e.Item2)
            .ThenBy(e => e.Key, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>A refiner value as SharePoint tokenizes it: "ǂǂ" and the value's UTF-8 bytes in hex.</summary>
    private static string Token(string value) => $"\"ǂǂ{Convert.ToHexString(Encoding.UTF8.GetBytes(value)).ToLowerInvariant()}\"";

    private static string Highlight(string text, IReadOnlyList<string> terms)
    {
        foreach (var term in terms.Where(t => t.Length > 0))
            text = Regex.Replace(text, $@"(?<![\w>])({Regex.Escape(term.TrimEnd('*'))}\w*)", "<c0>$1</c0>", RegexOptions.IgnoreCase);
        return text + " <ddd/>";
    }

    private static JsonNode List(JsonArray items, bool verbose) => verbose ? new JsonObject { ["results"] = items } : items;

    private static JsonObject KeyValue(string key, string value, string type, bool verbose)
    {
        var kv = new JsonObject { ["Key"] = key, ["Value"] = value, ["ValueType"] = type };
        if (verbose) kv["__metadata"] = new JsonObject { ["type"] = "SP.KeyValue" };
        return kv;
    }

    private static readonly HashSet<string> s_dates = new(StringComparer.OrdinalIgnoreCase) { "LastModifiedTime", "Created", "Write", "FirstPublishedDate" };
    private static readonly HashSet<string> s_numbers = new(StringComparer.OrdinalIgnoreCase)
    {
        "Size", "DocId", "WorkId", "ListItemID", "ViewsLifeTime", "ViewsRecent", "ViewsLifeTimeUniqueUsers", "ViewsRecentUniqueUsers", "UrlZone", "ResultTypeId",
    };
    private static readonly HashSet<string> s_bools = new(StringComparer.OrdinalIgnoreCase) { "IsDocument", "IsContainer", "IsListItem", "IsExternalContent", "IsHubSite" };
    private static readonly HashSet<string> s_guids = new(StringComparer.OrdinalIgnoreCase)
    {
        "UniqueId", "SiteId", "WebId", "ListId", "PartitionId", "NormUniqueID", "NormSiteID", "NormWebID", "NormListID", "GroupId", "DepartmentId",
    };

    private static string ValueType(string name) =>
        s_dates.Contains(name) ? "Edm.DateTime"
        : name.Equals("Rank", StringComparison.OrdinalIgnoreCase) ? "Edm.Double"
        : s_numbers.Contains(name) ? "Edm.Int64"
        : s_bools.Contains(name) ? "Edm.Boolean"
        : s_guids.Contains(name) ? "Edm.Guid"
        : "Edm.String";

    #endregion

    #region KQL and refinement filters

    /// <summary>
    /// Enough KQL for the demo: free-text terms and phrases, property restrictions (<c>:</c>, <c>=</c>,
    /// <c>&lt;&gt;</c>, <c>&lt;</c>, <c>&gt;</c>), trailing <c>*</c>, <c>-</c>/<c>NOT</c>, and <c>OR</c> between groups of terms.
    /// </summary>
    private sealed class Kql
    {
        private readonly List<List<(bool Negated, string Term)>> _groups = [];

        public List<string> FreeTextTerms { get; } = [];

        public static Kql Parse(string text)
        {
            var kql = new Kql();
            var group = new List<(bool, string)>();
            kql._groups.Add(group);
            bool negateNext = false;
            foreach (var token in Tokenize(text))
            {
                if (token == "AND") continue;
                if (token == "OR") { group = []; kql._groups.Add(group); continue; }
                if (token == "NOT") { negateNext = true; continue; }

                bool negated = negateNext || token.StartsWith('-');
                negateNext = false;
                var term = token.TrimStart('-', '+');
                if (term.Length == 0) continue;
                group.Add((negated, term));
                if (!negated && term != "*" && !s_restriction.IsMatch(term)) kql.FreeTextTerms.Add(Unquote(term));
            }
            return kql;
        }

        public bool Matches(DemoItem item) =>
            _groups.Any(g => g.All(t => TermMatches(item, t.Term) != t.Negated));

        /// <summary>A SharePoint-looking rank: higher for matches in the title than in the body.</summary>
        public double Score(DemoItem item)
        {
            double score = item.Popularity;
            foreach (var term in FreeTextTerms)
            {
                if (Contains(item.Get("Title"), term)) score += 4.1;
                if (Contains(item.Get("Filename"), term)) score += 1.3;
                if (Contains(item.Summary, term)) score += 1.7;
                if (Contains(item.Get("Tags"), term)) score += 0.9;
            }
            return 10 + score + (item.Get("WorkId").GetHashCode() & 0xFFFF) / 1_000_000.0;
        }

        private static readonly Regex s_restriction = new(@"^([A-Za-z_][\w.]*)(<>|>=|<=|:|=|>|<)(.*)$");

        private static bool TermMatches(DemoItem item, string term)
        {
            if (term == "*") return true;
            var m = s_restriction.Match(term);
            if (!m.Success) return Contains(item.FullText, Unquote(term));

            var name = m.Groups[1].Value;
            var op = m.Groups[2].Value;
            var value = Unquote(m.Groups[3].Value);
            var actual = item.Get(name);
            if (name.Equals("contentclass", StringComparison.OrdinalIgnoreCase) && value.StartsWith("urn:content-class:", StringComparison.OrdinalIgnoreCase))
                return op == "<>" ^ value.EndsWith("SPSPeople", StringComparison.OrdinalIgnoreCase);

            switch (op)
            {
                case ":":
                case "=":
                    if (actual.Length == 0) return false;
                    if (name.Equals("Path", StringComparison.OrdinalIgnoreCase) || value.EndsWith('*'))
                        return actual.StartsWith(value.TrimEnd('*'), StringComparison.OrdinalIgnoreCase) || (op == ":" && Contains(actual, value));
                    return actual.Equals(value, StringComparison.OrdinalIgnoreCase)
                        || (op == ":" && actual.Split(';').Any(a => a.Trim().Equals(value, StringComparison.OrdinalIgnoreCase)))
                        || (op == ":" && Contains(actual, value));
                case "<>":
                    return !actual.Equals(value, StringComparison.OrdinalIgnoreCase);
                default:
                    int cmp;
                    if (double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out var a)
                        && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var b))
                        cmp = a.CompareTo(b);
                    else if (DateTime.TryParse(actual, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var da)
                        && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var db))
                        cmp = da.ToUniversalTime().CompareTo(db);
                    else
                        return false;
                    return op switch { ">" => cmp > 0, "<" => cmp < 0, ">=" => cmp >= 0, _ => cmp <= 0 };
            }
        }

        /// <summary>Whole-word (or word-prefix, for a trailing *) match, case-insensitive.</summary>
        private static bool Contains(string haystack, string needle)
        {
            if (needle.Length == 0) return true;
            bool prefix = needle.EndsWith('*');
            var pattern = $@"(?<!\w){Regex.Escape(needle.TrimEnd('*'))}" + (prefix ? "" : @"(?!\w)");
            return Regex.IsMatch(haystack, pattern, RegexOptions.IgnoreCase);
        }

        private static string Unquote(string s) => s.Length >= 2 && s[0] == '"' && s[^1] == '"' ? s[1..^1] : s.Trim('"');

        private static IEnumerable<string> Tokenize(string text)
        {
            var current = new StringBuilder();
            bool inQuotes = false;
            foreach (var c in text)
            {
                if (c == '"') inQuotes = !inQuotes;
                if (!inQuotes && (char.IsWhiteSpace(c) || c == '(' || c == ')'))
                {
                    if (current.Length > 0) yield return current.ToString();
                    current.Clear();
                    continue;
                }
                current.Append(c);
            }
            if (current.Length > 0) yield return current.ToString();
        }
    }

    /// <summary>
    /// Parses FQL-style refinement filters, e.g. <c>FileType:or("ǂǂ646f6378","ǂǂ786c7378")</c> or
    /// <c>andnot(ContentClass:"ǂǂ…",FileType:or(…))</c>, into a predicate.
    /// </summary>
    private sealed class RefinementFilter(string text)
    {
        private int _i;

        public static Func<DemoItem, bool> Parse(string text)
        {
            try
            {
                return new RefinementFilter(text).Expression();
            }
            catch (Exception)
            {
                return _ => true;
            }
        }

        private Func<DemoItem, bool> Expression()
        {
            SkipSpaces();
            var name = Identifier();
            SkipSpaces();
            if (Peek == '(')
            {
                var args = Arguments(Expression);
                return name.ToLowerInvariant() switch
                {
                    "and" or "all" => item => args.All(a => a(item)),
                    "or" or "any" => item => args.Any(a => a(item)),
                    "not" => item => !args[0](item),
                    "andnot" => item => args[0](item) && !args.Skip(1).Any(a => a(item)),
                    _ => _ => true,
                };
            }

            Expect(':');
            SkipSpaces();
            if (Peek == '"')
            {
                var value = String();
                return item => Equal(item, name, value);
            }
            var function = Identifier().ToLowerInvariant();
            var values = Arguments(String);
            return function == "and"
                ? item => values.All(v => Equal(item, name, v))
                : item => values.Any(v => Equal(item, name, v));
        }

        private static bool Equal(DemoItem item, string name, string value) =>
            item.Get(name).Split(';').Any(v => v.Trim().Equals(value, StringComparison.OrdinalIgnoreCase));

        private List<T> Arguments<T>(Func<T> parse)
        {
            Expect('(');
            var list = new List<T>();
            while (true)
            {
                SkipSpaces();
                if (Peek == ')') { _i++; break; }
                list.Add(parse());
                SkipSpaces();
                if (Peek == ',') _i++;
            }
            return list;
        }

        private string String()
        {
            SkipSpaces();
            Expect('"');
            var sb = new StringBuilder();
            while (_i < text.Length && text[_i] != '"')
            {
                if (text[_i] == '\\' && _i + 1 < text.Length) _i++;
                sb.Append(text[_i++]);
            }
            Expect('"');
            var s = sb.ToString();
            return s.StartsWith("ǂǂ", StringComparison.Ordinal)
                ? Encoding.UTF8.GetString(Convert.FromHexString(s[2..]))
                : s;
        }

        private string Identifier()
        {
            int start = _i;
            while (_i < text.Length && (char.IsLetterOrDigit(text[_i]) || text[_i] is '_' or '.')) _i++;
            return text[start.._i];
        }

        private char Peek => _i < text.Length ? text[_i] : '\0';

        private void SkipSpaces()
        {
            while (_i < text.Length && char.IsWhiteSpace(text[_i])) _i++;
        }

        private void Expect(char c)
        {
            if (Peek != c) throw new FormatException($"Expected '{c}' at {_i}.");
            _i++;
        }
    }

    #endregion

    #region Contoso dataset

    private sealed class DemoItem
    {
        public Dictionary<string, string> Managed { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> Crawled { get; } = [];
        public string Summary { get; set; } = "";
        public double Popularity { get; set; }

        private string? _fullText;

        /// <summary>The text free-text terms are matched against.</summary>
        public string FullText => _fullText ??= string.Join(" ",
            Get("Title"), Get("Filename"), Summary, Get("Author"), Get("ModifiedBy"), Get("SiteTitle"), Get("Tags"), Get("Description"));

        public string Get(string name) => Managed.TryGetValue(name, out var v) ? v : "";
    }

    private sealed record Person(string Name, string Email, string JobTitle, string Department);

    private sealed record Site(string Key, string Url, string Title, string Description, string Template, string Department);

    private static class Dataset
    {
        private const string Tenant = "https://contoso.sharepoint.com";

        public static Guid SiteCollectionId => G("siteintranet");
        public static Guid WebId => G("webintranet");

        // Property sets used for crawled property names (see SearchService.s_propertySets).
        private const string SetSharePoint = "00130329-0000-0130-c000-000000131346";
        private const string SetSharePointRaw = "00140329-0000-0140-c000-000000141446";
        private const string SetSharePointSite = "158d7563-aeff-4dbf-bf16-4a1445f0366c";
        private const string SetBasic = "0b63e343-9ccc-11d0-bcdb-00805fccce04";
        private const string SetInternal = "012357bd-1113-171d-1f25-292bb0b0b0b0";
        private const string SetOffice = "F29F85E0-4FF9-1068-AB91-08002B27B3D9";
        private const string SetParser = "64ae120f-487d-445a-8d5a-5258f99cb970";
        private const string SetExtractor = "e835446c-937b-4492-95f3-d89988b01039";

        public static readonly Person[] People =
        [
            new("Megan Bowen", "megan.bowen@contoso.com", "Marketing Manager", "Marketing"),
            new("Adele Vance", "adele.vance@contoso.com", "Retail Manager", "Retail"),
            new("Alex Wilber", "alex.wilber@contoso.com", "Marketing Assistant", "Marketing"),
            new("Diego Siciliani", "diego.siciliani@contoso.com", "HR Manager", "Human Resources"),
            new("Isaiah Langer", "isaiah.langer@contoso.com", "Sales Rep", "Sales"),
            new("Johanna Lorenz", "johanna.lorenz@contoso.com", "Senior Engineer", "Engineering"),
            new("Lee Gu", "lee.gu@contoso.com", "Director", "Manufacturing"),
            new("Lidia Holloway", "lidia.holloway@contoso.com", "Product Manager", "Engineering"),
            new("Lynne Robbins", "lynne.robbins@contoso.com", "Planner", "Retail"),
            new("Miriam Graham", "miriam.graham@contoso.com", "Director", "Sales and Marketing"),
            new("Nestor Wilke", "nestor.wilke@contoso.com", "Director of Operations", "Operations"),
            new("Patti Fernandez", "patti.fernandez@contoso.com", "President", "Executive Management"),
            new("Pradeep Gupta", "pradeep.gupta@contoso.com", "Accountant", "Finance"),
            new("Grady Archie", "grady.archie@contoso.com", "Designer", "Research and Development"),
            new("Joni Sherman", "joni.sherman@contoso.com", "Paralegal", "Legal"),
            new("Christie Cline", "christie.cline@contoso.com", "Buyer", "Retail"),
            new("Allan Deyoung", "allan.deyoung@contoso.com", "IT Admin", "IT"),
            new("Henrietta Mueller", "henrietta.mueller@contoso.com", "Developer", "Engineering"),
            new("Irvin Sayers", "irvin.sayers@contoso.com", "Senior Engineer", "Engineering"),
        ];

        private static readonly Site[] s_sites =
        [
            new("intranet", "/sites/intranet", "Contoso Intranet", "News, events and resources for everyone at Contoso.", "SITEPAGEPUBLISHING#0", "Communications"),
            new("mark8", "/sites/Mark8ProjectTeam", "Mark 8 Project Team", "Collaboration space for the Mark 8 drone program.", "GROUP#0", "Engineering"),
            new("marketing", "/sites/SalesAndMarketing", "Sales and Marketing", "Campaigns, brand assets and sales enablement.", "GROUP#0", "Marketing"),
            new("finance", "/sites/Finance", "Finance", "Budgets, forecasts, policies and financial reporting.", "GROUP#0", "Finance"),
            new("hr", "/sites/HumanResources", "Human Resources", "Policies, benefits and onboarding for Contoso employees.", "SITEPAGEPUBLISHING#0", "Human Resources"),
            new("legal", "/sites/Legal", "Legal", "Contracts, agreements and intellectual property.", "GROUP#0", "Legal"),
            new("retail", "/sites/Retail", "Retail", "Store operations, sales and merchandising.", "GROUP#0", "Retail"),
            new("it", "/sites/ITHelpDesk", "IT Help Desk", "How-to guides, tickets and IT announcements.", "SITEPAGEPUBLISHING#0", "IT"),
            new("eng", "/sites/Engineering", "Engineering", "Firmware, software and hardware engineering.", "GROUP#0", "Engineering"),
        ];

        public static readonly string[] PopularQueries =
        [
            "mark 8", "mark 8 launch plan", "mark 8 specification", "marketing budget", "fy27 budget", "brand guidelines",
            "benefits enrollment", "employee handbook", "expense policy", "onboarding checklist", "vpn setup", "store sales dashboard",
        ];

        public static readonly List<DemoItem> Items = Build();

        private static int s_nextId = 1;

        private static List<DemoItem> Build()
        {
            var items = new List<DemoItem>();

            foreach (var site in s_sites) items.Add(SiteItem(site));

            // Mark 8 Project Team
            Doc(items, "mark8", "", "Mark 8 Product Specification.docx", "Johanna Lorenz;Lidia Holloway", "Lidia Holloway", 120, 0, 862_412,
                "Technical specification for the Mark 8 drone: flight controller, 4K gimbal camera, 32-minute battery life and obstacle avoidance sensors. Includes performance targets and acceptance criteria.",
                "Mark 8;Specification;Engineering", 4.2, label: "Confidential");
            Doc(items, "mark8", "", "Mark 8 Launch Plan.pptx", "Lidia Holloway", "Megan Bowen", 60, 1, 6_389_120,
                "Go-to-market plan for the Mark 8 launch, covering pricing, channel partners, launch events and the press timeline for Q1 FY27.",
                "Mark 8;Launch;Marketing", 4.0);
            Doc(items, "mark8", "", "Mark 8 Bill of Materials.xlsx", "Irvin Sayers", "Johanna Lorenz", 90, 2, 397_304,
                "Component costs, supplier lead times and target unit cost for Mark 8 production builds, by assembly and supplier.",
                "Mark 8;Manufacturing;Suppliers", 3.1);
            Doc(items, "mark8", "", "Mark 8 Flight Test Results.xlsx", "Henrietta Mueller", "Irvin Sayers", 45, 3, 1_238_016,
                "Results from 140 Mark 8 test flights: hover stability, wind resistance, battery discharge curves and sensor accuracy.",
                "Mark 8;Testing", 2.8);
            Doc(items, "mark8", "", "Mark 8 Field Trial.mp4", "Grady Archie", "Grady Archie", 20, 18, 188_743_680,
                "Footage from the Mark 8 field trial at the Redmond test range, including autonomous return-to-home and obstacle avoidance.",
                "Mark 8;Video", 1.9);
            Doc(items, "mark8", "", "Mark 8 Hero Image.png", "Grady Archie", "Grady Archie", 30, 25, 2_406_400,
                "Product photography of the Mark 8 drone for the launch campaign and retail packaging.",
                "Mark 8;Images", 1.2);
            Doc(items, "mark8", "", "Mark 8 Project Plan.mpp", "Lynne Robbins", "Lynne Robbins", 150, 4, 626_688,
                "Mark 8 program schedule from design freeze to general availability, with milestones, dependencies and owners.",
                "Mark 8;Schedule", 2.2);
            Doc(items, "mark8", "", "Mark 8 System Architecture.vsdx", "Johanna Lorenz", "Johanna Lorenz", 110, 12, 962_560,
                "System architecture of the Mark 8: flight controller, sensor fusion, telemetry radio and companion app services.",
                "Mark 8;Architecture", 2.0);
            Doc(items, "mark8", "", "Mark 8 Safety Certification.pdf", "Irvin Sayers", "Joni Sherman", 70, 40, 3_174_400,
                "Safety certification package for the Mark 8, including FAA Part 107 compliance notes and battery safety testing.",
                "Mark 8;Compliance", 1.6);
            Doc(items, "mark8", "", "Mark 8 Engineering Notebook.one", "Henrietta Mueller", "Johanna Lorenz", 100, 6, 5_529_600,
                "Engineering notes for the Mark 8: daily stand-ups, design decisions and open issues from the firmware and hardware teams.",
                "Mark 8;Notes", 1.4);
            Doc(items, "mark8", "", "Mark 8 Customer Feedback Summary.docx", "Lidia Holloway", "Alex Wilber", 14, 5, 214_016,
                "Summary of beta customer feedback on the Mark 8: ease of setup, camera quality, battery life and app experience.",
                "Mark 8;Customers;Research", 2.6);
            Folder(items, "mark8", "Design Reviews", "Johanna Lorenz", 130, 0);
            Doc(items, "mark8", "Design Reviews", "Mark 8 Design Review - October.pptx", "Johanna Lorenz", "Irvin Sayers", 5, 0, 4_718_592,
                "October design review for the Mark 8: gimbal redesign, motor thermal margins and the updated landing gear.",
                "Mark 8;Design Review", 3.4);
            Doc(items, "mark8", "Design Reviews", "Mark 8 Design Review - September.pptx", "Johanna Lorenz", "Johanna Lorenz", 36, 33, 4_194_304,
                "September design review for the Mark 8: propeller noise reduction, battery enclosure and EMI test findings.",
                "Mark 8;Design Review", 2.1);
            Library(items, "mark8", "Shared Documents", "Documents", 160, 0);
            ListItem(items, "mark8", "Tasks", "STS_ListItem_Tasks", 12, "Finalize battery supplier contract", "Lynne Robbins", "Joni Sherman", 21, 2,
                "Agree volume pricing and delivery schedule with the battery cell supplier for Mark 8 production.", 1.0);
            ListItem(items, "mark8", "Tasks", "STS_ListItem_Tasks", 15, "Complete FAA certification paperwork", "Irvin Sayers", "Irvin Sayers", 18, 1,
                "Submit the remaining Mark 8 certification documents and schedule the inspection.", 1.0);
            ListItem(items, "mark8", "Tasks", "STS_ListItem_Tasks", 19, "Order packaging samples", "Christie Cline", "Christie Cline", 9, 4,
                "Request three packaging samples for the Mark 8 retail box from approved vendors.", 0.6);

            // Sales and Marketing
            Doc(items, "marketing", "", "FY27 Marketing Budget.xlsx", "Megan Bowen", "Pradeep Gupta", 40, 1, 219_136,
                "Quarterly marketing budget by campaign and region, with actuals against plan for Q1 FY27 and the Mark 8 launch reserve.",
                "Budget;FY27;Marketing", 3.6);
            Doc(items, "marketing", "", "Brand Guidelines 2026.pdf", "Grady Archie", "Megan Bowen", 200, 60, 9_113_600,
                "Contoso brand guidelines: logo usage, colour palette, typography, photography style and voice and tone.",
                "Brand;Guidelines", 3.0);
            Doc(items, "marketing", "", "Contoso Brand Toolkit.zip", "Grady Archie", "Grady Archie", 180, 90, 52_428_800,
                "Logos, templates, icons and fonts for Contoso marketing materials.",
                "Brand", 1.1);
            Folder(items, "marketing", "Campaigns", "Megan Bowen", 300, 1);
            Doc(items, "marketing", "Campaigns", "Q3 Campaign Performance.pptx", "Alex Wilber", "Megan Bowen", 21, 7, 3_670_016,
                "Results of the Q3 campaigns: reach, engagement, leads and cost per lead by channel.",
                "Campaigns;Reporting", 2.4);
            Doc(items, "marketing", "Campaigns", "Social Media Calendar.xlsx", "Alex Wilber", "Alex Wilber", 35, 0, 98_304,
                "Social media posting calendar for October to December, including Mark 8 teaser posts.",
                "Social;Campaigns", 2.0);
            Doc(items, "marketing", "Campaigns", "Webinar Registrations.csv", "Alex Wilber", "Alex Wilber", 8, 8, 45_056,
                "Registrations for the Introducing Mark 8 webinar, exported from the event platform.",
                "Events;Leads", 0.8);
            Doc(items, "marketing", "", "Press Release - Mark 8 Announcement.docx", "Megan Bowen", "Patti Fernandez", 10, 2, 128_000,
                "Draft press release announcing the Mark 8, Contoso's next-generation consumer drone, with quotes from Patti Fernandez.",
                "Mark 8;Press;Launch", 3.2);
            Doc(items, "marketing", "", "Customer Personas.docx", "Megan Bowen", "Miriam Graham", 300, 120, 1_572_864,
                "Research-based personas for Contoso's consumer and professional customers.",
                "Research;Customers", 1.5);

            // Finance
            Doc(items, "finance", "", "FY27 Annual Budget.xlsx", "Pradeep Gupta", "Nestor Wilke", 60, 2, 1_048_576,
                "Company-wide FY27 operating budget by department, with headcount, capital expenditure and revenue assumptions.",
                "Budget;FY27", 3.8, label: "Confidential");
            Doc(items, "finance", "", "Q1 FY27 Financial Results.pptx", "Pradeep Gupta", "Patti Fernandez", 15, 3, 2_883_584,
                "Q1 FY27 financial results for the town hall: revenue, margin, operating expenses and outlook.",
                "Results;FY27", 2.9);
            Doc(items, "finance", "", "Budget Forecast Model.xlsm", "Pradeep Gupta", "Pradeep Gupta", 75, 9, 1_835_008,
                "Rolling budget forecast model with macros for scenario planning across business units.",
                "Budget;Forecast", 2.3);
            Doc(items, "finance", "", "Vendor Payment Schedule.xlsx", "Pradeep Gupta", "Pradeep Gupta", 30, 4, 167_936,
                "Payment schedule and terms for Contoso's top 40 vendors, including Fabrikam and Northwind Traders.",
                "Vendors;Payments", 1.4);
            Doc(items, "finance", "", "Expense Policy.pdf", "Pradeep Gupta", "Nestor Wilke", 400, 150, 512_000,
                "Contoso expense policy: travel, meals, accommodation, approvals and reimbursement timelines.",
                "Policy;Expenses", 2.7);
            Doc(items, "finance", "", "Travel Expense Report.xltx", "Pradeep Gupta", "Pradeep Gupta", 500, 200, 61_440,
                "Template for submitting travel expenses for reimbursement.",
                "Template;Expenses", 1.2);
            Doc(items, "finance", "", "Internal Audit Findings 2026.docx", "Nestor Wilke", "Nestor Wilke", 25, 11, 344_064,
                "Findings and recommendations from the 2026 internal audit of procurement and expense processes.",
                "Audit", 1.3, label: "Highly Confidential");

            // Human Resources
            Folder(items, "hr", "Policies", "Diego Siciliani", 400, 30);
            Doc(items, "hr", "Policies", "Employee Handbook 2026.docx", "Diego Siciliani", "Diego Siciliani", 280, 30, 2_097_152,
                "The Contoso employee handbook: code of conduct, working hours, time off, remote work and workplace safety.",
                "Handbook;Policy", 3.3);
            Doc(items, "hr", "Policies", "Remote Work Policy.pdf", "Diego Siciliani", "Patti Fernandez", 220, 70, 420_864,
                "Guidelines for hybrid and remote work, equipment allowances and expectations for collaboration.",
                "Policy", 2.1);
            Doc(items, "hr", "", "Benefits Overview 2027.pdf", "Diego Siciliani", "Diego Siciliani", 20, 3, 1_310_720,
                "Overview of 2027 benefits: medical, dental and vision plans, retirement matching and wellness programs.",
                "Benefits", 2.9);
            Doc(items, "hr", "", "New Hire Onboarding Checklist.docx", "Diego Siciliani", "Lynne Robbins", 210, 6, 88_064,
                "Onboarding checklist for managers and new hires: accounts, equipment, training and first-week meetings.",
                "Onboarding", 2.5);
            Doc(items, "hr", "", "Open Positions.xlsx", "Diego Siciliani", "Diego Siciliani", 50, 1, 73_728,
                "Open roles across Contoso, including four Mark 8 engineering positions.",
                "Recruiting", 1.8);
            Doc(items, "hr", "", "Performance Review Guide.pptx", "Diego Siciliani", "Diego Siciliani", 150, 20, 3_145_728,
                "How the annual performance review works: goals, self-assessment, calibration and feedback conversations.",
                "Performance", 1.7);

            // Legal
            Doc(items, "legal", "", "Mutual NDA Template.dotx", "Joni Sherman", "Joni Sherman", 600, 90, 52_224,
                "Standard mutual non-disclosure agreement template for partners and suppliers.",
                "Template;NDA", 1.9, retention: "Contract");
            Doc(items, "legal", "", "Supplier Agreement - Fabrikam.docx", "Joni Sherman", "Lee Gu", 40, 3, 297_984,
                "Master supply agreement with Fabrikam for Mark 8 motors and propellers, including warranty and liability terms.",
                "Contracts;Suppliers;Mark 8", 2.0, label: "Confidential", retention: "Contract");
            Doc(items, "legal", "", "Patent Filing - Obstacle Avoidance.pdf", "Joni Sherman", "Johanna Lorenz", 120, 33, 2_621_440,
                "Patent application for the Mark 8 obstacle avoidance system using stereo vision and ultrasonic sensors.",
                "Patents;Mark 8", 1.5, label: "Highly Confidential");
            Doc(items, "legal", "", "Contract Register.xlsx", "Joni Sherman", "Joni Sherman", 300, 2, 143_360,
                "Register of active contracts with renewal dates, owners and obligations.",
                "Contracts", 1.6);

            // Retail
            Doc(items, "retail", "", "Store Sales Dashboard.xlsx", "Adele Vance", "Adele Vance", 90, 0, 2_359_296,
                "Weekly store sales by region and product line, with year-over-year comparisons and Mark 8 pre-orders.",
                "Sales;Dashboard", 3.0);
            Doc(items, "retail", "", "Holiday Season Planogram.pptx", "Christie Cline", "Adele Vance", 28, 6, 7_340_032,
                "Planogram for holiday displays, with endcap placement for the Mark 8 and accessories.",
                "Merchandising;Holiday", 1.8);
            Doc(items, "retail", "", "Inventory Forecast.xlsx", "Christie Cline", "Lynne Robbins", 60, 3, 655_360,
                "Inventory forecast for Q2 FY27 by SKU and store cluster, including Mark 8 launch allocation.",
                "Inventory;Forecast", 2.2);
            Doc(items, "retail", "", "Store Opening Checklist.docx", "Adele Vance", "Adele Vance", 120, 15, 110_592,
                "Daily store opening and closing checklist for store managers.",
                "Operations", 1.0);

            // IT Help Desk
            Doc(items, "it", "", "VPN Setup Guide.docx", "Allan Deyoung", "Allan Deyoung", 330, 40, 1_101_824,
                "Step-by-step guide to installing and connecting to the Contoso VPN on Windows and macOS.",
                "How-to;VPN", 2.4);
            Doc(items, "it", "", "Device Refresh Plan.xlsx", "Allan Deyoung", "Nestor Wilke", 50, 8, 204_800,
                "Laptop and phone refresh schedule for FY27 by department, with budget and vendor quotes.",
                "Devices;Budget;FY27", 1.7);
            Doc(items, "it", "", "Security Awareness Training.pptx", "Allan Deyoung", "Allan Deyoung", 140, 22, 5_242_880,
                "Annual security awareness training: phishing, passwords, multifactor authentication and data handling.",
                "Security;Training", 2.0);
            Doc(items, "it", "", "Network Diagram.vsdx", "Allan Deyoung", "Allan Deyoung", 400, 70, 1_478_656,
                "Network topology for the Redmond campus and branch offices.",
                "Network", 1.1);
            Doc(items, "it", "", "Backup-FileShares.ps1", "Allan Deyoung", "Allan Deyoung", 90, 12, 6_144,
                "PowerShell script that backs up departmental file shares to Azure storage each night.",
                "Scripts", 0.7);

            // Engineering
            Doc(items, "eng", "", "Firmware Roadmap.pptx", "Johanna Lorenz", "Lidia Holloway", 100, 10, 3_932_160,
                "Firmware roadmap for FY27: Mark 8 flight modes, over-the-air updates and the companion app SDK.",
                "Roadmap;Firmware;Mark 8", 2.5);
            Doc(items, "eng", "", "Release Notes 4.2.md", "Henrietta Mueller", "Henrietta Mueller", 7, 1, 18_432,
                "Release notes for companion app 4.2: Mark 8 pairing, offline maps and bug fixes.",
                "Release Notes", 1.3);
            Doc(items, "eng", "", "API Design Guidelines.docx", "Henrietta Mueller", "Irvin Sayers", 260, 17, 403_456,
                "Guidelines for designing Contoso REST APIs: naming, versioning, paging, errors and authentication.",
                "Guidelines;API", 1.6);
            Doc(items, "eng", "", "build-pipeline.json", "Henrietta Mueller", "Henrietta Mueller", 80, 9, 9_216,
                "Build pipeline configuration for the firmware and companion app.",
                "Build", 0.5);

            // Pages and news
            Page(items, "intranet", "Home.aspx", "Contoso Intranet", "Patti Fernandez", "Megan Bowen", 700, 1,
                "Welcome to the Contoso intranet. Find news, events, policies and tools for your work.", 3.5);
            Page(items, "intranet", "The-Mark-8-is-coming.aspx", "The Mark 8 is coming: meet our next-generation drone", "Patti Fernandez", "Megan Bowen", 6, 6,
                "After two years of development, the Mark 8 launches in January. Here's what makes it our most capable drone yet.", 3.9);
            Page(items, "intranet", "Welcome-to-FY27.aspx", "Welcome to FY27: a message from Patti", "Patti Fernandez", "Patti Fernandez", 16, 16,
                "Our priorities for FY27: launching the Mark 8, growing retail, and investing in our people.", 2.7);
            Page(items, "hr", "Benefits-Open-Enrollment.aspx", "Open enrollment for 2027 benefits starts November 1", "Diego Siciliani", "Diego Siciliani", 3, 3,
                "Review your 2027 benefits choices between November 1 and November 15. Here's what's new this year.", 2.8);
            Page(items, "it", "Reset-Your-Password.aspx", "How to reset your password", "Allan Deyoung", "Allan Deyoung", 200, 40,
                "Use self-service password reset to change a forgotten or expired password without calling the help desk.", 2.3);

            // List items
            ListItem(items, "intranet", "Events", "STS_ListItem_Events", 4, "Mark 8 Launch Event - Seattle", "Megan Bowen", "Megan Bowen", 25, 2,
                "Join us at the Seattle Center on January 14 for the public launch of the Mark 8.", 2.0);
            ListItem(items, "intranet", "Events", "STS_ListItem_Events", 7, "Town Hall: Q1 FY27 Results", "Patti Fernandez", "Patti Fernandez", 12, 4,
                "Quarterly all-hands with Patti Fernandez and the leadership team.", 1.6);
            ListItem(items, "intranet", "Announcements", "STS_ListItem_Announcements", 31, "Offices closed for Thanksgiving", "Nestor Wilke", "Nestor Wilke", 2, 2,
                "All Contoso offices will be closed on November 26 and 27.", 1.2);
            ListItem(items, "intranet", "Quick Links", "STS_ListItem_Links", 3, "Benefits portal", "Diego Siciliani", "Diego Siciliani", 500, 20,
                "Enroll in or change your benefits.", 0.9);
            ListItem(items, "it", "Help Desk Tickets", "STS_ListItem_GenericList", 1042, "Docking station not detected after update", "Isaiah Langer", "Allan Deyoung", 3, 0,
                "After the latest driver update, my docking station's monitors and network are no longer detected.", 0.8);
            ListItem(items, "it", "Help Desk Tickets", "STS_ListItem_GenericList", 1047, "Request: Visio licence for the Mark 8 team", "Johanna Lorenz", "Allan Deyoung", 1, 1,
                "Three engineers on the Mark 8 project need Visio Plan 2 for architecture diagrams.", 0.8);
            ListItem(items, "hr", "Contacts", "STS_ListItem_Contacts", 8, "Benefits help line", "Diego Siciliani", "Diego Siciliani", 300, 60,
                "Contact the benefits team at extension 4400 or benefits@contoso.com.", 0.7);

            return items;
        }

        private static Site SiteOf(string key) => s_sites.First(s => s.Key == key);

        private static Person PersonOf(string name) => People.First(p => p.Name == name);

        private static DemoItem SiteItem(Site site)
        {
            var url = Tenant + site.Url;
            var item = Base(site, url, site.Title, "", "STS_Site", "Patti Fernandez", "Patti Fernandez", 900, 0, null, site.Description, 2.5 + site.Title.Length % 5 * 0.3);
            item.Managed["IsDocument"] = "false";
            item.Managed["IsContainer"] = "false";
            item.Managed["SiteDescription"] = site.Description;
            item.Managed["IsHubSite"] = site.Key == "intranet" ? "true" : "false";
            item.Managed["ParentLink"] = Tenant;
            item.Managed["Description"] = site.Description;
            return item;
        }


        private static void Doc(
            List<DemoItem> items, string siteKey, string folder, string file, string authors, string editor, int createdDaysAgo,
            int modifiedDaysAgo, long size, string summary, string tags, double popularity, string? label = null, string? retention = null)
        {
            var site = SiteOf(siteKey);
            var web = Tenant + site.Url;
            var folderUrl = $"{web}/Shared Documents" + (folder.Length > 0 ? "/" + folder : "");
            var ext = Path.GetExtension(file).TrimStart('.').ToLowerInvariant();
            var title = Path.GetFileNameWithoutExtension(file);
            var item = Base(site, $"{folderUrl}/{file}", title, ext, "STS_ListItem_DocumentLibrary", authors, editor, createdDaysAgo, modifiedDaysAgo, size, summary, popularity);
            var m = item.Managed;
            var uid = m["UniqueId"];
            m["Filename"] = file;
            m["FileType"] = ext;
            m["FileExtension"] = ext;
            m["SecondaryFileExtension"] = ext;
            m["IsDocument"] = "true";
            m["IsContainer"] = "false";
            m["IsListItem"] = "false";
            m["ParentLink"] = folder.Length > 0 ? folderUrl : $"{web}/Shared Documents/Forms/AllItems.aspx";
            m["ListUrl"] = $"{web}/Shared Documents";
            m["ContentType"] = "Document";
            m["ContentTypeId"] = "0x0101008C2F0F6C27E2B04CB10E8E8F4F2A4E8B";
            m["Tags"] = tags;
            m["owstaxIdMetadataAllTagsInfo"] = string.Join(";", tags.Split(';').Select(t => $"L0|#0{G("tag" + t)}|{t}"));
            m["SensitivityLabel"] = label == null ? "" : G("label" + label).ToString();
            m["InformationProtectionLabelId"] = label == null ? "" : G("label" + label).ToString();
            m["ComplianceTag"] = retention ?? "";
            m["LastSharedByUser"] = label == null ? PersonOf(editor).Name : "";
            m["SharedWithUsersOWSUSER"] = label == null ? $"{PersonOf(editor).Email} | {PersonOf(editor).Name}" : "";
            m["CheckoutUserOWSUSER"] = "";
            m["ServerRedirectedURL"] = $"{web}/_layouts/15/Doc.aspx?sourcedoc=%7B{uid}%7D&file={Uri.EscapeDataString(file)}&action=default&mobileredirect=true";
            m["ServerRedirectedEmbedURL"] = $"{web}/_layouts/15/Doc.aspx?sourcedoc=%7B{uid}%7D&file={Uri.EscapeDataString(file)}&action=interactivepreview";
            m["ServerRedirectedPreviewURL"] = $"{web}/_layouts/15/getpreview.ashx?guidSite={m["SiteId"]}&guidWeb={m["WebId"]}&guidFile={uid}";
            m["PictureThumbnailURL"] = $"{Tenant}/_api/v2.1/drives/b!{m["ListId"][..8]}/items/{m["DocId"]}/thumbnails/0/c400x99999/content";
            m["DefaultEncodingURL"] = $"{folderUrl}/{file}";
            if (ext is "mp4") m["MediaDuration"] = "00:03:42";

            var c = item.Crawled;
            c.AddRange(new[]
            {
                "ows_FileLeafRef", "ows_FileRef", "ows_File_x0020_Type", "ows_File_x0020_Size", "ows_CheckoutUser", "ows__UIVersionString",
                "ows_DocIcon", "ows_ProgId", "ows_ServerRedirectedEmbedUri", "ows_ParentVersionString", "ows_ContentTypeId", "ows_MediaServiceTags",
            }.Select(n => $"{SetSharePoint}:{n}"));
            if (tags.Length > 0) c.Add($"{SetSharePoint}:ows_TaxKeyword");
            if (retention != null) c.Add($"{SetSharePoint}:ows__ComplianceTag");
            if (label != null) c.Add($"{SetSharePoint}:ows__IpLabelId");
            if (ext is "docx" or "xlsx" or "xlsm" or "pptx" or "dotx" or "xltx" or "vsdx" or "one" or "mpp")
                c.AddRange(new[] { "DocSubject", "DocComments", "DocLastAuthor", "DocRevNumber", "DocAppName", "DocLastSaved", "DocCompany" }.Select(n => $"{SetOffice}:{n}"));
            if (ext is "docx" or "pdf" or "pptx" or "xlsx")
                c.AddRange(new[] { "PageCount", "DocumentLanguage", "DocTitle", "DocAuthor", "DocKeywords" }.Select(n => $"{SetParser}:{n}"));
            c.AddRange(new[] { "MetadataExtractorTitle", "MetadataExtractorAuthor", "MetadataExtractorDate" }.Select(n => $"{SetExtractor}:{n}"));
            items.Add(item);
        }

        private static void Folder(List<DemoItem> items, string siteKey, string name, string author, int createdDaysAgo, int modifiedDaysAgo)
        {
            var site = SiteOf(siteKey);
            var web = Tenant + site.Url;
            var item = Base(site, $"{web}/Shared Documents/{name}", name, "", "STS_ListItem_DocumentLibrary", author, author, createdDaysAgo, modifiedDaysAgo, null,
                $"Folder in {site.Title} › Documents.", 1.0);
            item.Managed["Filename"] = name;
            item.Managed["IsDocument"] = "false";
            item.Managed["IsContainer"] = "true";
            item.Managed["ContentType"] = "Folder";
            item.Managed["ContentTypeId"] = "0x0120002A4DDA2B4F2E8D4B8CF26E9F3D0B5B1E";
            item.Managed["ParentLink"] = $"{web}/Shared Documents/Forms/AllItems.aspx";
            item.Crawled.Add($"{SetSharePoint}:ows_FileLeafRef");
            item.Crawled.Add($"{SetSharePoint}:ows_ItemChildCount");
            item.Crawled.Add($"{SetSharePoint}:ows_FolderChildCount");
            items.Add(item);
        }

        private static void Library(List<DemoItem> items, string siteKey, string url, string title, int createdDaysAgo, int modifiedDaysAgo)
        {
            var site = SiteOf(siteKey);
            var web = Tenant + site.Url;
            var item = Base(site, $"{web}/{url}/Forms/AllItems.aspx", title, "", "STS_List_DocumentLibrary", "Patti Fernandez", "Patti Fernandez",
                createdDaysAgo, modifiedDaysAgo, null, $"The {title} library of {site.Title}.", 1.3);
            item.Managed["IsDocument"] = "false";
            item.Managed["IsContainer"] = "false";
            item.Managed["ParentLink"] = web;
            items.Add(item);
        }

        private static void Page(List<DemoItem> items, string siteKey, string file, string title, string author, string editor, int createdDaysAgo, int modifiedDaysAgo,
            string summary, double popularity)
        {
            var site = SiteOf(siteKey);
            var web = Tenant + site.Url;
            var item = Base(site, $"{web}/SitePages/{file}", title, "aspx", "STS_ListItem_WebPageLibrary", author, editor, createdDaysAgo, modifiedDaysAgo,
                Math.Abs(G(file).GetHashCode() % 60_000) + 20_000, summary, popularity);
            var m = item.Managed;
            m["Filename"] = file;
            m["FileType"] = "aspx";
            m["FileExtension"] = "aspx";
            m["SecondaryFileExtension"] = "aspx";
            m["IsDocument"] = "true";
            m["IsContainer"] = "false";
            m["ContentType"] = file.StartsWith("Home", StringComparison.Ordinal) ? "Site Page" : "News Post";
            m["ContentTypeId"] = "0x0101009D1CB255DA76424F860D91F20E6C411800";
            m["ParentLink"] = $"{web}/SitePages/Forms/ByAuthor.aspx";
            m["ListUrl"] = $"{web}/SitePages";
            m["FirstPublishedDate"] = m["Created"];
            m["PromotedState"] = m["ContentType"] == "News Post" ? "2" : "0";
            m["Description"] = summary;
            item.Crawled.AddRange(new[] { "ows_CanvasContent1", "ows_BannerImageUrl", "ows_PromotedState", "ows_FirstPublishedDate", "ows_LayoutWebpartsContent", "ows_FileLeafRef" }
                .Select(n => $"{SetSharePoint}:{n}"));
            items.Add(item);
        }

        private static void ListItem(List<DemoItem> items, string siteKey, string list, string contentClass, int id, string title, string author, string editor,
            int createdDaysAgo, int modifiedDaysAgo, string summary, double popularity)
        {
            var site = SiteOf(siteKey);
            var web = Tenant + site.Url;
            var item = Base(site, $"{web}/Lists/{list.Replace(" ", "")}/DispForm.aspx?ID={id}", title, "", contentClass, author, editor, createdDaysAgo, modifiedDaysAgo,
                null, summary, popularity, listItemId: id, listName: list);
            var m = item.Managed;
            m["IsDocument"] = "false";
            m["IsContainer"] = "false";
            m["IsListItem"] = "true";
            m["ContentType"] = contentClass switch
            {
                "STS_ListItem_Tasks" => "Task",
                "STS_ListItem_Events" => "Event",
                "STS_ListItem_Announcements" => "Announcement",
                "STS_ListItem_Links" => "Link",
                "STS_ListItem_Contacts" => "Contact",
                _ => "Item",
            };
            m["ParentLink"] = $"{web}/Lists/{list.Replace(" ", "")}/AllItems.aspx";
            m["ListUrl"] = $"{web}/Lists/{list.Replace(" ", "")}";
            m["Description"] = summary;
            item.Crawled.AddRange(new[] { "ows_Body", "ows_LinkTitle", "ows_ID", "ows_Attachments" }.Select(n => $"{SetSharePoint}:{n}"));
            if (contentClass == "STS_ListItem_Tasks")
                item.Crawled.AddRange(new[] { "ows_AssignedTo", "ows_Status", "ows_Priority", "ows_DueDate", "ows_PercentComplete" }.Select(n => $"{SetSharePoint}:{n}"));
            if (contentClass == "STS_ListItem_Events")
                item.Crawled.AddRange(new[] { "ows_EventDate", "ows_EndDate", "ows_Location", "ows_fAllDayEvent" }.Select(n => $"{SetSharePoint}:{n}"));
            items.Add(item);
        }


        /// <summary>The properties every result has.</summary>
        private static DemoItem Base(
            Site site, string path, string title, string fileType, string contentClass, string authors, string editor, int createdDaysAgo,
            int modifiedDaysAgo, long? size, string summary, double popularity, int? listItemId = null, string? listName = null)
        {
            int n = s_nextId++;
            var web = Tenant + site.Url;
            var now = DateTime.UtcNow;
            // Fixed times of day per item, so the list doesn't look generated.
            var created = now.Date.AddDays(-createdDaysAgo).AddMinutes(8 * 60 + (n * 47) % 540);
            var modified = modifiedDaysAgo == 0
                ? now.AddMinutes(-(15 + n * 13 % 300))
                : now.Date.AddDays(-modifiedDaysAgo).AddMinutes(9 * 60 + (n * 71) % 480);
            if (created > modified) created = modified.AddHours(-3);

            var authorPeople = authors.Split(';').Select(PersonOf).ToList();
            var editorPerson = PersonOf(editor);
            var uid = G(path);
            var siteId = G("site" + site.Key);
            var webId = G("web" + site.Key);
            var listId = G("list" + site.Key + (listName ?? (contentClass == "STS_ListItem_WebPageLibrary" ? "SitePages" : "Documents")));
            long docId = 17_592_186_044_416 + n * 3_217L;
            var workId = docId.ToString(CultureInfo.InvariantCulture);

            var item = new DemoItem { Summary = summary, Popularity = popularity };
            var m = item.Managed;
            m["Title"] = title;
            m["Path"] = path;
            m["OriginalPath"] = path;
            m["ContentClass"] = contentClass;
            m["LastModifiedTime"] = Iso(modified);
            m["Write"] = Iso(modified);
            m["Created"] = Iso(created);
            m["Author"] = string.Join(";", authorPeople.Select(p => p.Name));
            m["DisplayAuthor"] = string.Join(";", authorPeople.Select(p => p.Name));
            m["AuthorOWSUSER"] = string.Join("\n", authorPeople.Select(User));
            m["CreatedBy"] = authorPeople[0].Name;
            m["ModifiedBy"] = editorPerson.Name;
            m["EditorOWSUSER"] = User(editorPerson);
            if (size is { } bytes) m["Size"] = bytes.ToString(CultureInfo.InvariantCulture);
            m["SiteName"] = web;
            m["SiteTitle"] = site.Title;
            m["SPSiteUrl"] = web;
            m["SPWebUrl"] = web;
            m["SiteId"] = siteId.ToString();
            m["WebId"] = webId.ToString();
            m["ListId"] = listId.ToString();
            m["NormSiteID"] = siteId.ToString();
            m["NormWebID"] = webId.ToString();
            m["NormListID"] = listId.ToString();
            m["UniqueId"] = "{" + uid.ToString().ToUpperInvariant() + "}";
            m["NormUniqueID"] = uid.ToString();
            m["ListItemID"] = (listItemId ?? n).ToString(CultureInfo.InvariantCulture);
            m["DocId"] = workId;
            m["WorkId"] = workId;
            m["DocumentSignature"] = $"{(uint)G("sig" + path).GetHashCode()};{(uint)G("sig2" + path).GetHashCode()}";
            m["ContentSource"] = "Local SharePoint";
            m["SiteTemplate"] = site.Template.Split('#')[0];
            m["SiteTemplateId"] = site.Template.StartsWith("GROUP", StringComparison.Ordinal) ? "64" : "68";
            m["WebTemplate"] = site.Template.Split('#')[0];
            m["DepartmentId"] = "{" + G("hub").ToString() + "}";
            m["RelatedHubSites"] = G("hub").ToString();
            if (site.Template.StartsWith("GROUP", StringComparison.Ordinal)) m["GroupId"] = G("group" + site.Key).ToString();
            m["language"] = "en";
            m["DetectedLanguage"] = "en";
            m["Culture"] = "en-US";
            m["UrlZone"] = "0";
            m["PartitionId"] = "a0c1d5e5-2b4f-4ad8-9b0e-7d4c8f1e2a3b";
            m["ResultTypeId"] = "0";
            m["RenderTemplateId"] = fileType switch
            {
                "docx" or "dotx" => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Word.js",
                "xlsx" or "xlsm" or "xltx" => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Excel.js",
                "pptx" => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_PowerPoint.js",
                "pdf" => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_PDF.js",
                "mp4" => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Video.js",
                _ when contentClass == "STS_Site" => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Site.js",
                _ => "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js",
            };
            m["piSearchResultId"] = $"{G("pi" + path):N}";
            m["IsExternalContent"] = "false";
            m["GeoLocationSource"] = "NAM";
            int views = (int)(popularity * 173) + n * 7 % 90;
            m["ViewsLifeTime"] = views.ToString(CultureInfo.InvariantCulture);
            m["ViewsLifeTimeUniqueUsers"] = (views / 3 + 1).ToString(CultureInfo.InvariantCulture);
            m["ViewsRecent"] = (views / 9).ToString(CultureInfo.InvariantCulture);
            m["ViewsRecentUniqueUsers"] = (views / 20).ToString(CultureInfo.InvariantCulture);
            m["Description"] = "";
            m["importance"] = "0";

            var c = item.Crawled;
            c.AddRange(new[] { "ows_Title", "ows_Modified", "ows_Created", "ows_Author", "ows_Editor", "ows_ContentType", "ows_UniqueId", "ows_FSObjType", "ows_GUID", "ows_ID", "ows_MetaInfo" }
                .Select(p => $"{SetSharePoint}:{p}"));
            c.AddRange(new[] { "ows_r_DAT_Created", "ows_r_DAT_Modified", "ows_r_USER_Author", "ows_r_USER_Editor" }.Select(p => $"{SetSharePointRaw}:{p}"));
            c.AddRange(new[] { "Site", "SiteID", "WebID", "ListID", "spsiteurl", "spweburl" }.Select(p => $"{SetSharePointSite}:{p}"));
            c.AddRange(new[] { "FileName", "FileExtension", "FileSize", "FileWriteTime", "FileCreateTime", "Attrib" }.Select(p => $"{SetBasic}:{p}"));
            c.AddRange(new[] { "url", "urlzone", "HostName", "UrlDepth" }.Select(p => $"{SetInternal}:{p}"));
            return item;
        }

        private static string User(Person p) => $"{p.Email} | {p.Name} | i:0#.f|membership|{p.Email}";

        private static string Iso(DateTime utc) => utc.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);

        /// <summary>A stable GUID for a seed string, so IDs stay the same between launches.</summary>
        private static Guid G(string seed) => new(MD5.HashData(Encoding.UTF8.GetBytes(seed)));
    }


    #endregion
}
