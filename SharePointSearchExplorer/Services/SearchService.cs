using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Services;

/// <summary>
/// Calls the SharePoint Search REST API (<c>/_api/search/query</c> and <c>/_api/search/postquery</c>).
/// </summary>
public sealed class SearchService
{
    /// <summary>Properties requested for every result in addition to the user's extras.</summary>
    public static readonly string[] DefaultSelectProperties =
    [
        "Title", "Path", "Filename", "FileType", "FileExtension", "SecondaryFileExtension", "ContentClass",
        "IsDocument", "IsContainer", "LastModifiedTime", "Created", "Author", "ModifiedBy", "Size",
        "SiteName", "SiteTitle", "SPWebUrl", "HitHighlightedSummary", "WorkId", "DocId", "Rank",
        "UniqueId", "ListItemID", "ParentLink",
    ];

    private const string ManagedPropertiesHelp =
        "SharePoint did not return any managed property names for this item. The ManagedProperties refiner " +
        "is only populated for items indexed after it was enabled; re-indexing the site or library usually fixes this.";

    // Property sets used to label crawled properties (same mapping as the original tool).
    private static readonly Dictionary<string, string> s_propertySets = new(StringComparer.OrdinalIgnoreCase)
    {
        ["00020329-0000-0000-C000-000000000046"] = "SharePoint",
        ["00130329-0000-0130-c000-000000131346"] = "SharePoint",
        ["00140329-0000-0140-c000-000000141446"] = "SharePoint",
        ["158d7563-aeff-4dbf-bf16-4a1445f0366c"] = "SharePoint",
        ["ED280121-B677-4E2A-8FBC-0D9E2325B0A2"] = "SharePoint",
        ["0b63e343-9ccc-11d0-bcdb-00805fccce04"] = "Basic",
        ["49691c90-7e17-101a-a91c-08002b2ecda9"] = "Basic",
        ["B725F130-47EF-101A-A5F1-02608C9EEBAC"] = "Basic",
        ["012357bd-1113-171d-1f25-292bb0b0b0b0"] = "Internal",
        ["F29F85E0-4FF9-1068-AB91-08002B27B3D9"] = "Office",
        ["64ae120f-487d-445a-8d5a-5258f99cb970"] = "Document Parser",
        ["e835446c-937b-4492-95f3-d89988b01039"] = "Metadata Extractor",
    };

    private static readonly string[] s_excludedManagedProperties =
        ["ClassificationLastScan", "ClassificationConfidence", "ClassificationCount", "ClassificationContext"];

    private const string UserAgent = "NONISV|SharePointSearchExplorer/1.0";

    private readonly AuthService _auth;
    private readonly HttpClient _http = CreateClient(windowsAuth: false);
    private readonly HttpClient _windowsHttp = CreateClient(windowsAuth: true);

    public SearchService(AuthService auth) => _auth = auth;

    /// <summary>Runs a query with an HTTP GET, like pasting the REST URL in a browser.</summary>
    public async Task<SearchResponse> QueryAsync(ConnectionProfile profile, SearchQuery query, CancellationToken ct)
    {
        var url = BuildGetUrl(profile.SiteUrl, query);
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        return ParseResponse(await SendAsync(profile, request, ct, query));
    }

    /// <summary>
    /// Sends a request URL built on the Advanced Query Rules page as it is. The URL has to be on the profile's
    /// site, because the profile's access token goes with it.
    /// </summary>
    public async Task<HttpExchange> GetAsync(ConnectionProfile profile, string url, CancellationToken ct, SearchQuery? query = null)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
            !Uri.TryCreate(profile.SiteUrl, UriKind.Absolute, out var site) ||
            uri.Scheme != site.Scheme || !uri.Host.Equals(site.Host, StringComparison.OrdinalIgnoreCase))
        {
            throw new SearchException($"The request URL has to start with the profile's site, {profile.SiteUrl.TrimEnd('/')}/.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Get, uri);
        request.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        return await SendAsync(profile, request, ct, query);
    }

    /// <summary>The suggestions in a <c>/_api/search/suggest</c> response: queries, then people, then personal results.</summary>
    public static IReadOnlyList<string> ParseSuggestions(HttpExchange http)
    {
        using var doc = JsonDocument.Parse(http.ResponseBody);
        var root = Unwrap(doc.RootElement);
        root = Prop(root, "suggest") ?? root;

        var suggestions = new List<string>();
        if (Prop(root, "Queries") is { ValueKind: JsonValueKind.Array } queries)
        {
            foreach (var q in queries.EnumerateArray())
            {
                if (Prop(q, "Query")?.GetString() is { Length: > 0 } text) suggestions.Add(text);
            }
        }
        if (Prop(root, "PeopleNames") is { ValueKind: JsonValueKind.Array } people)
        {
            foreach (var p in people.EnumerateArray())
            {
                if (p.ValueKind == JsonValueKind.String && p.GetString() is { Length: > 0 } name) suggestions.Add($"{name} (person)");
            }
        }
        if (Prop(root, "PersonalResults") is { ValueKind: JsonValueKind.Array } results)
        {
            foreach (var r in results.EnumerateArray())
            {
                var title = Prop(r, "Title")?.GetString();
                var link = Prop(r, "Url")?.GetString();
                if (!string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(link)) suggestions.Add($"{title} ({link})");
            }
        }
        return suggestions;
    }

    /// <summary>Runs a query with an HTTP POST; used when the select list is too long for a URL.</summary>
    public async Task<SearchResponse> PostQueryAsync(ConnectionProfile profile, SearchQuery query, CancellationToken ct)
    {
        var url = CombineUrl(profile.SiteUrl, "_api/search/postquery");
        using var request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Accept.ParseAdd("application/json;odata=verbose");
        request.Content = new StringContent(BuildPostBody(query), Encoding.UTF8);
        request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose;charset=utf-8");

        // Bearer-token calls to SharePoint Online don't need a form digest; Windows auth does.
        if (profile.AuthMode == AuthMode.Windows)
            request.Headers.Add("X-RequestDigest", await GetFormDigestAsync(profile, ct));

        return ParseResponse(await SendAsync(profile, request, ct, query));
    }

    /// <summary>
    /// All managed properties with values for one item: first asks for the names via the
    /// <c>ManagedProperties</c> refiner, then re-queries the item selecting all of them.
    /// </summary>
    public async Task<IReadOnlyList<PropertyEntry>> GetManagedPropertiesAsync(
        ConnectionProfile profile, SearchResultItem item, SearchQuery baseQuery, CancellationToken ct)
    {
        var names = (await GetRefinerNamesAsync(profile, item, baseQuery, "ManagedProperties", ct))
            .Except(s_excludedManagedProperties, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (names.Count == 0) throw new SearchException(ManagedPropertiesHelp);

        var response = await PostQueryAsync(profile, ItemQuery(item, baseQuery, selectProperties: names), ct);
        var row = response.Rows.FirstOrDefault()
            ?? throw new SearchException("The item was not returned when re-queried for its properties.");

        return row.ReturnedProperties
            .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>Crawled property names (and their property set) for one item, via the <c>CrawledProperties</c> refiner.</summary>
    public async Task<IReadOnlyList<PropertyEntry>> GetCrawledPropertiesAsync(
        ConnectionProfile profile, SearchResultItem item, SearchQuery baseQuery, CancellationToken ct)
    {
        var names = await GetRefinerNamesAsync(profile, item, baseQuery, "CrawledProperties", ct);
        if (names.Count == 0)
            throw new SearchException("SharePoint did not return any crawled property names for this item.");

        return names
            .Select(ToCrawledProperty)
            .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private async Task<List<string>> GetRefinerNamesAsync(
        ConnectionProfile profile, SearchResultItem item, SearchQuery baseQuery, string refiner, CancellationToken ct)
    {
        var query = ItemQuery(item, baseQuery, refiners: $"{refiner}(filter=5000/0/*)");
        var response = await PostQueryAsync(profile, query, ct);
        return response.Refiners.TryGetValue(refiner, out var entries)
            ? entries.Select(e => e.Name).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct().ToList()
            : [];
    }

    /// <summary>A query that returns just this item, pinned by its WorkId.</summary>
    private static SearchQuery ItemQuery(
        SearchResultItem item, SearchQuery baseQuery, string? refiners = null, IReadOnlyList<string>? selectProperties = null)
    {
        if (string.IsNullOrEmpty(item.WorkId))
            throw new SearchException("This result has no WorkId, so its properties can't be looked up.");

        return new SearchQuery
        {
            QueryText = baseQuery.QueryText,
            QueryTemplate = baseQuery.QueryTemplate,
            SourceId = baseQuery.SourceId,
            HiddenConstraints = $"WorkId:\"{item.WorkId}\"",
            RowLimit = 1,
            TrimDuplicates = false,
            Refiners = refiners,
            SelectProperties = selectProperties ?? [],
        };
    }

    private static PropertyEntry ToCrawledProperty(string raw)
    {
        // Refiner values look like "00130329-0000-0130-c000-000000131346:ows_Title".
        int colon = raw.IndexOf(':');
        if (colon > 0 && Guid.TryParse(raw[..colon], out _))
        {
            var set = raw[..colon];
            var category = s_propertySets.TryGetValue(set, out var name) ? name : "Other";
            return new PropertyEntry(raw[(colon + 1)..], set, category);
        }
        return new PropertyEntry(raw, "", "Other");
    }

    #region HTTP

    /// <summary>Sends the request; a failure throws a <see cref="SearchException"/> that carries the exchange.</summary>
    private async Task<HttpExchange> SendAsync(ConnectionProfile profile, HttpRequestMessage request, CancellationToken ct, SearchQuery? query = null)
    {
        HttpClient client;
        if (profile.AuthMode == AuthMode.Windows)
        {
            client = _windowsHttp;
        }
        else
        {
            var token = await _auth.GetAccessTokenAsync(profile, allowInteractive: true, ct: ct);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            client = _http;
        }

        var requestBody = request.Content == null ? "" : await request.Content.ReadAsStringAsync(ct);

        HttpResponseMessage response;
        try
        {
            response = await client.SendAsync(request, ct);
        }
        catch (HttpRequestException ex)
        {
            var error = $"Could not reach {new Uri(profile.SiteUrl).Host}: {ex.Message}";
            throw new SearchException(error, http: Exchange(request, requestBody, query, error: error));
        }

        using (response)
        {
            var body = await response.Content.ReadAsStringAsync(ct);
            var http = Exchange(request, requestBody, query, response, body);
            if (response.IsSuccessStatusCode)
            {
                profile.IsSignedIn = true;
                return http;
            }
            if (response.StatusCode == HttpStatusCode.Unauthorized) profile.IsSignedIn = false;

            var message = ExtractError(body) ?? response.ReasonPhrase ?? "Request failed";
            var hint = response.StatusCode switch
            {
                HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden =>
                    " Check that the account has access to the site and that the app registration has the SharePoint AllSites.Read (delegated) permission.",
                HttpStatusCode.TooManyRequests or HttpStatusCode.ServiceUnavailable =>
                    " SharePoint is throttling requests; wait a moment and try again.",
                _ => "",
            };
            throw new SearchException($"{(int)response.StatusCode} {response.StatusCode}: {message}{hint}", (int)response.StatusCode, http);
        }
    }

    /// <summary>Records a request (after sending, so it includes the client's default headers) and its response.</summary>
    private static HttpExchange Exchange(
        HttpRequestMessage request, string requestBody, SearchQuery? query, HttpResponseMessage? response = null, string responseBody = "", string? error = null)
    {
        return new HttpExchange
        {
            Method = request.Method.Method,
            Url = request.RequestUri?.ToString() ?? "",
            RequestHeaders = Headers(request.Headers, request.Content?.Headers),
            RequestBody = requestBody,
            StatusCode = response == null ? null : (int)response.StatusCode,
            ReasonPhrase = response?.ReasonPhrase ?? "",
            ProtocolVersion = response?.Version.ToString() ?? "",
            ResponseHeaders = response == null ? [] : Headers(response.Headers, response.Content.Headers),
            ResponseBody = responseBody,
            Error = error,
            Query = query,
        };

        static List<KeyValuePair<string, string>> Headers(HttpHeaders headers, HttpHeaders? contentHeaders)
        {
            var all = contentHeaders == null ? headers : headers.Concat(contentHeaders);
            return all.Select(h => new KeyValuePair<string, string>(h.Key, IsSecret(h.Key) ? Redact(h.Value) : string.Join(", ", h.Value))).ToList();
        }

        // The details can be copied and shared, so the access token and form digest are left out.
        static bool IsSecret(string header) =>
            header.Equals("Authorization", StringComparison.OrdinalIgnoreCase) || header.Equals("X-RequestDigest", StringComparison.OrdinalIgnoreCase);

        static string Redact(IEnumerable<string> values)
        {
            var value = values.FirstOrDefault() ?? "";
            int space = value.IndexOf(' ');
            return space > 0 ? $"{value[..space]} [redacted]" : "[redacted]";
        }
    }

    private async Task<string> GetFormDigestAsync(ConnectionProfile profile, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, CombineUrl(profile.SiteUrl, "_api/contextinfo"));
        request.Headers.Accept.ParseAdd("application/json;odata=verbose");
        request.Content = new StringContent("");
        var json = (await SendAsync(profile, request, ct)).ResponseBody;

        using var doc = JsonDocument.Parse(json);
        var root = Unwrap(doc.RootElement);
        var info = Prop(root, "GetContextWebInformation") ?? root;
        return Prop(info, "FormDigestValue")?.GetString() ?? "";
    }

    private static HttpClient CreateClient(bool windowsAuth)
    {
        var handler = new SocketsHttpHandler
        {
            AutomaticDecompression = DecompressionMethods.All,
            PooledConnectionLifetime = TimeSpan.FromMinutes(5),
        };
        if (windowsAuth)
        {
            handler.Credentials = CredentialCache.DefaultCredentials;
            handler.PreAuthenticate = true;
        }

        var client = new HttpClient(DemoMode.IsEnabled ? DemoMode.CreateHandler() : handler) { Timeout = TimeSpan.FromSeconds(100) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd(UserAgent);
        return client;
    }

    private static string CombineUrl(string siteUrl, string relative) => siteUrl.TrimEnd('/') + "/" + relative;

    #endregion

    #region Request building

    private static string BuildGetUrl(string siteUrl, SearchQuery q)
    {
        var parameters = new List<string>
        {
            "querytext=" + Quote(q.QueryText),
            "startrow=" + q.StartRow,
            "rowlimit=" + q.RowLimit,
            "trimduplicates=" + (q.TrimDuplicates ? "true" : "false"),
            "clienttype=" + Quote("ContentSearchRegular"),
        };
        if (q.SelectProperties.Count > 0) parameters.Add("selectproperties=" + Quote(string.Join(",", q.SelectProperties)));
        if (!string.IsNullOrWhiteSpace(q.Refiners)) parameters.Add("refiners=" + Quote(q.Refiners));
        if (!string.IsNullOrWhiteSpace(q.RefinementFilters)) parameters.Add("refinementfilters=" + Quote(q.RefinementFilters));
        if (!string.IsNullOrWhiteSpace(q.HiddenConstraints)) parameters.Add("hiddenconstraints=" + Quote(q.HiddenConstraints));
        if (!string.IsNullOrWhiteSpace(q.SortList)) parameters.Add("sortlist=" + Quote(q.SortList));
        if (!string.IsNullOrWhiteSpace(q.QueryTemplate)) parameters.Add("querytemplate=" + Quote(q.QueryTemplate));
        if (!string.IsNullOrWhiteSpace(q.SourceId)) parameters.Add("sourceid=" + Quote(q.SourceId.Trim()));

        return CombineUrl(siteUrl, "_api/search/query?" + string.Join("&", parameters));

        // KQL/REST string values are single-quoted with embedded quotes doubled, then URL-encoded.
        static string Quote(string value) => Uri.EscapeDataString("'" + value.Replace("'", "''") + "'");
    }

    private static string BuildPostBody(SearchQuery q)
    {
        var request = new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "Microsoft.Office.Server.Search.REST.SearchRequest" },
            ["Querytext"] = q.QueryText,
            ["StartRow"] = q.StartRow,
            ["RowLimit"] = q.RowLimit,
            ["TrimDuplicates"] = q.TrimDuplicates,
            ["ClientType"] = "ContentSearchRegular",
        };
        if (q.SelectProperties.Count > 0)
            request["SelectProperties"] = new JsonObject { ["results"] = new JsonArray(q.SelectProperties.Select(p => (JsonNode)p!).ToArray()) };
        if (!string.IsNullOrWhiteSpace(q.Refiners)) request["Refiners"] = q.Refiners;
        if (!string.IsNullOrWhiteSpace(q.HiddenConstraints)) request["HiddenConstraints"] = q.HiddenConstraints;
        if (!string.IsNullOrWhiteSpace(q.QueryTemplate)) request["QueryTemplate"] = q.QueryTemplate;
        if (!string.IsNullOrWhiteSpace(q.SourceId)) request["SourceId"] = q.SourceId.Trim();
        if (!string.IsNullOrWhiteSpace(q.RefinementFilters))
            request["RefinementFilters"] = new JsonObject { ["results"] = new JsonArray(q.RefinementFilters) };

        return new JsonObject { ["request"] = request }.ToJsonString();
    }

    #endregion

    #region Response parsing

    /// <summary>Parses both odata=nometadata and odata=verbose response shapes.</summary>
    public static SearchResponse ParseResponse(HttpExchange http)
    {
        using var doc = JsonDocument.Parse(http.ResponseBody);
        var root = Unwrap(doc.RootElement);
        root = Prop(root, "query") ?? Prop(root, "postquery") ?? root;

        var primary = Prop(root, "PrimaryQueryResult");
        var relevant = primary is { } p ? Prop(p, "RelevantResults") : null;

        var rows = new List<SearchResultItem>();
        long total = 0;
        if (relevant is { } rel)
        {
            total = ReadLong(Prop(rel, "TotalRows"));
            var table = Prop(rel, "Table");
            if (table is { } t && Prop(t, "Rows") is { ValueKind: JsonValueKind.Array } rowArray)
            {
                foreach (var row in rowArray.EnumerateArray())
                {
                    if (Prop(row, "Cells") is not { ValueKind: JsonValueKind.Array } cells) continue;
                    rows.Add(new SearchResultItem(cells.EnumerateArray().Select(c =>
                        new KeyValuePair<string, string?>(
                            Prop(c, "Key")?.GetString() ?? "",
                            Prop(c, "Value") is { ValueKind: JsonValueKind.String } v ? v.GetString() : null))));
                }
            }
        }

        var refiners = new Dictionary<string, IReadOnlyList<RefinerEntry>>(StringComparer.OrdinalIgnoreCase);
        if (primary is { } pr && Prop(pr, "RefinementResults") is { } rr && Prop(rr, "Refiners") is { ValueKind: JsonValueKind.Array } refArray)
        {
            foreach (var refiner in refArray.EnumerateArray())
            {
                var name = Prop(refiner, "Name")?.GetString() ?? "";
                var entries = new List<RefinerEntry>();
                if (Prop(refiner, "Entries") is { ValueKind: JsonValueKind.Array } entryArray)
                {
                    foreach (var e in entryArray.EnumerateArray())
                    {
                        entries.Add(new RefinerEntry(
                            Prop(e, "RefinementName")?.GetString() ?? Prop(e, "RefinementValue")?.GetString() ?? "",
                            Prop(e, "RefinementToken")?.GetString() ?? "",
                            ReadLong(Prop(e, "RefinementCount"))));
                    }
                }
                refiners[name] = entries;
            }
        }

        return new SearchResponse
        {
            TotalRows = total,
            ElapsedMilliseconds = (int)ReadLong(Prop(root, "ElapsedTime")),
            Rows = rows,
            Refiners = refiners,
            Http = http,
        };
    }

    private static string? ExtractError(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;
            var error = Prop(root, "odata.error") ?? Prop(root, "error");
            if (error is not { } e) return null;
            var message = Prop(e, "message");
            return message is { ValueKind: JsonValueKind.Object } m ? Prop(m, "value")?.GetString() : message?.GetString();
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>Strips the verbose "d" wrapper.</summary>
    internal static JsonElement Unwrap(JsonElement root) => Prop(root, "d") ?? root;

    /// <summary>Case-insensitive property lookup that also unwraps verbose { "results": [...] } collections.</summary>
    internal static JsonElement? Prop(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object) return null;
        foreach (var p in element.EnumerateObject())
        {
            if (!string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)) continue;
            var value = p.Value;
            if (value.ValueKind == JsonValueKind.Object && value.TryGetProperty("results", out var results) && results.ValueKind == JsonValueKind.Array)
                return results;
            return value;
        }
        return null;
    }

    private static long ReadLong(JsonElement? element) => element switch
    {
        { ValueKind: JsonValueKind.Number } n => n.TryGetInt64(out var l) ? l : (long)n.GetDouble(),
        { ValueKind: JsonValueKind.String } s => long.TryParse(s.GetString(), out var l) ? l : 0,
        _ => 0,
    };

    #endregion
}
