using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Services;

/// <summary>
/// Text for the tabs of the Server Response page (the original tool's Status, Raw and Debug tabs).
/// </summary>
internal static class HttpDetailsFormatter
{
    // Relaxed escaping keeps markup such as <c0> in HitHighlightedSummary, and non-ASCII text, readable.
    private static readonly JsonSerializerOptions s_indented = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>"HTTP Details": the request and status lines, then the query statistics SharePoint returned.</summary>
    public static string Details(HttpExchange http)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{http.Method} {http.Url}");
        sb.AppendLine(StatusLine(http));

        if (!http.IsSuccess)
        {
            if (http.ResponseBody.Length > 0) sb.AppendLine().AppendLine(Pretty(http.ResponseBody));
            return sb.ToString().TrimEnd();
        }

        try
        {
            using var doc = JsonDocument.Parse(http.ResponseBody);
            AppendQueryDetails(sb, doc.RootElement);
        }
        catch (JsonException)
        {
            // Not JSON: the status line is all there is to show.
        }
        return sb.ToString().TrimEnd();
    }

    /// <summary>"Full API Response": request and response headers, then the response body (indented JSON).</summary>
    public static string Response(HttpExchange http)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{http.Method} {http.Url}");

        sb.AppendLine().AppendLine("Request:");
        foreach (var header in http.RequestHeaders) sb.AppendLine($"\t{header.Key}: {header.Value}");
        if (http.RequestBody.Length > 0) sb.AppendLine().AppendLine(Pretty(http.RequestBody));

        sb.AppendLine().AppendLine("Response:");
        sb.AppendLine("\t" + StatusLine(http));
        foreach (var header in http.ResponseHeaders) sb.AppendLine($"\t{header.Key}: {header.Value}");
        if (http.ResponseBody.Length > 0) sb.AppendLine().AppendLine(Pretty(http.ResponseBody));

        return sb.ToString().TrimEnd();
    }

    // Response headers that identify the request in SharePoint's logs or describe how the server handled it.
    private static readonly string[] s_diagnosticHeaders =
    [
        "SPRequestGuid", "request-id", "X-MS-Correlation-Id", "SPRequestDuration", "SPIisLatency",
        "X-SharePointHealthScore", "MicrosoftSharePointTeamServices", "X-MSEdge-Ref", "Date",
    ];

    /// <summary>
    /// "Debug": how the query was sent and how SharePoint processed it. The original tool's Debug tab listed
    /// similar fields from the server's ULS logs; this one uses what the REST request and response contain.
    /// </summary>
    public static string Debug(HttpExchange http)
    {
        var sb = new StringBuilder();

        if (http.Query is { } q)
        {
            sb.AppendLine($"Query: {q.QueryText}");
            sb.AppendLine($"Template: {q.QueryTemplate}");
            sb.AppendLine($"Hidden Constraint: {q.HiddenConstraints}");
            sb.AppendLine($"Source Id: {q.SourceId}");
            sb.AppendLine($"Refiners: {q.Refiners}");
            sb.AppendLine($"Refinement Filters: {q.RefinementFilters}");
            sb.AppendLine($"Sort List: {q.SortList}");
            sb.AppendLine($"Start Row: {q.StartRow}");
            sb.AppendLine($"Row Limit: {q.RowLimit}");
            sb.AppendLine($"Trim Duplicates: {(q.TrimDuplicates ? "true" : "false")}");
            sb.AppendLine($"Select Properties: {string.Join(", ", q.SelectProperties)}");
            sb.AppendLine();
        }

        sb.AppendLine(StatusLine(http));

        if (http.IsSuccess)
        {
            try
            {
                using var doc = JsonDocument.Parse(http.ResponseBody);
                AppendDebugDetails(sb, doc.RootElement);
            }
            catch (JsonException)
            {
                // Not JSON: only the request and the headers below can be shown.
            }
        }

        var headers = http.ResponseHeaders
            .Where(h => s_diagnosticHeaders.Contains(h.Key, StringComparer.OrdinalIgnoreCase))
            .ToList();
        if (headers.Count > 0)
        {
            sb.AppendLine().AppendLine("Diagnostic Response Headers:");
            foreach (var header in headers) sb.AppendLine($"\t{header.Key}: {header.Value}");
        }

        return sb.ToString().TrimEnd();
    }

    private static void AppendDebugDetails(StringBuilder sb, JsonElement document)
    {
        JsonElement? root = SearchService.Unwrap(document);
        root = Prop(root, "query") ?? Prop(root, "postquery") ?? root;

        var primary = Prop(root, "PrimaryQueryResult");
        var relevant = Prop(primary, "RelevantResults");
        var properties = Prop(root, "Properties");
        var relevantProperties = Prop(relevant, "Properties");

        sb.AppendLine();
        sb.AppendLine($"Query Modification: {KeyValue(relevantProperties, "QueryModification")}");
        sb.AppendLine($"Serialized Query: {KeyValue(properties, "SerializedQuery")}");
        sb.AppendLine($"Correlation Id: {KeyValue(properties, "CorrelationId")}");
        sb.AppendLine($"Elapsed Time (ms): {Text(Prop(root, "ElapsedTime"))}");
        sb.AppendLine($"Relevant Results: {Text(Prop(relevant, "TotalRows"))} ({Text(Prop(relevant, "RowCount"))} returned)");
        var refiners = Prop(Prop(primary, "RefinementResults"), "Refiners");
        sb.AppendLine($"Refiner Results: {(refiners is { ValueKind: JsonValueKind.Array } r ? r.GetArrayLength() : 0)}");
        sb.AppendLine($"Query Suggestion: {Text(Prop(root, "SpellingSuggestion"))}");

        if (Prop(root, "TriggeredRules") is { ValueKind: JsonValueKind.Array } rules && rules.GetArrayLength() > 0)
        {
            sb.AppendLine().AppendLine("Triggered Rules:");
            foreach (var rule in rules.EnumerateArray()) sb.AppendLine($"\tQuery Rule Id: {Text(rule)}");
        }

        AppendProperties(sb, "Query Properties:", properties);
        AppendProperties(sb, "Relevant Results Properties:", relevantProperties);
    }

    /// <summary>Every entry of a SharePoint { Key, Value, ValueType } list, one per line.</summary>
    private static void AppendProperties(StringBuilder sb, string title, JsonElement? properties)
    {
        if (properties is not { ValueKind: JsonValueKind.Array } list || list.GetArrayLength() == 0) return;
        sb.AppendLine().AppendLine(title);
        foreach (var entry in list.EnumerateArray())
            sb.AppendLine($"\t{Text(Prop(entry, "Key"))}: {Text(Prop(entry, "Value"))}");
    }

    private static string StatusLine(HttpExchange http) => http.StatusCode is { } code
        ? $"HTTP/{http.ProtocolVersion} {code} {http.ReasonPhrase}".TrimEnd()
        : $"No response: {http.Error}";

    private static void AppendQueryDetails(StringBuilder sb, JsonElement document)
    {
        JsonElement? root = SearchService.Unwrap(document);
        root = Prop(root, "query") ?? Prop(root, "postquery") ?? root;

        if (Text(Prop(root, "ElapsedTime")) is { Length: > 0 } elapsed)
            sb.AppendLine().AppendLine($"Elapsed Time (ms): {elapsed}");

        if (KeyValue(Prop(root, "Properties"), "SerializedQuery") is { Length: > 0 } serialized)
            sb.AppendLine().AppendLine("Serialized Query:").AppendLine("\t" + serialized);

        if (Prop(root, "TriggeredRules") is { ValueKind: JsonValueKind.Array } rules && rules.GetArrayLength() > 0)
        {
            sb.AppendLine().AppendLine("Triggered Rules:");
            foreach (var rule in rules.EnumerateArray()) sb.AppendLine($"\tQuery Rule Id: {Text(rule)}");
        }

        if (Prop(root, "PrimaryQueryResult") is { ValueKind: JsonValueKind.Object } primary)
        {
            sb.AppendLine().AppendLine("Primary Query Results:");
            AppendResult(sb, primary);
        }

        if (Prop(root, "SecondaryQueryResults") is { ValueKind: JsonValueKind.Array } secondary)
        {
            foreach (var result in secondary.EnumerateArray().Where(r => r.ValueKind == JsonValueKind.Object))
            {
                sb.AppendLine().AppendLine("Secondary Query Results:");
                AppendResult(sb, result);
            }
        }
    }

    private static void AppendResult(StringBuilder sb, JsonElement result)
    {
        var relevant = Prop(result, "RelevantResults");
        sb.AppendLine($"\tTotal Rows: {Text(Prop(relevant, "TotalRows"))}");
        sb.AppendLine($"\tTotal Rows Including Duplicates: {Text(Prop(relevant, "TotalRowsIncludingDuplicates"))}");
        sb.AppendLine($"\tQuery Id: {Text(Prop(result, "QueryId"))}");
        sb.AppendLine($"\tQuery Rule Id: {Text(Prop(result, "QueryRuleId"))}");
        sb.AppendLine($"\tQuery Modification: {KeyValue(Prop(relevant, "Properties"), "QueryModification")}");
    }

    private static JsonElement? Prop(JsonElement? element, string name) =>
        element is { } e ? SearchService.Prop(e, name) : null;

    /// <summary>The Value of the entry with the given Key in a SharePoint { Key, Value, ValueType } list.</summary>
    private static string KeyValue(JsonElement? properties, string key)
    {
        if (properties is not { ValueKind: JsonValueKind.Array } list) return "";
        foreach (var entry in list.EnumerateArray())
        {
            if (string.Equals(Text(Prop(entry, "Key")), key, StringComparison.OrdinalIgnoreCase))
                return Text(Prop(entry, "Value"));
        }
        return "";
    }

    private static string Text(JsonElement? element) => element switch
    {
        { ValueKind: JsonValueKind.String } s => s.GetString() ?? "",
        { ValueKind: JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False } v => v.GetRawText(),
        _ => "",
    };

    /// <summary>Indented JSON; anything that isn't JSON (e.g. an HTML error page) is returned unchanged.</summary>
    private static string Pretty(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            return JsonSerializer.Serialize(doc.RootElement, s_indented);
        }
        catch (JsonException)
        {
            return body;
        }
    }
}
