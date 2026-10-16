namespace SharePointSearchExplorer.Helpers;

/// <summary>Delimiters offered in Settings for copying property names and values.</summary>
public static class Delimiters
{
    public sealed record Option(string Key, string Label, string Value);

    public static IReadOnlyList<Option> All { get; } =
    [
        new("Comma", "Comma  ( , )", ","),
        new("Semicolon", "Semicolon  ( ; )", ";"),
        new("Colon", "Colon  ( : )", ":"),
        new("Pipe", "Pipe  ( | )", "|"),
        new("Space", "Single space", " "),
    ];

    public static string Get(string key) =>
        All.FirstOrDefault(o => string.Equals(o.Key, key, StringComparison.OrdinalIgnoreCase))?.Value ?? ",";

    /// <summary>
    /// Joins fields with the delimiter. For comma and semicolon (CSV-style), a field containing the
    /// delimiter, a quote or a line break is quoted so the result still opens correctly in Excel.
    /// </summary>
    public static string Join(string key, params string[] fields)
    {
        var delimiter = Get(key);
        bool csv = delimiter is "," or ";";
        return string.Join(delimiter, fields.Select(f => csv && NeedsQuotes(f, delimiter) ? Quote(f) : f));

        static bool NeedsQuotes(string f, string d) => f.Contains(d) || f.Contains('"') || f.Contains('\n') || f.Contains('\r');
        static string Quote(string f) => "\"" + f.Replace("\"", "\"\"") + "\"";
    }
}
