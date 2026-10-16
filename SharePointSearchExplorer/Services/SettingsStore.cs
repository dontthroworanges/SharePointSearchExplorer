using System.Diagnostics;
using System.Text.Json;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Services;

/// <summary>Loads and saves <see cref="AppSettings"/> as JSON under %LOCALAPPDATA%.</summary>
public sealed class SettingsStore
{
    public const int MaxHistory = 25;

    public static string DataFolder { get; } = ResolveDataFolder();

    /// <summary>
    /// %LOCALAPPDATA%\SharePointSearchExplorer. The app was called "SharePoint Search Query Tool" and, before that,
    /// "Simple Search Query Tool", so the newest existing folder of those names (settings, sign-in session, token
    /// cache) is moved across on first run; if it can't be moved, it keeps being used.
    /// </summary>
    private static string ResolveDataFolder()
    {
        if (DemoMode.IsEnabled) return DemoMode.DataFolder;
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var folder = Path.Combine(root, "SharePointSearchExplorer");
        if (Directory.Exists(folder)) return folder;
        var legacy = new[] { "SharePointSearchQueryTool", "SimpleSearchQueryTool" }
            .Select(name => Path.Combine(root, name))
            .FirstOrDefault(Directory.Exists);
        if (legacy == null) return folder;
        try
        {
            Directory.Move(legacy, folder);
            return folder;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Debug.WriteLine($"Could not move {legacy} to {folder}, still using it: {ex.Message}");
            return legacy;
        }
    }

    private static string SettingsPath => Path.Combine(DataFolder, "settings.json");

    private static readonly JsonSerializerOptions s_json = new() { WriteIndented = true };

    private SettingsStore(AppSettings data) => Data = data;

    public AppSettings Data { get; }

    public event EventHandler? CurrentProfileChanged;

    /// <summary>Queries were added to or removed from the history (the query box and Query History page both show it).</summary>
    public event EventHandler? HistoryChanged;

    public ConnectionProfile? CurrentProfile =>
        Data.Profiles.FirstOrDefault(p => p.Id == Data.CurrentProfileId) ?? Data.Profiles.FirstOrDefault();

    public static SettingsStore Load()
    {
        if (DemoMode.IsEnabled) return new SettingsStore(DemoMode.CreateSettings());
        try
        {
            if (File.Exists(SettingsPath))
            {
                var data = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsPath), s_json);
                if (data != null) return new SettingsStore(data);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not read settings, starting fresh: {ex.Message}");
        }
        return new SettingsStore(new AppSettings());
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(DataFolder);
            var temp = SettingsPath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(Data, s_json));
            File.Move(temp, SettingsPath, overwrite: true);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not save settings: {ex.Message}");
        }
    }

    public void SetCurrentProfile(ConnectionProfile profile)
    {
        Data.CurrentProfileId = profile.Id;
        Save();
        CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Adds a new profile or saves changes to an existing one.</summary>
    public void SaveProfile(ConnectionProfile profile)
    {
        if (!Data.Profiles.Contains(profile)) Data.Profiles.Add(profile);
        Save();
        if (profile == CurrentProfile) CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }

    public void RemoveProfile(ConnectionProfile profile)
    {
        bool wasCurrent = profile == CurrentProfile;
        Data.Profiles.Remove(profile);
        if (wasCurrent) Data.CurrentProfileId = Data.Profiles.FirstOrDefault()?.Id;
        Save();
        if (wasCurrent) CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Restores every setting to its default and removes all profiles and history. Values are copied into
    /// the existing <see cref="Data"/> instance because pages bind straight to it.
    /// </summary>
    public void ResetToDefaults()
    {
        var defaults = new AppSettings();
        foreach (var property in typeof(AppSettings).GetProperties().Where(p => p.CanRead && p.CanWrite))
            property.SetValue(Data, property.GetValue(defaults));
        Save();
        CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
        HistoryChanged?.Invoke(this, EventArgs.Empty);
    }

    public void RemoveHistory(string query)
    {
        if (!Data.History.Remove(query)) return;
        Save();
        HistoryChanged?.Invoke(this, EventArgs.Empty);
    }

    public void ClearHistory()
    {
        Data.History.Clear();
        Save();
        HistoryChanged?.Invoke(this, EventArgs.Empty);
    }

    public void AddHistory(string query)
    {
        if (string.IsNullOrWhiteSpace(query)) return;
        Data.History.RemoveAll(h => string.Equals(h, query, StringComparison.OrdinalIgnoreCase));
        Data.History.Insert(0, query);
        if (Data.History.Count > MaxHistory) Data.History.RemoveRange(MaxHistory, Data.History.Count - MaxHistory);
        Save();
        HistoryChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Puts imported queries at the top of the history, in file order. Queries already in the history are skipped,
    /// and the oldest entries are dropped to stay within <see cref="MaxHistory"/>.
    /// </summary>
    public (int Added, int Skipped, int Dropped) ImportHistory(IReadOnlyList<string> queries)
    {
        var added = queries.Where(q => !Data.History.Contains(q, StringComparer.OrdinalIgnoreCase)).ToList();
        Data.History.InsertRange(0, added);
        int dropped = Math.Max(0, Data.History.Count - MaxHistory);
        if (dropped > 0) Data.History.RemoveRange(MaxHistory, dropped);
        if (added.Count > 0)
        {
            Save();
            HistoryChanged?.Invoke(this, EventArgs.Empty);
        }
        return (added.Count, queries.Count - added.Count, dropped);
    }
}
