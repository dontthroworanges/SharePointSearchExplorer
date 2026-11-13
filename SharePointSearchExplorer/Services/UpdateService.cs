using System.ComponentModel;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SharePointSearchExplorer.Services;

/// <summary>A newer release on GitHub, downloaded and ready to install.</summary>
public sealed record UpdateInfo(Version Version, string TagName, string ReleaseUrl, string? Notes, string DownloadedFile);

/// <summary>
/// Checks the project's GitHub releases for a newer version (by the release's tag, e.g. "v0.2.0"), downloads its .exe
/// in the background, and installs it by swapping it in for the running .exe and restarting.
/// </summary>
/// <remarks>
/// Windows won't let a running .exe be overwritten, but it can be renamed. So installing renames the running file to
/// "&lt;name&gt;.exe.old", copies the new one into its place and starts it; the next start deletes the .old file.
/// If the app's folder isn't writable (e.g. Program Files), the rename and copy run elevated, after a UAC prompt.
/// </remarks>
public static class UpdateService
{
    public const string Owner = "dontthroworanges";
    public const string Repo = "SharePointSearchExplorer";
    public const string RepoUrl = $"https://github.com/{Owner}/{Repo}";
    public const string ReleasesUrl = $"{RepoUrl}/releases";
    private const string LatestReleaseApi = $"https://api.github.com/repos/{Owner}/{Repo}/releases/latest";

    private const string OldSuffix = ".old";

    /// <summary>Downloads are kept here until installed (and cleared on the next start).</summary>
    public static string DownloadFolder { get; } = Path.Combine(SettingsStore.DataFolder, "Updates");

    /// <summary>This app's version (the csproj &lt;Version&gt;), as compared with release tags.</summary>
    public static Version CurrentVersion { get; } = Normalize(Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0));

    /// <summary>The running .exe (the single-file bundle when published, not a temporary extraction).</summary>
    public static string? ExecutablePath => Environment.ProcessPath;

    private static readonly HttpClient s_http = CreateClient();

    private static readonly JsonSerializerOptions s_json = new() { PropertyNameCaseInsensitive = true };

    private static readonly object s_lock = new();
    private static Task<UpdateInfo?>? s_inFlight;

    /// <summary>
    /// Raised when <see cref="State"/>, <see cref="AvailableVersion"/> or <see cref="DownloadPercent"/> changes.
    /// May be raised on a background thread.
    /// </summary>
    public static event EventHandler? StateChanged;

    public static UpdateState State { get; private set; } = UpdateState.Idle;

    /// <summary>The newer version found by the last check, while it downloads and once it's ready.</summary>
    public static Version? AvailableVersion { get; private set; }

    /// <summary>0–100 while downloading.</summary>
    public static int DownloadPercent { get; private set; }

    /// <summary>The downloaded update, ready to install (kept for the rest of the session).</summary>
    public static UpdateInfo? ReadyUpdate { get; private set; }

    /// <summary>Why the last check failed, when <see cref="State"/> is <see cref="UpdateState.Failed"/>.</summary>
    public static string? LastError { get; private set; }

    private static HttpClient CreateClient()
    {
        var client = new HttpClient { Timeout = Timeout.InfiniteTimeSpan }; // downloads are large; requests set their own timeouts
        // GitHub's API requires a User-Agent.
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue(Repo, CurrentVersion.ToString(3)));
        return client;
    }

    private static void SetState(UpdateState state)
    {
        State = state;
        StateChanged?.Invoke(null, EventArgs.Empty);
    }

    /// <summary>
    /// Finds the latest release. Returns null when this version is current (or newer). When there's a newer one,
    /// downloads its .exe in the background, verifies it against the digest GitHub publishes, and returns it.
    /// A call while a check is already running joins that check rather than starting another download.
    /// </summary>
    /// <exception cref="UpdateException">The release couldn't be read or downloaded.</exception>
    public static Task<UpdateInfo?> CheckAndDownloadAsync()
    {
        lock (s_lock)
        {
            if (s_inFlight is { IsCompleted: false } running) return running;
            return s_inFlight = Task.Run(CheckAndDownloadCoreAsync);
        }
    }

    private static async Task<UpdateInfo?> CheckAndDownloadCoreAsync()
    {
        LastError = null;
        DownloadPercent = 0;
        SetState(UpdateState.Checking);
        try
        {
            var release = await GetLatestReleaseAsync(CancellationToken.None);
            if (release == null || !TryParseTag(release.TagName, out var latest) || latest <= CurrentVersion)
            {
                SetState(UpdateState.UpToDate);
                return null;
            }

            AvailableVersion = latest;
            if (ReadyUpdate is { } done && done.Version == latest && File.Exists(done.DownloadedFile))
            {
                SetState(UpdateState.Ready);
                return done;
            }

            var asset = PickAsset(release.Assets)
                ?? throw new UpdateException($"Version {latest.ToString(3)} is available, but its release has no .exe to download. Get it from GitHub instead.");

            SetState(UpdateState.Downloading);
            Action<double> progress = fraction =>
            {
                int percent = (int)(fraction * 100);
                if (percent == DownloadPercent) return;
                DownloadPercent = percent;
                StateChanged?.Invoke(null, EventArgs.Empty);
            };
            var file = await DownloadAsync(asset, latest, progress, CancellationToken.None);
            ReadyUpdate = new UpdateInfo(latest, release.TagName, release.HtmlUrl ?? ReleasesUrl, release.Body, file);
            SetState(UpdateState.Ready);
            return ReadyUpdate;
        }
        catch (Exception ex)
        {
            LastError = ex is UpdateException ? ex.Message : $"Couldn't check for updates: {ex.Message}";
            SetState(UpdateState.Failed);
            throw ex as UpdateException ?? new UpdateException(LastError, ex);
        }
    }

    private static async Task<GitHubRelease?> GetLatestReleaseAsync(CancellationToken cancel)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseApi);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancel);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        try
        {
            using var response = await s_http.SendAsync(request, timeout.Token);
            // 404: no published releases yet.
            if (response.StatusCode == System.Net.HttpStatusCode.NotFound) return null;
            if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                throw new UpdateException("GitHub is limiting requests from this network. Try again later.");
            response.EnsureSuccessStatusCode();
            await using var stream = await response.Content.ReadAsStreamAsync(timeout.Token);
            return await JsonSerializer.DeserializeAsync<GitHubRelease>(stream, s_json, timeout.Token);
        }
        catch (OperationCanceledException) when (!cancel.IsCancellationRequested)
        {
            throw new UpdateException("GitHub didn't respond. Check your internet connection and try again.");
        }
        catch (HttpRequestException ex)
        {
            throw new UpdateException($"Couldn't reach GitHub: {ex.Message}", ex);
        }
        catch (JsonException ex)
        {
            throw new UpdateException("GitHub's response couldn't be read.", ex);
        }
    }

    /// <summary>
    /// The release's .exe for this machine: one named for this architecture (e.g. "…-arm64.exe") if there is one,
    /// otherwise one named like this app, otherwise the only .exe.
    /// </summary>
    private static GitHubAsset? PickAsset(List<GitHubAsset>? assets)
    {
        var exes = (assets ?? []).Where(a => a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) && a.BrowserDownloadUrl != null).ToList();
        if (exes.Count == 0) return null;

        string arch = RuntimeInformation.ProcessArchitecture.ToString(); // X64, X86, Arm64
        string exeName = Path.GetFileName(ExecutablePath) ?? "SharePointSearchExplorer.exe";
        return exes.FirstOrDefault(a => a.Name.Contains(arch, StringComparison.OrdinalIgnoreCase))
            ?? exes.FirstOrDefault(a => string.Equals(a.Name, exeName, StringComparison.OrdinalIgnoreCase))
            ?? (exes.Count == 1 ? exes[0] : exes.FirstOrDefault(a => a.Name.StartsWith(Repo, StringComparison.OrdinalIgnoreCase)));
    }

    private static async Task<string> DownloadAsync(GitHubAsset asset, Version version, Action<double>? progress, CancellationToken cancel)
    {
        var folder = Path.Combine(DownloadFolder, version.ToString(3));
        Directory.CreateDirectory(folder);
        var target = Path.Combine(folder, asset.Name);
        var partial = target + ".download";

        try
        {
            using var response = await s_http.GetAsync(asset.BrowserDownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancel);
            response.EnsureSuccessStatusCode();
            long? total = response.Content.Headers.ContentLength ?? (asset.Size > 0 ? asset.Size : null);

            using var sha = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            await using (var source = await response.Content.ReadAsStreamAsync(cancel))
            await using (var output = new FileStream(partial, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            {
                var buffer = new byte[81920];
                long read = 0;
                int n;
                while ((n = await source.ReadAsync(buffer, cancel)) > 0)
                {
                    await output.WriteAsync(buffer.AsMemory(0, n), cancel);
                    sha.AppendData(buffer, 0, n);
                    read += n;
                    if (total > 0) progress?.Invoke((double)read / total.Value);
                }
            }

            // GitHub publishes "sha256:<hex>" for each asset; refuse a download that doesn't match.
            var expected = asset.Digest is { } d && d.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase) ? d[7..] : null;
            var actual = Convert.ToHexString(sha.GetHashAndReset());
            if (expected != null && !string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                throw new UpdateException("The downloaded update didn't match the release's checksum, so it wasn't kept. Try again later.");

            File.Move(partial, target, overwrite: true);
            return target;
        }
        catch (HttpRequestException ex)
        {
            throw new UpdateException($"Couldn't download the update: {ex.Message}", ex);
        }
        catch (IOException ex)
        {
            throw new UpdateException($"Couldn't save the update: {ex.Message}", ex);
        }
        finally
        {
            TryDelete(partial);
        }
    }

    /// <summary>
    /// Replaces the running .exe with the downloaded one, in the same folder, and starts it. The caller then closes
    /// this instance. The new file takes the release's file name (normally SharePointSearchExplorer.exe, the same as
    /// the running one; a copy with another name, e.g. a renamed download, is replaced by the release's name).
    /// Returns false if the user declined the UAC prompt needed for a protected folder.
    /// </summary>
    /// <exception cref="UpdateException">The files couldn't be swapped.</exception>
    public static bool Install(UpdateInfo update)
    {
        var exe = ExecutablePath ?? throw new UpdateException("Couldn't find where the app is running from.");
        if (!File.Exists(update.DownloadedFile)) throw new UpdateException("The downloaded update is missing. Check for updates again.");
        var old = exe + OldSuffix;
        var target = Path.Combine(Path.GetDirectoryName(exe)!, Path.GetFileName(update.DownloadedFile));

        try
        {
            TryDelete(old);
            // A running .exe can't be overwritten or deleted, but it can be renamed.
            File.Move(exe, old);
            try
            {
                File.Copy(update.DownloadedFile, target, overwrite: true);
            }
            catch
            {
                File.Move(old, exe); // put the running version back
                throw;
            }
        }
        catch (UnauthorizedAccessException)
        {
            if (!InstallElevated(exe, old, update.DownloadedFile, target)) return false;
        }
        catch (IOException ex)
        {
            throw new UpdateException($"Couldn't replace {Path.GetFileName(exe)}: {ex.Message}", ex);
        }

        Process.Start(new ProcessStartInfo(target) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(target) });
        return true;
    }

    /// <summary>Does the rename and copy from an elevated PowerShell, for an app in a folder only admins can write to.</summary>
    private static bool InstallElevated(string exe, string old, string downloaded, string target)
    {
        static string Quote(string path) => "'" + path.Replace("'", "''") + "'";
        var script =
            "$ErrorActionPreference = 'Stop'; " +
            $"if (Test-Path -LiteralPath {Quote(old)}) {{ Remove-Item -LiteralPath {Quote(old)} -Force }}; " +
            $"Move-Item -LiteralPath {Quote(exe)} -Destination {Quote(old)}; " +
            $"try {{ Copy-Item -LiteralPath {Quote(downloaded)} -Destination {Quote(target)} -Force }} " +
            $"catch {{ Move-Item -LiteralPath {Quote(old)} -Destination {Quote(exe)}; exit 1 }}";
        var encoded = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(script));

        try
        {
            using var process = Process.Start(new ProcessStartInfo("powershell.exe",
                $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand {encoded}")
            {
                UseShellExecute = true,
                Verb = "runas",
                WindowStyle = ProcessWindowStyle.Hidden,
            }) ?? throw new UpdateException("Couldn't start the installer.");
            process.WaitForExit();
            if (process.ExitCode != 0 || !File.Exists(target))
                throw new UpdateException($"Couldn't replace {Path.GetFileName(exe)} in {Path.GetDirectoryName(exe)}.");
            return true;
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223) // ERROR_CANCELLED: UAC prompt declined
        {
            return false;
        }
    }

    /// <summary>At startup: removes the previous version left by an update, and downloads no longer needed.</summary>
    public static void CleanUpAfterUpdate() => _ = Task.Run(async () =>
    {
        // The previous version was renamed to "<its name>.exe.old" (its name may differ from this one's), so look
        // for any. Right after an update it may still be closing, so it can't be deleted straight away.
        if (Path.GetDirectoryName(ExecutablePath) is { Length: > 0 } folder)
        {
            for (int attempt = 0; attempt < 10; attempt++)
            {
                string[] leftovers;
                try
                {
                    leftovers = Directory.GetFiles(folder, $"{Repo}*.exe{OldSuffix}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Could not look for the previous version: {ex.Message}");
                    break;
                }
                if (leftovers.Length == 0) break;
                if (attempt > 0) await Task.Delay(TimeSpan.FromSeconds(1));
                foreach (var old in leftovers) TryDelete(old);
            }
        }
        try
        {
            if (Directory.Exists(DownloadFolder))
            {
                foreach (var dir in Directory.GetDirectories(DownloadFolder))
                {
                    // Keep a download for a version newer than this one (not installed yet).
                    if (Version.TryParse(Path.GetFileName(dir), out var v) && Normalize(v) > CurrentVersion) continue;
                    Directory.Delete(dir, recursive: true);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not clean up old updates: {ex.Message}");
        }
    });

    /// <summary>"v1.2.3", "1.2", "V1.2.3-beta" → 1.2.3 (anything after a '-' or '+' is ignored).</summary>
    public static bool TryParseTag(string? tag, out Version version)
    {
        version = new Version(0, 0, 0);
        if (string.IsNullOrWhiteSpace(tag)) return false;
        var text = tag.Trim().TrimStart('v', 'V');
        int cut = text.IndexOfAny(['-', '+', ' ']);
        if (cut >= 0) text = text[..cut];
        if (!text.Contains('.')) text += ".0";
        if (!Version.TryParse(text, out var parsed)) return false;
        version = Normalize(parsed);
        return true;
    }

    /// <summary>Major.Minor.Build, so 1.2 == 1.2.0 and the assembly's 1.2.0.0 == tag 1.2.0.</summary>
    private static Version Normalize(Version v) => new(v.Major, Math.Max(v.Minor, 0), Math.Max(v.Build, 0));

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not delete {path}: {ex.Message}");
        }
    }

    private sealed class GitHubRelease
    {
        [JsonPropertyName("tag_name")] public string TagName { get; set; } = "";
        [JsonPropertyName("html_url")] public string? HtmlUrl { get; set; }
        [JsonPropertyName("body")] public string? Body { get; set; }
        [JsonPropertyName("assets")] public List<GitHubAsset>? Assets { get; set; }
    }

    private sealed class GitHubAsset
    {
        [JsonPropertyName("name")] public string Name { get; set; } = "";
        [JsonPropertyName("browser_download_url")] public string? BrowserDownloadUrl { get; set; }
        [JsonPropertyName("size")] public long Size { get; set; }
        [JsonPropertyName("digest")] public string? Digest { get; set; }
    }
}

public enum UpdateState
{
    /// <summary>Not checked this session.</summary>
    Idle,
    Checking,
    UpToDate,
    Downloading,
    /// <summary>Downloaded and waiting to be installed (<see cref="UpdateService.ReadyUpdate"/>).</summary>
    Ready,
    Failed,
}

/// <summary>An update check or install failed; the message is shown to the user.</summary>
public sealed class UpdateException(string message, Exception? inner = null) : Exception(message, inner);
