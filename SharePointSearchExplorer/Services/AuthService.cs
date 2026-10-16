using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensibility;
using Microsoft.Identity.Client.Extensions.Msal;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Services;

/// <summary>Thrown when a token cannot be obtained without showing sign-in UI.</summary>
public sealed class SignInRequiredException(string message) : Exception(message);

/// <summary>
/// Microsoft Entra sign-in for SharePoint Online.
///
/// By default this matches the original tool's "SharePoint Online" method: Microsoft's
/// public SharePoint Online Management Shell client, signed in through an embedded
/// sign-in window (<see cref="WebViewSignIn"/>), so no app registration is needed.
/// A profile may instead name its own app registration, which signs in through the
/// default web browser. Either way tokens are kept in an encrypted (DPAPI) MSAL cache
/// on disk and refreshed silently on later runs.
/// </summary>
public sealed class AuthService
{
    /// <summary>SharePoint Online Management Shell, the client the original tool uses.</summary>
    public const string BuiltInClientId = "9bc3ab49-b65d-410a-85ad-de819febfddc";
    private const string BuiltInRedirectUri = "https://oauth.spops.microsoft.com/";
    private const string CustomAppRedirectUri = "http://localhost";

    private readonly Dictionary<string, IPublicClientApplication> _apps = new();
    private readonly SemaphoreSlim _gate = new(1, 1);
    private MsalCacheHelper? _cacheHelper;

    /// <summary>Raised after a profile's signed-in account changes.</summary>
    public event EventHandler<ConnectionProfile>? AccountChanged;

    /// <summary>"https://contoso.sharepoint.com" for any URL on that host.</summary>
    public static string GetResource(string siteUrl)
    {
        var uri = new Uri(siteUrl);
        return $"{uri.Scheme}://{uri.Authority}";
    }

    /// <summary>
    /// Tenant used for the authority. contoso.sharepoint.com / contoso-my / contoso-admin
    /// map to contoso.onmicrosoft.com; anything else falls back to "organizations".
    /// </summary>
    public static string GetTenant(ConnectionProfile profile)
    {
        if (!string.IsNullOrWhiteSpace(profile.TenantId)) return profile.TenantId.Trim();

        var host = new Uri(profile.SiteUrl).Host;
        const string suffix = ".sharepoint.com";
        if (host.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
        {
            var name = host[..^suffix.Length];
            foreach (var s in new[] { "-my", "-admin", "-df" })
            {
                if (name.EndsWith(s, StringComparison.OrdinalIgnoreCase)) name = name[..^s.Length];
            }
            return $"{name}.onmicrosoft.com";
        }
        return "organizations";
    }

    /// <summary>
    /// Returns an access token for the profile's SharePoint host.
    /// Tries the token cache first; when <paramref name="allowInteractive"/> is set and the
    /// cache can't satisfy the request, opens the default browser to sign in.
    /// </summary>
    /// <param name="chooseAccount">Force the account picker (used for "sign in with a different account").</param>
    public async Task<string> GetAccessTokenAsync(
        ConnectionProfile profile, bool allowInteractive, bool chooseAccount = false, CancellationToken ct = default)
    {
        if (DemoMode.IsEnabled)
        {
            profile.IsSignedIn = true;
            return "demo-access-token";
        }
        await _gate.WaitAsync(ct);
        try
        {
            var app = await GetAppAsync(profile);
            string[] scopes = [GetResource(profile.SiteUrl) + "/.default"];

            if (!chooseAccount)
            {
                var account = await FindAccountAsync(app, profile);
                if (account != null)
                {
                    try
                    {
                        var silent = await app.AcquireTokenSilent(scopes, account).ExecuteAsync(ct);
                        Remember(profile, silent);
                        return silent.AccessToken;
                    }
                    catch (MsalUiRequiredException) when (allowInteractive)
                    {
                        // Fall through to browser sign-in.
                    }
                    catch (MsalUiRequiredException)
                    {
                        profile.IsSignedIn = false;
                        throw new SignInRequiredException("Your sign-in has expired. Search again to sign in.");
                    }
                }
            }

            if (!allowInteractive)
            {
                profile.IsSignedIn = false;
                throw new SignInRequiredException("Not signed in yet. You'll be asked to sign in on the first search.");
            }

            var builder = app.AcquireTokenInteractive(scopes)
                .WithParentActivityOrWindow(App.WindowHandle)
                // NoPrompt lets an existing session complete silently; SelectAccount shows the picker.
                .WithPrompt(chooseAccount ? Prompt.SelectAccount : Prompt.NoPrompt);

            if (UsesBuiltInClient(profile))
            {
                builder = builder.WithCustomWebUi(new WebViewSignIn());
            }
            else
            {
                builder = builder
                    .WithUseEmbeddedWebView(false) // system browser => reuse its existing Microsoft 365 session
                    .WithSystemWebViewOptions(new SystemWebViewOptions
                    {
                        HtmlMessageSuccess = BrowserPage("Signed in", "You can close this tab and return to SharePoint Search Explorer."),
                        HtmlMessageError = BrowserPage("Sign-in failed", "Return to SharePoint Search Explorer for details. Error: {0} {1}"),
                    });
            }

            if (!chooseAccount && !string.IsNullOrEmpty(profile.UserName))
                builder = builder.WithLoginHint(profile.UserName);

            var result = await builder.ExecuteAsync(ct);
            Remember(profile, result);
            return result.AccessToken;
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>Removes the profile's account from the token cache.</summary>
    public async Task SignOutAsync(ConnectionProfile profile)
    {
        if (profile.AuthMode == AuthMode.Browser)
        {
            var app = await GetAppAsync(profile);
            var account = await FindAccountAsync(app, profile);
            if (account != null) await app.RemoveAsync(account);
        }
        profile.AccountId = "";
        profile.UserName = "";
        profile.IsSignedIn = false;
        AccountChanged?.Invoke(this, profile);
    }

    /// <summary>True when the profile has no app ID of its own and uses the original tool's built-in client.</summary>
    public static bool UsesBuiltInClient(ConnectionProfile profile) =>
        string.IsNullOrWhiteSpace(profile.ClientId)
        || string.Equals(profile.ClientId.Trim(), BuiltInClientId, StringComparison.OrdinalIgnoreCase);

    private async Task<IPublicClientApplication> GetAppAsync(ConnectionProfile profile)
    {
        bool builtIn = UsesBuiltInClient(profile);
        var clientId = builtIn ? BuiltInClientId : profile.ClientId.Trim();
        var tenant = GetTenant(profile);
        var key = $"{clientId}|{tenant}".ToLowerInvariant();
        if (_apps.TryGetValue(key, out var existing)) return existing;

        var app = PublicClientApplicationBuilder.Create(clientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, tenant)
            .WithRedirectUri(builtIn ? BuiltInRedirectUri : CustomAppRedirectUri)
            .WithClientName("SharePointSearchExplorer")
            .Build();

        _cacheHelper ??= await MsalCacheHelper.CreateAsync(
            new StorageCreationPropertiesBuilder("msal_token_cache.bin", SettingsStore.DataFolder).Build());
        _cacheHelper.RegisterCache(app.UserTokenCache);

        _apps[key] = app;
        return app;
    }

    private static async Task<IAccount?> FindAccountAsync(IPublicClientApplication app, ConnectionProfile profile)
    {
        if (!string.IsNullOrEmpty(profile.AccountId))
        {
            var byId = await app.GetAccountAsync(profile.AccountId);
            if (byId != null) return byId;
        }

        var accounts = await app.GetAccountsAsync();
        if (!string.IsNullOrEmpty(profile.UserName))
            return accounts.FirstOrDefault(a => string.Equals(a.Username, profile.UserName, StringComparison.OrdinalIgnoreCase));

        return null;
    }

    private void Remember(ConnectionProfile profile, AuthenticationResult result)
    {
        var id = result.Account?.HomeAccountId?.Identifier ?? "";
        var user = result.Account?.Username ?? "";
        profile.IsSignedIn = true;
        if (id == profile.AccountId && user == profile.UserName) return;

        profile.AccountId = id;
        profile.UserName = user;
        AccountChanged?.Invoke(this, profile);
    }

    private static string BrowserPage(string heading, string text) =>
        "<html><head><title>" + heading + "</title></head>" +
        "<body style=\"font-family:'Segoe UI',sans-serif;display:flex;align-items:center;justify-content:center;height:90vh;margin:0\">" +
        "<div style=\"text-align:center\"><h2 style=\"font-weight:600\">" + heading + "</h2><p>" + text + "</p></div></body></html>";
}
