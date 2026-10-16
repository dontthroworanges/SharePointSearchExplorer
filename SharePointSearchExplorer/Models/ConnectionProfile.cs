using System.Text.Json.Serialization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSearchExplorer.Models;

[JsonConverter(typeof(JsonStringEnumConverter<AuthMode>))]
public enum AuthMode
{
    /// <summary>Microsoft Entra ID sign-in through the system browser (SharePoint Online).</summary>
    Browser,

    /// <summary>Current Windows user via Negotiate/NTLM (SharePoint Server on-premises).</summary>
    Windows,
}

/// <summary>
/// A saved connection: which SharePoint site to query and how to sign in to it.
/// Switching profiles switches site, tenant and/or signed-in account.
/// </summary>
public partial class ConnectionProfile : ObservableObject
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    [ObservableProperty]
    private string _name = "";

    [ObservableProperty]
    private string _siteUrl = "";

    [ObservableProperty]
    private AuthMode _authMode = AuthMode.Browser;

    /// <summary>Optional tenant ID or domain. Derived from the site URL when empty.</summary>
    [ObservableProperty]
    private string _tenantId = "";

    /// <summary>Microsoft Entra app registration (public client) used to sign in.</summary>
    [ObservableProperty]
    private string _clientId = "";

    /// <summary>MSAL home account identifier of the last account used with this profile.</summary>
    [ObservableProperty]
    private string _accountId = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(AccountDisplay))]
    private string _userName = "";

    /// <summary>True once a token (or Windows credential) has been confirmed this session. Not persisted.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(AccountDisplay))]
    [property: JsonIgnore]
    private bool _isSignedIn;

    [JsonIgnore]
    public string AccountDisplay => AuthMode == AuthMode.Windows
        ? "Current Windows user"
        : IsSignedIn && !string.IsNullOrEmpty(UserName) ? $"Signed in as {UserName}" : "Not signed in";

    [JsonIgnore]
    public string EffectiveName => string.IsNullOrWhiteSpace(Name) ? SiteUrl : Name;
}
