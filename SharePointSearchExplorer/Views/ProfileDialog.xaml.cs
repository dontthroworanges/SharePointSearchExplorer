using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Views;

/// <summary>Adds a new connection profile or edits an existing one.</summary>
public sealed partial class ProfileDialog : ContentDialog
{
    private readonly ConnectionProfile? _existing;
    private bool _nameEdited;

    /// <summary>Last name filled in from the URL; any other text in the name box means the user typed it.</summary>
    private string _suggestedName = "";

    public ProfileDialog(ConnectionProfile? existing, string defaultClientId)
    {
        InitializeComponent();
        _existing = existing;

        Title = existing == null ? "Add SharePoint profile" : "Edit profile";
        PrimaryButtonText = existing == null ? "Save and sign in" : "Save";

        SiteUrlBox.Text = existing?.SiteUrl ?? "";
        NameBox.Text = existing?.Name ?? "";
        ClientIdBox.Text = existing?.ClientId ?? defaultClientId;
        TenantBox.Text = existing?.TenantId ?? "";
        AuthModeBox.SelectedIndex = existing?.AuthMode == AuthMode.Windows ? 1 : 0;
        AdvancedExpander.IsExpanded = ClientIdBox.Text.Length > 0 || TenantBox.Text.Length > 0;

        _nameEdited = !string.IsNullOrEmpty(existing?.Name);

        // TextChanged is raised asynchronously, after the URL handler has returned, so a "setting the name"
        // flag can't tell our updates from typing. Compare against the suggestion instead; clearing the
        // box hands naming back to the URL.
        NameBox.TextChanged += (_, _) =>
            _nameEdited = NameBox.Text.Length > 0 && NameBox.Text != _suggestedName;
    }

    /// <summary>The saved profile (the edited instance, or a new one).</summary>
    public ConnectionProfile? Result { get; private set; }

    /// <summary>True when the site host, tenant, app ID or sign-in method changed.</summary>
    public bool ConnectionChanged { get; private set; }

    private AuthMode SelectedAuthMode => AuthModeBox.SelectedIndex == 1 ? AuthMode.Windows : AuthMode.Browser;

    private void AuthModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e) =>
        BrowserOptions.Visibility = SelectedAuthMode == AuthMode.Browser ? Visibility.Visible : Visibility.Collapsed;

    private void SiteUrlBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        // Default the profile name to the tenant name until the user types their own.
        if (_nameEdited || !Uri.TryCreate(SiteUrlBox.Text.Trim(), UriKind.Absolute, out var uri)) return;
        _suggestedName = TenantName(uri.Host);
        NameBox.Text = _suggestedName;
    }

    /// <summary>"Contoso" for contoso.sharepoint.com, contoso-my.sharepoint.com or contoso-admin.sharepoint.com.</summary>
    private static string TenantName(string host)
    {
        var name = host.Split('.')[0];
        foreach (var suffix in new[] { "-my", "-admin" })
        {
            if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) name = name[..^suffix.Length];
        }
        return name.Length == 0 ? name : char.ToUpperInvariant(name[0]) + name[1..];
    }

    private void ContentDialog_PrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        var url = SiteUrlBox.Text.Trim().TrimEnd('/');
        var clientId = ClientIdBox.Text.Trim();
        var mode = SelectedAuthMode;

        string? error = null;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp))
            error = "Enter the full site URL, starting with https://";
        else if (mode == AuthMode.Browser && uri.Scheme != Uri.UriSchemeHttps)
            error = "Microsoft 365 sites must use https://";
        else if (mode == AuthMode.Browser && clientId.Length > 0 && !Guid.TryParse(clientId, out _))
            error = "The app (client) ID must be a GUID, or leave it empty to use the built-in SharePoint Online sign-in.";

        if (error != null)
        {
            ErrorBar.Message = error;
            ErrorBar.IsOpen = true;
            args.Cancel = true;
            return;
        }

        var profile = _existing ?? new ConnectionProfile();
        ConnectionChanged = _existing == null
            || !string.Equals(new Uri(profile.SiteUrl).Authority, uri!.Authority, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(profile.ClientId, clientId, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(profile.TenantId, TenantBox.Text.Trim(), StringComparison.OrdinalIgnoreCase)
            || profile.AuthMode != mode;

        profile.SiteUrl = url;
        profile.Name = NameBox.Text.Trim();
        profile.AuthMode = mode;
        profile.ClientId = clientId;
        profile.TenantId = TenantBox.Text.Trim();
        Result = profile;
    }
}
