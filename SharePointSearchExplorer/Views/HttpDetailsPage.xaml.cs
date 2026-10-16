using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointSearchExplorer.Models;
using SharePointSearchExplorer.Services;
using SharePointSearchExplorer.ViewModels;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// The "Server Response" page: the HTTP request and response of the last search, as a summary
/// ("HTTP Details"), the headers and body ("Full API Response") and how the query was processed ("Debug").
/// </summary>
public sealed partial class HttpDetailsPage : Page
{
    private enum Tab { Details, Response, Debug }

    private Tab _tab;

    // Each tab's text for the current exchange, built the first time the tab is shown.
    private readonly Dictionary<Tab, string> _text = new();

    public HttpDetailsPage()
    {
        InitializeComponent();
        Finder.Target = Code;
        ShowActiveTab();
        App.Search.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != nameof(SearchViewModel.LastExchange)) return;
            _text.Clear();
            ShowActiveTab();
        };
    }

    private static HttpExchange? Exchange => App.Search.LastExchange;

    private void Tab_Checked(object sender, RoutedEventArgs e)
    {
        if (sender is RadioButton { Tag: string tag } && Enum.TryParse<Tab>(tag, out var tab)) _tab = tab;
        // Null while the page is still loading.
        if (Code != null) ShowActiveTab();
    }

    private void ShowActiveTab()
    {
        if (Exchange is not { } http)
        {
            Code.Text = "";
        }
        else
        {
            if (!_text.TryGetValue(_tab, out var text))
            {
                _text[_tab] = text = _tab switch
                {
                    Tab.Response => HttpDetailsFormatter.Response(http),
                    Tab.Debug => HttpDetailsFormatter.Debug(http),
                    _ => HttpDetailsFormatter.Details(http),
                };
            }
            Code.Text = text;
        }
        CopyButton.IsEnabled = PopOutButton.IsEnabled = Code.Text.Length > 0;
    }

    private void Copy_Click(object sender, RoutedEventArgs e) => PropertyClipboard.CopyText(Code.Text);

    /// <summary>Opens the active tab's text in its own window, as a snapshot that later searches don't change.</summary>
    private void PopOut_Click(object sender, RoutedEventArgs e)
    {
        if (Exchange is not { } http) return;
        var tabName = _tab switch
        {
            Tab.Response => "Full API Response",
            Tab.Debug => "Debug",
            _ => "HTTP Details",
        };
        CodeWindow.Show(tabName, $"{http.Method} {http.Url}", Code.Text);
    }
}
