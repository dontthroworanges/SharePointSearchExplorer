using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using SharePointSearchExplorer.Helpers;
using SharePointSearchExplorer.Models;
using SharePointSearchExplorer.Services;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// Recent queries, with "Run again", "Export" and "Remove" for each, plus "Clear history" and XML import/export
/// of the whole history.
/// </summary>
public sealed partial class QueryHistoryPage : Page
{
    public QueryHistoryPage()
    {
        InitializeComponent();
        UpdateHistory();

        // The query box (in the title bar, beside this page) can remove queries too.
        App.Settings.HistoryChanged += OnHistoryChanged;
        Unloaded += (_, _) => App.Settings.HistoryChanged -= OnHistoryChanged;
    }

    private void OnHistoryChanged(object? sender, EventArgs e) => UpdateHistory();

    private static AppSettings Settings => App.Settings.Data;

    private void ContentScroller_SizeChanged(object sender, SizeChangedEventArgs e) =>
        SettingsPage.CenterContent(ContentPanel, e.NewSize.Width);

    /// <summary>Always asks first (unlike removing one query, this can't be turned off in Settings).</summary>
    private async void ClearHistory_Click(object sender, RoutedEventArgs e)
    {
        int count = Settings.History.Count;
        bool confirmed = await App.MainWindow.ConfirmAsync(
            "Clear query history?",
            $"All {count} saved quer{(count == 1 ? "y" : "ies")} will be removed from the query box suggestions and this page. This can't be undone.\n\nTo keep a copy, cancel and use Export all first.",
            "Clear history");
        if (confirmed) App.Settings.ClearHistory();
    }

    #region Import and export

    private async void Import_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FileOpenPicker { SuggestedStartLocation = PickerLocationId.DocumentsLibrary };
        picker.FileTypeFilter.Add(QueryFile.Extension);
        InitializeWithWindow.Initialize(picker, App.WindowHandle);
        if (await picker.PickSingleFileAsync() is not { } file) return;

        List<string> queries;
        try
        {
            queries = QueryFile.Load(file.Path);
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
        {
            ShowStatus(InfoBarSeverity.Error, "Couldn't import queries", $"{file.Name}: {ex.Message}");
            return;
        }
        if (queries.Count == 0)
        {
            ShowStatus(InfoBarSeverity.Warning, "No queries to import", $"{file.Name} doesn't contain any queries.");
            return;
        }

        var (added, skipped, dropped) = App.Settings.ImportHistory(queries);
        UpdateHistory();

        var message = $"{Plural(added, "query was", "queries were")} added from {file.Name}.";
        if (skipped > 0) message += $" {Plural(skipped, "was", "were")} already in your history.";
        if (dropped > 0) message += $" History keeps {SettingsStore.MaxHistory} queries, so the {Plural(dropped, "oldest was", "oldest were")} removed.";
        ShowStatus(added > 0 ? InfoBarSeverity.Success : InfoBarSeverity.Informational, "Import complete", message);
    }

    private async void ExportAll_Click(object sender, RoutedEventArgs e) =>
        await ExportAsync(Settings.History.ToList(), $"Saved queries {DateTime.Now:yyyy-MM-dd}");

    private async Task ExportAsync(IReadOnlyList<string> queries, string suggestedName)
    {
        var picker = new FileSavePicker
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            SuggestedFileName = suggestedName,
            DefaultFileExtension = QueryFile.Extension,
        };
        picker.FileTypeChoices.Add("Saved queries (XML)", new List<string> { QueryFile.Extension });
        InitializeWithWindow.Initialize(picker, App.WindowHandle);
        if (await picker.PickSaveFileAsync() is not { } file) return;

        try
        {
            QueryFile.Save(file.Path, queries);
            ShowStatus(InfoBarSeverity.Success, "Export complete", $"{Plural(queries.Count, "query was", "queries were")} saved to {file.Path}.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowStatus(InfoBarSeverity.Error, "Couldn't export queries", ex.Message);
        }
    }

    private void ShowStatus(InfoBarSeverity severity, string title, string message)
    {
        StatusBar.Severity = severity;
        StatusBar.Title = title;
        StatusBar.Message = message;
        StatusBar.IsOpen = true;
    }

    /// <summary>"1 query was" / "3 queries were"; with no noun, "1 was" / "3 were".</summary>
    private static string Plural(int count, string one, string many) => $"{count} {(count == 1 ? one : many)}";

    #endregion

    private void UpdateHistory()
    {
        int count = Settings.History.Count;
        HistoryCard.Description = count == 0
            ? "No saved queries."
            : $"{count} recent quer{(count == 1 ? "y is" : "ies are")} suggested as you type in the query box.";
        ClearHistoryButton.IsEnabled = count > 0;
        ExportAllButton.IsEnabled = count > 0;

        HistoryList.Children.Clear();
        foreach (var query in Settings.History) HistoryList.Children.Add(HistoryRow(query));
        SavedQueriesHeader.Visibility = HistoryList.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    /// <summary>A saved query with "Run again" and "Remove" buttons.</summary>
    private Border HistoryRow(string query)
    {
        var text = new TextBlock { Text = query, VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
        ToolTipService.SetToolTip(text, query);

        var rerun = IconButton(FluentIcons.ArrowClockwise, "Run again", async () => await App.MainWindow.RunQueryAsync(query));
        var export = IconButton(FluentIcons.ArrowExportLtr, "Export to XML file", async () => await ExportAsync(new[] { query }, QueryFile.SuggestedName(query)));
        var remove = IconButton(FluentIcons.Dismiss, "Remove from history", () => App.Settings.RemoveHistory(query));

        var grid = new Grid
        {
            ColumnSpacing = 16,
            ColumnDefinitions =
            {
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                new ColumnDefinition { Width = GridLength.Auto },
            },
        };
        // A request saved from the Advanced Query Rules page has that page's icon; Run again opens it there.
        bool advanced = AdvancedRequest.IsSaved(query);
        var icon = FluentIcons.Create(advanced ? FluentIcons.SearchSettings : FluentIcons.Search, 16);
        icon.VerticalAlignment = VerticalAlignment.Center;
        if (advanced) ToolTipService.SetToolTip(icon, "Advanced Query Rules request");
        if (advanced && !App.Settings.Data.ShowAdvancedPages)
        {
            // Run again opens the Advanced Query Rules page, which is turned off in Settings.
            rerun.IsEnabled = false;
            ToolTipService.SetToolTip(rerun, "Turn on Settings › Search › Advanced pages to run this request");
        }
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 2, Children = { rerun, export, remove } };
        Grid.SetColumn(text, 1);
        Grid.SetColumn(buttons, 2);
        grid.Children.Add(icon);
        grid.Children.Add(text);
        grid.Children.Add(buttons);

        return new Border { Style = (Style)Resources["HistoryRowStyle"], Child = grid };

        static Button IconButton(string glyph, string label, Action onClick)
        {
            var button = new Button
            {
                Content = FluentIcons.Create(glyph, 16),
                Style = (Style)Application.Current.Resources["IconButtonStyle"],
            };
            AutomationProperties.SetName(button, label);
            ToolTipService.SetToolTip(button, label);
            button.Click += (_, _) => onClick();
            return button;
        }
    }
}
