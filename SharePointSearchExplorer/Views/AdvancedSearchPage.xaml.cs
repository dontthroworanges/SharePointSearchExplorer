using System.ComponentModel;
using System.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using SharePointSearchExplorer.Services;
using SharePointSearchExplorer.ViewModels;
using Windows.Foundation;
using Windows.System;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// Builds a query or suggest request from the original tool's advanced settings. The fields, the query box
/// in the title bar and the editable request URL describe the same <see cref="AdvancedRequest"/>: a change
/// to a field or the query text rebuilds the URL, and an edit to the URL is parsed back into them.
/// </summary>
public sealed partial class AdvancedSearchPage : Page
{
    private readonly Dictionary<AdvancedMode, AdvancedRequest> _requests = new()
    {
        [AdvancedMode.Query] = AdvancedRequest.CreateDefault(AdvancedMode.Query),
        [AdvancedMode.Suggest] = AdvancedRequest.CreateDefault(AdvancedMode.Suggest),
    };

    /// <summary>Per request type: how to show a parameter's value in its control.</summary>
    private readonly Dictionary<AdvancedMode, Dictionary<string, Action<string>>> _setters = new()
    {
        [AdvancedMode.Query] = new(StringComparer.OrdinalIgnoreCase),
        [AdvancedMode.Suggest] = new(StringComparer.OrdinalIgnoreCase),
    };

    /// <summary>Per request type: refreshes the options button's summary (and the rank detail box).</summary>
    private readonly Dictionary<AdvancedMode, Action> _refreshOptions = new();

    /// <summary>The request the URL box shows and Run sends.</summary>
    private AdvancedMode _mode = AdvancedMode.Query;

    /// <summary>The URL the current request corresponds to; a different URL in the box is an edit to parse.</summary>
    private string _url = "";

    public AdvancedSearchPage()
    {
        InitializeComponent();
        BuildFields(QueryFields, AdvancedMode.Query);
        BuildFields(SuggestFields, AdvancedMode.Suggest);
        ShowValues(AdvancedMode.Query);
        ShowValues(AdvancedMode.Suggest);
        UseQueryText();
        UseResultsSortAndFilter();
        RebuildUrl();

        ViewModel.SortOrFilterChanged += (_, _) => UseResultsSortAndFilter();
        // The query text is the title bar's query box, and the URL starts with the current profile's site.
        ViewModel.PropertyChanged += ViewModel_PropertyChanged;
        App.Settings.CurrentProfileChanged += (_, _) => RebuildUrl();
    }

    public SearchViewModel ViewModel => App.Search;

    public bool IsNot(bool value) => !value;

    private void ErrorBar_Closed(InfoBar sender, InfoBarClosedEventArgs args) => ViewModel.ErrorMessage = null;

    #region Fields and query text to URL

    /// <summary>
    /// A field's value changed: store it and rebuild the URL, for the field's request type. A no-op when the
    /// value came from the URL.
    /// </summary>
    private void OnFieldChanged(AdvancedMode mode, string key, string value)
    {
        var request = _requests[mode];
        if (request.Get(key) == value) return;
        request.Set(key, value);
        _refreshOptions[mode]();
        _mode = mode;
        ShowMode();
        RebuildUrl();
    }

    private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        // Not while the query text is being set from a URL edit: the URL is then left exactly as typed.
        if (e.PropertyName != nameof(SearchViewModel.QueryText) || _parsingUrl) return;
        UseQueryText();
        RebuildUrl();
    }

    /// <summary>Both requests take their query text from the title bar's query box.</summary>
    private void UseQueryText()
    {
        foreach (var request in _requests.Values) request.Set(AdvancedRequest.QueryTextKey, ViewModel.QueryText.Trim());
    }

    // What the Sort list and Refinement filters fields last received from the Search Results page.
    private string _resultsSort = "";
    private string _resultsFilter = "";

    /// <summary>
    /// The sort and file type filter chosen above Primary Results go into the query's Sort list and Refinement
    /// filters fields. Turning one off there clears the field only if it still holds what was copied, so a
    /// value typed here is not lost.
    /// </summary>
    private void UseResultsSortAndFilter()
    {
        bool changed = Use("sortlist", ViewModel.SortList ?? "", ref _resultsSort);
        changed |= Use("refinementfilters", ViewModel.RefinementFilters ?? "", ref _resultsFilter);
        if (changed && _mode == AdvancedMode.Query) RebuildUrl();

        bool Use(string key, string value, ref string copied)
        {
            if (value == copied) return false;
            var request = _requests[AdvancedMode.Query];
            bool keep = value.Length == 0 && request.Get(key) != copied;
            copied = value;
            if (keep || request.Get(key) == value) return false;

            request.Set(key, value);
            _setters[AdvancedMode.Query][key](value);
            return true;
        }
    }

    private void RebuildUrl()
    {
        _url = _requests[_mode].Build(App.Settings.CurrentProfile?.SiteUrl ?? "");
        UrlBox.Text = WrapUrl(_url);
        UpdateUrlLength();
        ActionText.Text = "";
    }

    private void UpdateUrlLength() => UrlLengthText.Text = $"HTTP GET · {_url.Length:N0} characters";

    #endregion

    #region URL to fields

    private bool _parsingUrl;

    private void UrlBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        var url = UnwrapUrl(UrlBox.Text);

        // Typing moves the line breaks: wrap again and put the caret back where it was in the URL.
        var wrapped = WrapUrl(url);
        if (wrapped != NormalizeBreaks(UrlBox.Text))
        {
            int caret = UnwrapUrl(UrlBox.Text[..Math.Min(UrlBox.SelectionStart, UrlBox.Text.Length)]).Length;
            UrlBox.Text = wrapped;
            UrlBox.SelectionStart = WrappedIndex(caret, url.Length);
            return; // Raised again for the wrapped text.
        }

        // Edited by hand: the fields and the query box follow, and the URL is left exactly as typed.
        if (url != _url) UseUrl(url);
    }

    /// <summary>Makes a URL the current request: its type, its field values and the query text.</summary>
    private void UseUrl(string url)
    {
        _url = url;
        UpdateUrlLength();
        ActionText.Text = "";
        var request = AdvancedRequest.Parse(url);
        _requests[request.Mode] = request;
        _mode = request.Mode;
        ShowMode();
        ShowValues(request.Mode);

        _parsingUrl = true;
        ViewModel.QueryText = request.Get(AdvancedRequest.QueryTextKey);
        _parsingUrl = false;
        UseQueryText();
    }

    /// <summary>Puts the request's values in its controls (their change handlers then find nothing new).</summary>
    private void ShowValues(AdvancedMode mode)
    {
        var request = _requests[mode];
        _showingValues = true;
        foreach (var (key, set) in _setters[mode]) set(request.Get(key));
        _showingValues = false;
        _refreshOptions[mode]();
    }

    /// <summary>
    /// Set while values are pushed into the controls. Check boxes report a change at once, and a half-typed
    /// URL value (e.g. "tru") shows as a dash; without this, that would be written back over what was typed.
    /// </summary>
    private bool _showingValues;

    #endregion

    #region URL wrapping

    // The URL box wraps by itself only at the few characters where text may break (after "?" or "/", say),
    // which leaves most of each line empty. So wrapping is off, and line breaks are put in here instead, at
    // the number of characters that fit the box (its font is monospaced). They are not part of the URL.

    private const string UrlFont = "Consolas";
    private const double UrlFontSize = 13;

    /// <summary>The box's left and right padding and border, plus room for the vertical scroll bar.</summary>
    private const double UrlBoxChrome = 40;

    private double _urlCharWidth;
    private int _urlColumns = int.MaxValue;

    private void UrlBox_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (_urlCharWidth == 0)
        {
            var sample = new TextBlock { FontFamily = new Microsoft.UI.Xaml.Media.FontFamily(UrlFont), FontSize = UrlFontSize, Text = new string('0', 100) };
            sample.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            _urlCharWidth = sample.DesiredSize.Width / 100;
        }

        int columns = Math.Max(20, (int)((e.NewSize.Width - UrlBoxChrome) / _urlCharWidth));
        if (columns == _urlColumns) return;
        _urlColumns = columns;
        UrlBox.Text = WrapUrl(UnwrapUrl(UrlBox.Text));
    }

    private string WrapUrl(string url)
    {
        if (url.Length <= _urlColumns) return url;
        var wrapped = new StringBuilder(url.Length + url.Length / _urlColumns);
        for (int i = 0; i < url.Length; i += _urlColumns)
        {
            if (i > 0) wrapped.Append('\r');
            wrapped.Append(url, i, Math.Min(_urlColumns, url.Length - i));
        }
        return wrapped.ToString();
    }

    private static string UnwrapUrl(string text) => text.Replace("\r", "").Replace("\n", "");

    private static string NormalizeBreaks(string text) => text.Replace("\r\n", "\r").Replace('\n', '\r');

    /// <summary>Where a position in the URL is in the wrapped text (a position at a line's end stays on that line only at the very end).</summary>
    private int WrappedIndex(int urlIndex, int urlLength)
    {
        int breaks = urlIndex / _urlColumns;
        if (urlIndex > 0 && urlIndex == urlLength && urlIndex % _urlColumns == 0) breaks--;
        return urlIndex + breaks;
    }

    #endregion

    #region Request type

    private void ModeTab_Checked(object sender, RoutedEventArgs e)
    {
        // Null while the page is still loading.
        if (UrlBox == null || sender is not RadioButton { Tag: string tag } || !Enum.TryParse<AdvancedMode>(tag, out var mode) || mode == _mode) return;
        _mode = mode;
        ShowMode();
        RebuildUrl();
    }

    /// <summary>Shows which request the URL is for: the selected tab, and the Run button's text.</summary>
    private void ShowMode()
    {
        QueryTab.IsChecked = _mode == AdvancedMode.Query;
        SuggestTab.IsChecked = _mode == AdvancedMode.Suggest;
        RunButtonText.Text = _mode == AdvancedMode.Query ? "Run Query" : "Run Suggest";
    }

    #endregion

    #region Actions

    private async void Run_Click(object sender, RoutedEventArgs e) => await RunAsync();

    private async void UrlBox_PreviewKeyDown(object sender, KeyRoutedEventArgs e)
    {
        // Enter runs the request instead of adding a line break.
        if (e.Key != VirtualKey.Enter) return;
        e.Handled = true;
        await RunAsync();
    }

    /// <summary>Sends the URL as shown (also used by the title bar's search button while this page is open).</summary>
    public async Task RunAsync()
    {
        if (ViewModel.IsBusy) return;
        var url = UnwrapUrl(UrlBox.Text).Trim();

        if (_mode == AdvancedMode.Query)
        {
            if (await ViewModel.RunAdvancedQueryAsync(url, _requests[AdvancedMode.Query].ToSearchQuery()))
                App.MainWindow.ShowSearchPage();
            return;
        }

        if (await ViewModel.RunSuggestAsync(url) is not { } suggestions) return;
        ResultText.Text = suggestions.Count == 0
            ? "No suggestions were returned."
            : $"{suggestions.Count:N0} suggestion{(suggestions.Count == 1 ? "" : "s")}";
        SuggestionList.ItemsSource = suggestions;
        SuggestionList.Visibility = suggestions.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void CopyUrl_Click(object sender, RoutedEventArgs e) => PropertyClipboard.CopyText(UnwrapUrl(UrlBox.Text));

    /// <summary>Adds the request to the query history (without the site, so it runs against any profile).</summary>
    private void Save_Click(object sender, RoutedEventArgs e)
    {
        App.Settings.AddHistory(AdvancedRequest.ToSaved(UnwrapUrl(UrlBox.Text).Trim()));
        ActionText.Text = "Saved to Query History";
    }

    /// <summary>Opens a request saved with Save query and runs it (Run again on the Query History page).</summary>
    public async Task RunSavedAsync(string saved)
    {
        var url = AdvancedRequest.FromSaved(saved, App.Settings.CurrentProfile?.SiteUrl ?? "");
        UseUrl(url);
        UrlBox.Text = WrapUrl(url);
        await RunAsync();
    }

    private void Reset_Click(object sender, RoutedEventArgs e)
    {
        _requests[_mode] = AdvancedRequest.CreateDefault(_mode);
        UseQueryText();
        ShowValues(_mode);
        RebuildUrl();
    }

    #endregion

    #region Building the fields

    /// <summary>
    /// One control per parameter, in the parameter list's order: consecutive numbers share a row, and all the
    /// flags go into a single multi-select drop-down, placed where the first flag is. The query text has no
    /// field: it is the query box in the title bar.
    /// </summary>
    private void BuildFields(StackPanel panel, AdvancedMode mode)
    {
        var defs = AdvancedRequest.ParamsFor(mode).Where(d => d.Key != AdvancedRequest.QueryTextKey).ToList();
        var flags = defs.Where(d => d.Kind == ParamKind.Flag).ToList();
        bool flagsAdded = false;
        _refreshOptions[mode] = () => { };

        for (int i = 0; i < defs.Count; i++)
        {
            var def = defs[i];
            switch (def.Kind)
            {
                case ParamKind.Flag:
                    if (!flagsAdded) panel.Children.Add(OptionsButton(mode, flags));
                    flagsAdded = true;
                    break;

                case ParamKind.Number:
                    var numbers = defs.Skip(i).TakeWhile(d => d.Kind == ParamKind.Number).ToList();
                    panel.Children.Add(NumberRow(mode, numbers));
                    i += numbers.Count - 1;
                    break;

                default:
                    panel.Children.Add(def.Suggestions == null ? TextField(mode, def) : SuggestField(mode, def));
                    break;
            }
        }
    }

    private TextBox TextField(AdvancedMode mode, ParamDef def)
    {
        var box = new TextBox { Header = def.Label, IsSpellCheckEnabled = false, TextWrapping = TextWrapping.Wrap };
        if (def.Tip.Length > 0) ToolTipService.SetToolTip(box, def.Tip);
        box.TextChanged += (_, _) => OnFieldChanged(mode, def.Key, box.Text);
        _setters[mode][def.Key] = value =>
        {
            if (box.Text != value) box.Text = value;
        };
        return box;
    }

    /// <summary>
    /// A text field whose examples (or known values, e.g. cultures) are offered as suggestions: all of them
    /// when the field gets the focus, and the matching ones while typing.
    /// </summary>
    private AutoSuggestBox SuggestField(AdvancedMode mode, ParamDef def)
    {
        var suggestions = def.Suggestions!;
        var box = new AutoSuggestBox { Header = def.Label, UpdateTextOnSelect = false };
        if (def.Tip.Length > 0) ToolTipService.SetToolTip(box, def.Tip);

        box.TextChanged += (_, args) =>
        {
            if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            {
                // Match on the label or the value, so both "English" and "9" find English [en | 9].
                box.ItemsSource = suggestions
                    .Where(s => s.Label.Contains(box.Text, StringComparison.OrdinalIgnoreCase) || s.Value.Contains(box.Text, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }
            OnFieldChanged(mode, def.Key, box.Text);
        };
        box.GotFocus += (_, _) =>
        {
            if (box.FocusState == FocusState.Programmatic) return;
            box.ItemsSource = suggestions;
            box.IsSuggestionListOpen = true;
        };
        box.SuggestionChosen += (_, args) =>
        {
            if (args.SelectedItem is ParamSuggestion chosen) box.Text = chosen.Value;
        };

        _setters[mode][def.Key] = value =>
        {
            if (box.Text != value) box.Text = value;
        };
        return box;
    }

    private Grid NumberRow(AdvancedMode mode, List<ParamDef> defs)
    {
        var row = new Grid { ColumnSpacing = 8 };
        for (int i = 0; i < defs.Count; i++)
        {
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var field = TextField(mode, defs[i]);
            field.TextWrapping = TextWrapping.NoWrap;
            Grid.SetColumn(field, i);
            row.Children.Add(field);
        }
        return row;
    }

    /// <summary>
    /// The flags as a multi-select drop-down. Each box has three states, as in the original tool: checked
    /// sends true, cleared sends false, and the dash leaves the parameter out (SharePoint's default applies).
    /// </summary>
    private StackPanel OptionsButton(AdvancedMode mode, List<ParamDef> flags)
    {
        var summary = new TextBlock { TextTrimming = TextTrimming.CharacterEllipsis };
        var list = new StackPanel { Spacing = 0 };
        var boxes = new List<(ParamDef Def, CheckBox Box)>();

        foreach (var def in flags)
        {
            var box = new CheckBox { Content = def.Label, IsThreeState = true, IsChecked = null };
            if (def.Tip.Length > 0) ToolTipService.SetToolTip(box, def.Tip);
            void Changed(object s, RoutedEventArgs e)
            {
                if (!_showingValues) OnFieldChanged(mode, def.Key, box.IsChecked switch { true => "true", false => "false", null => "" });
            }
            box.Checked += Changed;
            box.Unchecked += Changed;
            box.Indeterminate += Changed;
            _setters[mode][def.Key] = value => box.IsChecked = value.ToLowerInvariant() switch { "true" => true, "false" => false, _ => null };
            boxes.Add((def, box));
            list.Children.Add(box);
        }

        // "Include rank detail" is a select property, not a parameter (see AdvancedRequest.SetRankDetail).
        CheckBox? rankDetail = null;
        if (mode == AdvancedMode.Query)
        {
            var rank = rankDetail = new CheckBox { Content = "Include rank detail" };
            ToolTipService.SetToolTip(rank, "Adds RankDetail to the select properties. Requires admin access, and fewer than 100 results, or the rank detail is not returned.");
            void RankChanged(object s, RoutedEventArgs e)
            {
                var request = _requests[mode];
                var select = AdvancedRequest.SetRankDetail(request.Get(AdvancedRequest.SelectPropertiesKey), rank.IsChecked == true);
                _setters[mode][AdvancedRequest.SelectPropertiesKey](select);
                OnFieldChanged(mode, AdvancedRequest.SelectPropertiesKey, select);
            }
            rank.Checked += RankChanged;
            rank.Unchecked += RankChanged;
            list.Children.Add(new MenuFlyoutSeparator { Margin = new Thickness(-4, 4, -4, 4) });
            list.Children.Add(rank);
        }

        var clear = new Button { Content = "Clear all", HorizontalAlignment = HorizontalAlignment.Right };
        clear.Click += (_, _) =>
        {
            foreach (var (_, box) in boxes) box.IsChecked = null;
            if (rankDetail != null) rankDetail.IsChecked = false;
        };

        var content = new StackPanel { Width = 300, Spacing = 8 };
        content.Children.Add(new TextBlock
        {
            Text = "Checked sends true, cleared sends false, and a dash leaves the option out of the request.",
            TextWrapping = TextWrapping.Wrap,
            Foreground = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["TextFillColorSecondaryBrush"],
            Style = (Style)Application.Current.Resources["CaptionTextBlockStyle"],
        });
        content.Children.Add(new ScrollViewer { MaxHeight = 460, Content = list });
        content.Children.Add(clear);

        var button = new DropDownButton
        {
            Content = summary,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Left,
            Flyout = new Flyout { Content = content, Placement = Microsoft.UI.Xaml.Controls.Primitives.FlyoutPlacementMode.BottomEdgeAlignedLeft },
        };
        AutomationProperties.SetName(button, "Options");

        _refreshOptions[mode] = () =>
        {
            var request = _requests[mode];
            if (rankDetail != null) rankDetail.IsChecked = AdvancedRequest.HasRankDetail(request.Get(AdvancedRequest.SelectPropertiesKey));

            // e.g. "Trim duplicates: off, Enable FQL: on"
            var set = boxes
                .Where(b => request.Get(b.Def.Key).Length > 0)
                .Select(b => $"{b.Def.Label}: {(request.Get(b.Def.Key).Equals("true", StringComparison.OrdinalIgnoreCase) ? "on" : "off")}")
                .ToList();
            if (rankDetail?.IsChecked == true) set.Add("Include rank detail: on");
            summary.Text = set.Count == 0 ? "None set" : string.Join(", ", set);
            ToolTipService.SetToolTip(button, set.Count == 0 ? null : string.Join("\n", set));
        };

        var field = new StackPanel { Spacing = 4 };
        field.Children.Add(new TextBlock { Text = "Options" });
        field.Children.Add(button);
        return field;
    }

    #endregion
}
