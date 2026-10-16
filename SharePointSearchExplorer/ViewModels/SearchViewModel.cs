using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.Identity.Client;
using SharePointSearchExplorer.Helpers;
using SharePointSearchExplorer.Models;
using SharePointSearchExplorer.Services;

namespace SharePointSearchExplorer.ViewModels;

public enum SortOption { Relevance, NewestFirst, OldestFirst, LargestFirst, SmallestFirst }

public partial class SearchViewModel : ObservableObject
{
    private const string FileTypeRefiner = "FileType";
    private const string ContentClassRefiner = "ContentClass";

    // Every value, not just the default top entries, so the filter accounts for all results.
    private const string AllEntries = "(filter=1000/0/*)";

    private readonly SettingsStore _settings;
    private readonly SearchService _search;
    private readonly AuthService _auth;

    private CancellationTokenSource? _searchCts;
    private CancellationTokenSource? _propertiesCts;

    /// <summary>The query that produced the current results (reused for paging and property lookups).</summary>
    private SearchQuery? _currentQuery;

    public SearchViewModel(SettingsStore settings, SearchService search, AuthService auth)
    {
        _settings = settings;
        _search = search;
        _auth = auth;

        Results.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(HasResults));
            OnPropertyChanged(nameof(ShowEmptyState));
        };
        _settings.CurrentProfileChanged += (_, _) => OnProfileChanged();
        _auth.AccountChanged += (_, profile) =>
        {
            _settings.Save();
            if (profile == Profile) OnPropertyChanged(nameof(Profile));
        };
        ShowEmptyProperties = settings.Data.ShowEmptyProperties;
        PropertiesHeader.SourceFilter = settings.Data.DefaultPropertySource;
        PropertiesHeader.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName is nameof(ColumnHeaderInfo.SortDescending) or nameof(ColumnHeaderInfo.SourceFilter))
                RefreshVisibleProperties();
        };
        UpdateEmptyStateText();
    }

    /// <summary>Raised when an action needs a profile but none exists yet.</summary>
    public event EventHandler? ProfileRequired;

    public ObservableCollection<SearchResultItem> Results { get; } = new();
    public ObservableCollection<RefinerOption> FileTypeOptions { get; } = new();
    public ObservableCollection<PropertyEntry> VisibleProperties { get; } = new();

    /// <summary>The selected result's favorite properties, for the Result Details summary.</summary>
    public ObservableCollection<PropertyEntry> SummaryFacts { get; } = new();

    public ConnectionProfile? Profile => _settings.CurrentProfile;

    [ObservableProperty]
    private string _queryText = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelection))]
    [NotifyPropertyChangedFor(nameof(ShowNoSelection))]
    private SearchResultItem? _selectedResult;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowEmptyState))]
    private bool _isBusy;

    [ObservableProperty]
    private string _statusText = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasError))]
    private string? _errorMessage;

    [ObservableProperty]
    private string _emptyStateText = "";

    /// <summary>The request and response of the last search, for the Server Response page; null until a search runs.</summary>
    [ObservableProperty]
    private HttpExchange? _lastExchange;

    [ObservableProperty]
    private bool _hasMoreResults;

    [ObservableProperty]
    private bool _isFilterActive;

    [ObservableProperty]
    private SortOption _sort = SortOption.Relevance;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsGrouped))]
    private GroupOption _groupBy = GroupOption.None;

    /// <summary>The loaded results in groups, while <see cref="GroupBy"/> is set; otherwise empty.</summary>
    [ObservableProperty]
    private IReadOnlyList<ResultGroup> _resultGroups = [];

    [ObservableProperty]
    private bool _isLoadingProperties;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPropertiesError))]
    private string? _propertiesError;

    [ObservableProperty]
    private string _propertyFilter = "";

    /// <summary>No favorite properties are chosen, so the summary shows where to add them.</summary>
    [ObservableProperty]
    private bool _hasNoFavorites;

    /// <summary>
    /// List properties with no value in the Result Details table. Starts from Settings › Appearance; the table's
    /// toggle changes it for this session only.
    /// </summary>
    [ObservableProperty]
    private bool _showEmptyProperties;

    /// <summary>The property table's header state: the property count, the sort direction and the Source filter.</summary>
    public ColumnHeaderInfo PropertiesHeader { get; } = new("Property");

    public bool HasResults => Results.Count > 0;
    public bool ShowEmptyState => Results.Count == 0 && !IsBusy;
    public bool HasSelection => SelectedResult != null;
    public bool ShowNoSelection => SelectedResult == null;
    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);
    public bool HasPropertiesError => !string.IsNullOrEmpty(PropertiesError);
    public bool IsGrouped => GroupBy != GroupOption.None;

    #region Search

    /// <summary>Runs a new query. <paramref name="keepRefiners"/> keeps the file-type filter (used by sort/filter changes).</summary>
    public async Task SearchAsync(bool keepRefiners = false)
    {
        var profile = Profile;
        if (profile == null)
        {
            ErrorMessage = "Add a SharePoint site profile to start searching.";
            ProfileRequired?.Invoke(this, EventArgs.Empty);
            return;
        }

        var text = string.IsNullOrWhiteSpace(QueryText) ? "*" : QueryText.Trim();
        // Sort and filter changes rerun the same query, so properties already loaded are still valid.
        if (!keepRefiners || _currentQuery?.QueryText != text) _propertyCache.Clear();
        if (!keepRefiners)
        {
            FileTypeOptions.Clear();
            IsFilterActive = false;
            SortOrFilterChanged?.Invoke(this, EventArgs.Empty);
        }

        var query = BuildQuery(text, startRow: 0);
        var ct = BeginSearch();
        try
        {
            var response = await _search.QueryAsync(profile, query, ct);
            ct.ThrowIfCancellationRequested();

            LastExchange = response.Http;
            _currentQuery = query;
            _advancedResults = false;
            Results.Clear();
            foreach (var row in response.Rows) Results.Add(row);
            RegroupResults();
            // The top of the list, which differs from the top result while grouped.
            SelectedResult = IsGrouped ? ResultGroups.FirstOrDefault()?.FirstOrDefault() : Results.FirstOrDefault();
            HasMoreResults = response.TotalRows > Results.Count;
            StatusText = $"{response.TotalRows:N0} result{(response.TotalRows == 1 ? "" : "s")} · {response.ElapsedMilliseconds:N0} ms";

            if (!keepRefiners && response.Refiners.TryGetValue(FileTypeRefiner, out var entries))
            {
                foreach (var e in entries) FileTypeOptions.Add(new RefinerOption(e.Name, e.Token, e.Count));
            }

            if (text != "*") _settings.AddHistory(text);
            UpdateEmptyStateText(searched: true);

            if (!keepRefiners && response.TotalRows > 0) await AddNoFileTypeOptionsAsync(profile, query, ct);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            StatusText = "Search canceled";
        }
        catch (Exception ex) when (IsSignInCanceled(ex))
        {
            StatusText = "Sign-in canceled";
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            if (ex is SearchException { Http: { } http }) LastExchange = http;
        }
        finally
        {
            EndSearch(ct);
        }
    }

    /// <summary>
    /// List items, folders, sites and other results without a file type have no FileType refiner entry, so they
    /// would be missing from the filter. A refiner-only query for just those results (none of the file types
    /// found) breaks them down by ContentClass, e.g. STS_ListItem_GenericList. Failure leaves the file types only.
    /// </summary>
    private async Task AddNoFileTypeOptionsAsync(ConnectionProfile profile, SearchQuery query, CancellationToken ct)
    {
        var classQuery = new SearchQuery
        {
            QueryText = query.QueryText,
            RowLimit = 0,
            TrimDuplicates = query.TrimDuplicates,
            Refiners = ContentClassRefiner + AllEntries,
            RefinementFilters = NoFileTypeFilter(),
        };
        try
        {
            var response = await _search.QueryAsync(profile, classQuery, ct);
            ct.ThrowIfCancellationRequested();
            if (!response.Refiners.TryGetValue(ContentClassRefiner, out var entries)) return;
            foreach (var e in entries.Where(e => e.Count > 0))
                FileTypeOptions.Add(new RefinerOption(e.Name, e.Token, e.Count, ContentClassRefiner));
        }
        catch (Exception) when (!ct.IsCancellationRequested)
        {
        }
    }

    /// <summary>The current results came from the Advanced Query Rules page, whose request the sort and filter buttons can't rebuild.</summary>
    private bool _advancedResults;

    /// <summary>
    /// Runs a query URL from the Advanced Query Rules page and shows its results as Primary Results; true when it
    /// succeeded. <paramref name="query"/> describes the request for property lookups and the Debug tab.
    /// </summary>
    public async Task<bool> RunAdvancedQueryAsync(string url, SearchQuery query)
    {
        var profile = Profile;
        if (profile == null)
        {
            ErrorMessage = "Add a SharePoint site profile to start searching.";
            ProfileRequired?.Invoke(this, EventArgs.Empty);
            return false;
        }

        var ct = BeginSearch();
        try
        {
            var http = await _search.GetAsync(profile, url, ct, query);
            ct.ThrowIfCancellationRequested();
            LastExchange = http;
            var response = SearchService.ParseResponse(http);

            _propertyCache.Clear();
            _currentQuery = query;
            _advancedResults = true;
            FileTypeOptions.Clear();
            IsFilterActive = false;
            Results.Clear();
            foreach (var row in response.Rows) Results.Add(row);
            RegroupResults();
            SelectedResult = IsGrouped ? ResultGroups.FirstOrDefault()?.FirstOrDefault() : Results.FirstOrDefault();
            // Paging is part of the advanced request (start row and row limit), so there is nothing to load more of.
            HasMoreResults = false;
            StatusText = $"{response.TotalRows:N0} result{(response.TotalRows == 1 ? "" : "s")} · {response.ElapsedMilliseconds:N0} ms";
            UpdateEmptyStateText(searched: true);
            return true;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            StatusText = "Search canceled";
        }
        catch (Exception ex) when (IsSignInCanceled(ex))
        {
            StatusText = "Sign-in canceled";
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            if (ex is SearchException { Http: { } failed }) LastExchange = failed;
        }
        finally
        {
            EndSearch(ct);
        }
        return false;
    }

    /// <summary>Runs a suggest URL from the Advanced Query Rules page; null when it failed or was canceled.</summary>
    public async Task<IReadOnlyList<string>?> RunSuggestAsync(string url)
    {
        var profile = Profile;
        if (profile == null)
        {
            ErrorMessage = "Add a SharePoint site profile to start searching.";
            ProfileRequired?.Invoke(this, EventArgs.Empty);
            return null;
        }

        var ct = BeginSearch();
        try
        {
            var http = await _search.GetAsync(profile, url, ct);
            ct.ThrowIfCancellationRequested();
            LastExchange = http;
            return SearchService.ParseSuggestions(http);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
        }
        catch (Exception ex) when (IsSignInCanceled(ex))
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            if (ex is SearchException { Http: { } failed }) LastExchange = failed;
        }
        finally
        {
            EndSearch(ct);
        }
        return null;
    }

    public async Task LoadMoreAsync()
    {
        var profile = Profile;
        if (profile == null || _currentQuery == null || IsBusy) return;

        var query = BuildQuery(_currentQuery.QueryText, startRow: Results.Count);
        var ct = BeginSearch();
        try
        {
            var response = await _search.QueryAsync(profile, query, ct);
            LastExchange = response.Http;
            foreach (var row in response.Rows) Results.Add(row);
            RegroupResults();
            HasMoreResults = response.Rows.Count > 0 && response.TotalRows > Results.Count;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            if (ex is SearchException { Http: { } http }) LastExchange = http;
        }
        finally
        {
            EndSearch(ct);
        }
    }

    public void CancelSearch() => _searchCts?.Cancel();

    public Task SetSortAsync(SortOption sort)
    {
        Sort = sort;
        SortOrFilterChanged?.Invoke(this, EventArgs.Empty);
        return _currentQuery == null || _advancedResults ? Task.CompletedTask : SearchAsync(keepRefiners: true);
    }

    public Task ApplyFiltersAsync()
    {
        if (_advancedResults) return Task.CompletedTask;
        IsFilterActive = FileTypeOptions.Any(o => o.IsSelected);
        SortOrFilterChanged?.Invoke(this, EventArgs.Empty);
        return SearchAsync(keepRefiners: true);
    }

    public Task ClearFiltersAsync()
    {
        foreach (var o in FileTypeOptions) o.IsSelected = false;
        return ApplyFiltersAsync();
    }

    partial void OnGroupByChanged(GroupOption value) => RegroupResults();

    /// <summary>Rebuilds <see cref="ResultGroups"/> from the loaded results (called after each batch, not per row).</summary>
    private void RegroupResults() => ResultGroups = IsGrouped ? ResultGrouping.Group(Results, GroupBy) : [];

    public IEnumerable<string> GetHistorySuggestions(string text) =>
        _settings.Data.History
            // Requests saved from the Advanced Query Rules page are in the history too, but they aren't query texts.
            .Where(h => !AdvancedRequest.IsSaved(h))
            .Where(h => string.IsNullOrWhiteSpace(text) || h.Contains(text, StringComparison.OrdinalIgnoreCase))
            .Take(10);

    private SearchQuery BuildQuery(string text, int startRow)
    {
        var data = _settings.Data;
        return new SearchQuery
        {
            QueryText = text,
            StartRow = startRow,
            RowLimit = Math.Clamp(data.RowLimit, 1, 500),
            // The favorites are requested too, so the summary has their values without waiting for the property table.
            SelectProperties = SearchService.DefaultSelectProperties
                .Union(data.FavoriteProperties, StringComparer.OrdinalIgnoreCase)
                .ToList(),
            Refiners = FileTypeRefiner + AllEntries,
            RefinementFilters = BuildFileTypeFilter(),
            SortList = SortList,
            TrimDuplicates = data.TrimDuplicates,
        };
    }

    /// <summary>
    /// Raised when the Primary Results sort or file type filter changes (see <see cref="SortList"/> and
    /// <see cref="RefinementFilters"/>); the Advanced Query Rules page copies them into its fields.
    /// </summary>
    public event EventHandler? SortOrFilterChanged;

    /// <summary>The Primary Results sort as a sortlist parameter; null for relevance.</summary>
    public string? SortList => Sort switch
    {
        SortOption.NewestFirst => "LastModifiedTime:descending",
        SortOption.OldestFirst => "LastModifiedTime:ascending",
        SortOption.LargestFirst => "Size:descending",
        SortOption.SmallestFirst => "Size:ascending",
        _ => null,
    };

    /// <summary>
    /// The Primary Results file type filter as a refinementfilters parameter, with readable values
    /// (<c>FileType:"xlsx"</c>); null when no type is ticked. A ticked content class matches only its results
    /// without a file type, as counted in the flyout:
    /// <c>andnot(ContentClass:"STS_ListItem_GenericList",FileType:or("docx","xlsx"))</c>.
    /// </summary>
    public string? RefinementFilters => BuildFileTypeFilter(readable: true);

    /// <param name="readable">
    /// SharePoint returns each refiner value as a token, <c>"ǂǂ786c7378"</c>: the value's UTF-8 bytes in hex.
    /// Searches from this page send the tokens as received; for display they are turned back into the value.
    /// </param>
    private string? BuildFileTypeFilter(bool readable = false)
    {
        var selected = FileTypeOptions.Where(o => o.IsSelected && !string.IsNullOrEmpty(o.Token)).ToList();
        var parts = new List<string>();

        var types = Tokens(selected.Where(o => o.Refiner == FileTypeRefiner), readable);
        if (types.Count > 0) parts.Add(Match(FileTypeRefiner, types));

        var classes = Tokens(selected.Where(o => o.Refiner == ContentClassRefiner), readable);
        if (classes.Count > 0)
        {
            var allTypes = AllFileTypeTokens(readable);
            parts.Add(allTypes.Count == 0
                ? Match(ContentClassRefiner, classes)
                : $"andnot({Match(ContentClassRefiner, classes)},{Match(FileTypeRefiner, allTypes)})");
        }

        return parts.Count switch
        {
            0 => null,
            1 => parts[0],
            _ => $"or({string.Join(",", parts)})",
        };
    }

    /// <summary>Matches the results with none of the file types found; null when there are none.</summary>
    private string? NoFileTypeFilter()
    {
        var allTypes = AllFileTypeTokens(readable: false);
        return allTypes.Count == 0 ? null : $"not({Match(FileTypeRefiner, allTypes)})";
    }

    private List<string> AllFileTypeTokens(bool readable) =>
        Tokens(FileTypeOptions.Where(o => o.Refiner == FileTypeRefiner && !string.IsNullOrEmpty(o.Token)), readable);

    private static List<string> Tokens(IEnumerable<RefinerOption> options, bool readable) =>
        options.Select(o => readable ? ReadableToken(o.Token) : o.Token).ToList();

    private static string Match(string refiner, List<string> tokens) =>
        tokens.Count == 1 ? $"{refiner}:{tokens[0]}" : $"{refiner}:or({string.Join(",", tokens)})";

    /// <summary><c>"ǂǂ786c7378"</c> as <c>"xlsx"</c>; a token in any other form is returned unchanged.</summary>
    private static string ReadableToken(string token)
    {
        const string prefix = "\"ǂǂ";
        if (!token.StartsWith(prefix, StringComparison.Ordinal) || !token.EndsWith('"') || token.Length <= prefix.Length) return token;
        try
        {
            var value = System.Text.Encoding.UTF8.GetString(Convert.FromHexString(token[prefix.Length..^1]));
            return $"\"{value.Replace("\\", "\\\\").Replace("\"", "\\\"")}\"";
        }
        catch (FormatException)
        {
            return token;
        }
    }

    private static bool IsSignInCanceled(Exception ex) =>
        ex is SignInCanceledException || ex is MsalClientException { ErrorCode: MsalError.AuthenticationCanceledError };

    private CancellationToken BeginSearch()
    {
        _searchCts?.Cancel();
        _searchCts = new CancellationTokenSource();
        ErrorMessage = null;
        IsBusy = true;
        return _searchCts.Token;
    }

    private void EndSearch(CancellationToken ct)
    {
        // Only the most recent search clears the busy state.
        if (_searchCts?.Token == ct) IsBusy = false;
    }

    #endregion

    #region Profiles

    /// <summary>Tries a silent sign-in so the profile shows who is signed in; never opens a sign-in window.</summary>
    public async Task RefreshSignInStatusAsync()
    {
        var profile = Profile;
        if (profile == null) return;
        if (profile.AuthMode == AuthMode.Windows)
        {
            // Windows profiles always use the signed-in Windows account; there is no separate sign-in.
            profile.IsSignedIn = true;
            return;
        }
        try
        {
            await _auth.GetAccessTokenAsync(profile, allowInteractive: false);
        }
        catch (Exception)
        {
            // Not signed in yet; the first search signs in through the browser.
        }
    }

    /// <summary>Opens the browser to sign in, optionally forcing the account picker.</summary>
    public async Task SignInAsync(bool chooseAccount)
    {
        var profile = Profile;
        if (profile is not { AuthMode: AuthMode.Browser }) return;

        var ct = BeginSearch();
        StatusText = "Waiting for sign-in…";
        try
        {
            await _auth.GetAccessTokenAsync(profile, allowInteractive: true, chooseAccount, ct);
            StatusText = $"Signed in as {profile.UserName}";
        }
        catch (Exception ex) when (ct.IsCancellationRequested || IsSignInCanceled(ex))
        {
            StatusText = "Sign-in canceled";
        }
        catch (Exception ex)
        {
            StatusText = "";
            ErrorMessage = ex.Message;
        }
        finally
        {
            EndSearch(ct);
        }
    }

    public async Task SignOutAsync()
    {
        if (Profile is { } profile)
        {
            await _auth.SignOutAsync(profile);
            StatusText = "Signed out";
        }
    }

    private void OnProfileChanged()
    {
        ClearResults();
        OnPropertyChanged(nameof(Profile));
        _ = RefreshSignInStatusAsync();
    }

    /// <summary>
    /// Empties Primary Results, the filter and the details pane, back to the state before any search
    /// (used when the profile changes and when the query box is cleared with its X button).
    /// </summary>
    public void ClearResults()
    {
        CancelSearch();
        _propertiesCts?.Cancel();
        _currentQuery = null;
        _propertyCache.Clear();
        LastExchange = null;
        Results.Clear();
        RegroupResults();
        FileTypeOptions.Clear();
        IsFilterActive = false;
        SortOrFilterChanged?.Invoke(this, EventArgs.Empty);
        HasMoreResults = false;
        SelectedResult = null;
        ErrorMessage = null;
        StatusText = "";
        UpdateEmptyStateText();
    }

    private void UpdateEmptyStateText(bool searched = false)
    {
        EmptyStateText = Profile is not { } p
            ? "Add a SharePoint site to get started."
            : searched
                ? "No results. Try a different query or clear the filters."
                : $"Search {p.EffectiveName}";
    }

    #endregion

    #region Property table

    /// <summary>
    /// Properties loaded for the current query, by WorkId. Sort, filter and Load more rerun the query and create
    /// new result items; this keeps their properties from being looked up again. Cleared by a new query, a profile
    /// change or the query box's X; Reload replaces one item's entry.
    /// </summary>
    private readonly Dictionary<string, (IReadOnlyList<PropertyEntry>? Managed, IReadOnlyList<PropertyEntry>? Crawled)> _propertyCache = new();

    private void RestoreCachedProperties(SearchResultItem item)
    {
        if (string.IsNullOrEmpty(item.WorkId) || !_propertyCache.TryGetValue(item.WorkId, out var cached)) return;
        item.ManagedProperties ??= cached.Managed;
        item.CrawledProperties ??= cached.Crawled;
    }

    private void CacheProperties(SearchResultItem item)
    {
        if (!string.IsNullOrEmpty(item.WorkId)) _propertyCache[item.WorkId] = (item.ManagedProperties, item.CrawledProperties);
    }

    partial void OnSelectedResultChanged(SearchResultItem? value)
    {
        if (value != null) RestoreCachedProperties(value);
        _propertiesCts?.Cancel();
        IsLoadingProperties = false;
        PropertiesError = null;
        RefreshVisibleProperties();
        _ = EnsurePropertiesLoadedAsync();
    }

    partial void OnPropertyFilterChanged(string value) => RefreshVisibleProperties();

    partial void OnShowEmptyPropertiesChanged(bool value) => RefreshVisibleProperties();

    public Task ReloadPropertiesAsync()
    {
        if (SelectedResult is { } item)
        {
            item.ManagedProperties = null;
            item.CrawledProperties = null;
            if (!string.IsNullOrEmpty(item.WorkId)) _propertyCache.Remove(item.WorkId);
        }
        return EnsurePropertiesLoadedAsync();
    }

    /// <summary>
    /// Loads the selected item's managed and crawled properties the first time it is selected. Each list is
    /// loaded on its own, so the table still shows one when the other fails.
    /// </summary>
    private async Task EnsurePropertiesLoadedAsync()
    {
        var item = SelectedResult;
        var profile = Profile;
        var query = _currentQuery;
        if (item == null || profile == null || query == null) return;
        bool needManaged = item.ManagedProperties == null;
        bool needCrawled = item.CrawledProperties == null;
        if (!needManaged && !needCrawled) return;

        _propertiesCts?.Cancel();
        var cts = _propertiesCts = new CancellationTokenSource();
        IsLoadingProperties = true;
        PropertiesError = null;
        RefreshVisibleProperties();

        var errors = new List<string>();
        try
        {
            if (needManaged)
            {
                try
                {
                    item.ManagedProperties = await _search.GetManagedPropertiesAsync(profile, item, query, cts.Token);
                }
                catch (Exception ex) when (!cts.IsCancellationRequested)
                {
                    errors.Add(ex.Message);
                }
            }
            if (needCrawled)
            {
                try
                {
                    item.CrawledProperties = await _search.GetCrawledPropertiesAsync(profile, item, query, cts.Token);
                }
                catch (Exception ex) when (!cts.IsCancellationRequested)
                {
                    errors.Add(ex.Message);
                }
            }
        }
        catch (Exception) when (cts.IsCancellationRequested)
        {
            return;
        }

        if (_propertiesCts != cts || cts.IsCancellationRequested) return;
        CacheProperties(item);
        if (errors.Count > 0) PropertiesError = string.Join(" ", errors);
        IsLoadingProperties = false;
        if (SelectedResult == item) RefreshVisibleProperties();
    }

    private void RefreshVisibleProperties()
    {
        PropertiesHeader.Fill(VisibleProperties, SelectedResult?.AllProperties, PropertyFilter, ShowEmptyProperties);
        RefreshSummary();
    }

    /// <summary>
    /// Rebuilds <see cref="SummaryFacts"/>: the favorite properties (Settings › Search) of the selected result.
    /// Values come from the search result, which requests the favorites, or else from the item's managed
    /// properties once loaded (results of an Advanced Query Rules request, or found before a favorite was added).
    /// </summary>
    public void RefreshSummary()
    {
        var favorites = _settings.Data.FavoriteProperties;
        HasNoFavorites = favorites.Count == 0;
        SummaryFacts.Clear();
        if (SelectedResult is not { } item) return;
        foreach (var name in favorites.Take(FavoritePropertyCatalog.MaxCount))
        {
            var value = item.Get(name);
            if (value.Length == 0)
                value = item.ManagedProperties?.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))?.Value ?? "";
            SummaryFacts.Add(new PropertyEntry(name, FavoritePropertyCatalog.Format(name, value)));
        }
    }

    /// <summary>Delimited text of the properties currently shown, for the clipboard.</summary>
    public string GetVisiblePropertiesText()
    {
        var key = _settings.Data.CopyAllDelimiter;
        return string.Join(Environment.NewLine, VisibleProperties.Select(p => Delimiters.Join(key, p.Name, p.Source, p.Value)));
    }

    /// <summary>Every property of the selected result (unfiltered); null until loaded.</summary>
    public IReadOnlyList<PropertyEntry>? GetSelectedProperties() => SelectedResult?.AllProperties;

    #endregion
}
