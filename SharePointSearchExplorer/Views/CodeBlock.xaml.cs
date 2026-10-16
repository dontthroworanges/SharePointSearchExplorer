using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media;
using Windows.UI;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// A read-only code block with text search, shared by the Server Response page and its pop-out windows.
/// </summary>
public sealed partial class CodeBlock : UserControl
{
    /// <summary>Navigation covers every match, but only this many are highlighted at once, to keep long texts responsive.</summary>
    private const int MaxHighlightedMatches = 2000;

    // Every match in translucent yellow (readable in both themes); the current one in solid orange.
    private static readonly Color s_matchColor = Color.FromArgb(0x66, 0xFF, 0xD3, 0x3D);
    private static readonly Color s_currentColor = Color.FromArgb(0xFF, 0xFF, 0x8C, 0x00);

    private readonly List<int> _matches = new();
    private int _matchLength;
    private string _text = "";

    // The text as displayed: line breaks are single characters, so match positions equal highlight positions.
    private string _shown = "";

    public CodeBlock() => InitializeComponent();

    /// <summary>Raised after <see cref="Text"/> is set; any search has to be run again.</summary>
    public event EventHandler? TextChanged;

    /// <summary>Shown in place of the code while <see cref="Text"/> is empty.</summary>
    public string PlaceholderText
    {
        get => Placeholder.Text;
        set => Placeholder.Text = value;
    }

    public string Text
    {
        get => _text;
        set
        {
            _text = value ?? "";
            _shown = _text.Replace("\r\n", "\n");
            _matches.Clear();
            CurrentMatch = -1;
            UpdateHighlights();
            CodeText.Text = _shown;
            Scroller.Visibility = _text.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
            Placeholder.Visibility = _text.Length > 0 ? Visibility.Collapsed : Visibility.Visible;
            Scroller.ChangeView(0, 0, null, disableAnimation: true);
            TextChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    #region Search

    public int MatchCount => _matches.Count;

    /// <summary>Index of the selected match (0-based); -1 when there are none.</summary>
    public int CurrentMatch { get; private set; } = -1;

    /// <summary>
    /// Highlights every occurrence of <paramref name="query"/> (ignoring case) and goes to the first. An empty
    /// query clears the highlights.
    /// </summary>
    public void Find(string query)
    {
        _matches.Clear();
        CurrentMatch = -1;

        if (!string.IsNullOrEmpty(query))
        {
            _matchLength = query.Length;
            for (int i = _shown.IndexOf(query, StringComparison.OrdinalIgnoreCase); i >= 0;
                 i = _shown.IndexOf(query, i + query.Length, StringComparison.OrdinalIgnoreCase))
            {
                _matches.Add(i);
            }
            if (_matches.Count > 0) CurrentMatch = 0;
        }

        UpdateHighlights();
        ScrollToCurrentMatch();
    }

    /// <summary>Goes to the next (or previous) match, wrapping around at the ends.</summary>
    public void FindNext(bool forward)
    {
        if (_matches.Count == 0) return;
        CurrentMatch = (CurrentMatch + (forward ? 1 : -1) + _matches.Count) % _matches.Count;
        UpdateHighlights();
        ScrollToCurrentMatch();
    }

    /// <summary>
    /// Replaces the highlighters with new ones for the current matches. They are rebuilt rather than edited
    /// because the TextBlock doesn't repaint when the ranges of a highlighter it already has are cleared.
    /// </summary>
    private void UpdateHighlights()
    {
        CodeText.TextHighlighters.Clear();
        if (_matches.Count == 0) return;

        var all = new TextHighlighter { Background = new SolidColorBrush(s_matchColor) };
        foreach (int start in _matches.Take(MaxHighlightedMatches))
            all.Ranges.Add(new TextRange { StartIndex = start, Length = _matchLength });

        var current = new TextHighlighter
        {
            Background = new SolidColorBrush(s_currentColor),
            Foreground = new SolidColorBrush(Colors.Black),
        };
        current.Ranges.Add(new TextRange { StartIndex = _matches[CurrentMatch], Length = _matchLength });

        CodeText.TextHighlighters.Add(all);
        CodeText.TextHighlighters.Add(current);
    }

    private void ScrollToCurrentMatch(bool retry = true)
    {
        if (CurrentMatch < 0) return;
        int start = _matches[CurrentMatch];

        // A scroll requested while the text or its highlights are still being laid out can be dropped, so the
        // position is checked once more after this pass (a no-op when the match is already in view).
        if (retry)
        {
            DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
            {
                if (CurrentMatch >= 0 && _matches[CurrentMatch] == start) ScrollToCurrentMatch(retry: false);
            });
        }

        // New text may not be laid out yet; the match's position is needed to scroll to it.
        Scroller.UpdateLayout();
        if (CodeText.ContentStart.GetPositionAtOffset(start, LogicalDirection.Forward) is not { } pointer) return;
        var rect = pointer.GetCharacterRect(LogicalDirection.Forward);
        var end = CodeText.ContentStart.GetPositionAtOffset(start + _matchLength, LogicalDirection.Forward)?.GetCharacterRect(LogicalDirection.Forward);
        // A match that runs onto the next line ends left of where it starts; then only its start is kept in view.
        double right = end is { } e && e.Left > rect.Left ? e.Left : rect.Right;

        // Scroll only when the match is outside the visible area, and then bring it a third of the way in.
        double x = Scroller.HorizontalOffset, y = Scroller.VerticalOffset;
        if (rect.Top < y || rect.Bottom > y + Scroller.ViewportHeight)
            y = Math.Max(0, rect.Top - Scroller.ViewportHeight / 3);
        if (rect.Left < x || right > x + Scroller.ViewportWidth)
            x = Math.Max(0, rect.Left - Scroller.ViewportWidth / 3);
        Scroller.ChangeView(x, y, null, disableAnimation: true);
    }

    #endregion
}
