using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;

namespace SharePointSearchExplorer.Helpers;

/// <summary>
/// Live display preferences that templates bind to with
/// <c>{x:Bind h:DisplayOptions.Current.Wrapping, Mode=OneWay}</c>, so a change in
/// Settings updates the results list and property tables immediately.
/// </summary>
public sealed partial class DisplayOptions : ObservableObject
{
    public static DisplayOptions Current { get; } = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Wrapping))]
    [NotifyPropertyChangedFor(nameof(Trimming))]
    [NotifyPropertyChangedFor(nameof(TitleMaxLines))]
    [NotifyPropertyChangedFor(nameof(CellMaxLines))]
    private bool _wordWrap = true;

    /// <summary>Wrap long lines in the Server Response code blocks instead of scrolling them sideways.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CodeWrapping))]
    [NotifyPropertyChangedFor(nameof(CodeHorizontalScrollBar))]
    [NotifyPropertyChangedFor(nameof(CodeHorizontalScrollMode))]
    private bool _wrapCodeBlocks;

    public TextWrapping CodeWrapping => WrapCodeBlocks ? TextWrapping.Wrap : TextWrapping.NoWrap;

    // Wrapped code is laid out at the block's width, so sideways scrolling is switched off.
    public Microsoft.UI.Xaml.Controls.ScrollBarVisibility CodeHorizontalScrollBar =>
        WrapCodeBlocks ? Microsoft.UI.Xaml.Controls.ScrollBarVisibility.Disabled : Microsoft.UI.Xaml.Controls.ScrollBarVisibility.Auto;

    public Microsoft.UI.Xaml.Controls.ScrollMode CodeHorizontalScrollMode =>
        WrapCodeBlocks ? Microsoft.UI.Xaml.Controls.ScrollMode.Disabled : Microsoft.UI.Xaml.Controls.ScrollMode.Auto;

    /// <summary>Colourful Fluent file type icons in the results list instead of outline glyphs.</summary>
    [ObservableProperty]
    private bool _useFileTypeIcons;

    /// <summary>Dense primary results: no last modified line, and tighter rows.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(LastModifiedVisibility))]
    [NotifyPropertyChangedFor(nameof(ResultPadding))]
    [NotifyPropertyChangedFor(nameof(ResultColumnSpacing))]
    private bool _denseResults;

    /// <summary>The last modified line under each result title, hidden when dense.</summary>
    public Visibility LastModifiedVisibility => DenseResults ? Visibility.Collapsed : Visibility.Visible;

    /// <summary>
    /// The result icon's box (and file type image) size: 20px in both densities, the same as the Result Details
    /// card. The outline glyphs are the font's 20px icons, so they're only pixel-aligned (crisp) at 20px.
    /// </summary>
    public double ResultIconSize => 20;

    /// <summary>The outline glyph size inside the icon box.</summary>
    public double ResultGlyphSize => 20;

    public Thickness ResultPadding => DenseResults ? new Thickness(4, 8, 4, 8) : new Thickness(4, 10, 4, 10);

    public double ResultColumnSpacing => DenseResults ? 12 : 14;

    public TextWrapping Wrapping => WordWrap ? TextWrapping.Wrap : TextWrapping.NoWrap;

    public TextTrimming Trimming => WordWrap ? TextTrimming.None : TextTrimming.CharacterEllipsis;

    /// <summary>Result titles: unlimited lines when wrapping, one trimmed line otherwise.</summary>
    public int TitleMaxLines => WordWrap ? 0 : 1;

    /// <summary>
    /// Property table cells: unlimited lines when wrapping, one trimmed line otherwise. Without the limit a
    /// value with line breaks of its own (a multi-value property) would still take several lines.
    /// </summary>
    public int CellMaxLines => WordWrap ? 0 : 1;
}
