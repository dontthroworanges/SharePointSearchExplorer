using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;

namespace SharePointSearchExplorer.Helpers;

public static partial class FluentIcons
{
    public const string RegularFontUri = "ms-appx:///Assets/Fonts/FluentSystemIcons-Regular.ttf#FluentSystemIcons-Regular";
    public const string FilledFontUri = "ms-appx:///Assets/Fonts/FluentSystemIcons-Filled.ttf#FluentSystemIcons-Filled";

    private static FontFamily? s_regular;
    private static FontFamily? s_filled;

    public static FontFamily RegularFontFamily => s_regular ??= new FontFamily(RegularFontUri);
    public static FontFamily FilledFontFamily => s_filled ??= new FontFamily(FilledFontUri);

    public static string Glyph(string name) =>
        ByName.TryGetValue(name, out var glyph) ? glyph : throw new ArgumentException($"Unknown Fluent icon '{name}'.", nameof(name));

    public static FontIcon Create(string glyph, double fontSize = double.NaN, bool filled = false)
    {
        var icon = new FontIcon { Glyph = glyph, FontFamily = filled ? FilledFontFamily : RegularFontFamily };
        if (!double.IsNaN(fontSize)) icon.FontSize = fontSize;
        return icon;
    }

    public static FontIconSource CreateSource(string glyph, double fontSize = double.NaN, bool filled = false)
    {
        var icon = new FontIconSource { Glyph = glyph, FontFamily = filled ? FilledFontFamily : RegularFontFamily };
        if (!double.IsNaN(fontSize)) icon.FontSize = fontSize;
        return icon;
    }
}

/// <summary><c>Icon="{h:FluentIcon Icon=Search}"</c> — a FontIcon from Fluent UI System Icons.</summary>
[MarkupExtensionReturnType(ReturnType = typeof(FontIcon))]
public sealed partial class FluentIconExtension : MarkupExtension
{
    public string Icon { get; set; } = "";
    public double FontSize { get; set; } = double.NaN;

    protected override object ProvideValue() => FluentIcons.Create(FluentIcons.Glyph(Icon), FontSize);
}

/// <summary><c>IconSource="{h:FluentIconSource Icon=Info}"</c> — a FontIconSource from Fluent UI System Icons.</summary>
[MarkupExtensionReturnType(ReturnType = typeof(FontIconSource))]
public sealed partial class FluentIconSourceExtension : MarkupExtension
{
    public string Icon { get; set; } = "";
    public double FontSize { get; set; } = double.NaN;

    protected override object ProvideValue() => FluentIcons.CreateSource(FluentIcons.Glyph(Icon), FontSize);
}
