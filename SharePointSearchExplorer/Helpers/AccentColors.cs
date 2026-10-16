using System.Globalization;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using Windows.UI;
using Windows.UI.ViewManagement;

namespace SharePointSearchExplorer.Helpers;

/// <summary>A colour from the Windows 11 Settings › Personalization › Colors palette.</summary>
public sealed class AccentSwatch(string name, string hex)
{
    public string Name { get; } = name;
    public string Hex { get; } = hex;
    public Color Color { get; } = AccentColors.Parse(hex)!.Value;

    private SolidColorBrush? _brush;
    public SolidColorBrush Brush => _brush ??= new SolidColorBrush(Color);
}

/// <summary>
/// The app's accent colour: either the Windows accent colour (the default) or one of the standard Windows
/// colours. SystemAccentColor and its light and dark shades, which every accent brush in WinUI is built from,
/// are always set in the app resources: to the Windows values or to shades of the chosen colour.
/// </summary>
/// <remarks>
/// The entries must exist from startup. Theme resources only pick up a change to an entry that was already in
/// the app resources, so adding them later (the first switch from the Windows colour to a custom one) had no
/// effect until the app was restarted.
/// </remarks>
public static class AccentColors
{
    /// <summary>The 48 "Windows colors" swatches, in the order Windows Settings shows them (8 per row).</summary>
    public static IReadOnlyList<AccentSwatch> Palette { get; } = new AccentSwatch[]
    {
        new("Yellow gold", "#FFB900"), new("Gold", "#FF8C00"), new("Orange bright", "#F7630C"), new("Orange dark", "#CA5010"),
        new("Rust", "#DA3B01"), new("Pale rust", "#EF6950"), new("Brick red", "#D13438"), new("Mod red", "#FF4343"),
        new("Pale red", "#E74856"), new("Red", "#E81123"), new("Rose bright", "#EA005E"), new("Rose", "#C30052"),
        new("Plum light", "#E3008C"), new("Plum", "#BF0077"), new("Orchid light", "#C239B3"), new("Orchid", "#9A0089"),
        new("Navy blue", "#0078D4"), new("Navy blue dark", "#0063B1"), new("Purple shadow", "#8E8CD8"), new("Purple shadow dark", "#6B69D6"),
        new("Iris pastel", "#8764B8"), new("Iris spring", "#744DA9"), new("Violet red light", "#B146C2"), new("Violet red", "#881798"),
        new("Cool blue bright", "#0099BC"), new("Cool blue", "#2D7D9A"), new("Seafoam", "#00B7C3"), new("Seafoam teal", "#038387"),
        new("Mint light", "#00B294"), new("Mint dark", "#018574"), new("Turf green", "#00CC6A"), new("Sport green", "#10893E"),
        new("Gray", "#7A7574"), new("Gray brown", "#5D5A58"), new("Steel blue", "#68768A"), new("Metal blue", "#515C6B"),
        new("Pale moss", "#567C73"), new("Moss", "#486860"), new("Meadow green", "#498205"), new("Green", "#107C10"),
        new("Overcast", "#767676"), new("Storm", "#4C4A48"), new("Blue gray", "#69797E"), new("Gray dark", "#4A5459"),
        new("Liddy green", "#647C64"), new("Sage", "#525E54"), new("Camouflage desert", "#847545"), new("Camouflage", "#7E735F"),
    };

    private const string Base = "SystemAccentColor";

    /// <summary>Resource key and how far to blend towards white (positive) or black (negative).</summary>
    private static readonly (string Key, double Amount)[] s_shades =
    {
        (Base, 0),
        (Base + "Light1", 0.25), (Base + "Light2", 0.5), (Base + "Light3", 0.7),
        (Base + "Dark1", -0.2), (Base + "Dark2", -0.4), (Base + "Dark3", -0.6),
    };

    /// <summary>The same resource keys with the Windows accent colour values they take when no colour is chosen.</summary>
    private static readonly (string Key, UIColorType Type)[] s_systemShades =
    {
        (Base, UIColorType.Accent),
        (Base + "Light1", UIColorType.AccentLight1), (Base + "Light2", UIColorType.AccentLight2), (Base + "Light3", UIColorType.AccentLight3),
        (Base + "Dark1", UIColorType.AccentDark1), (Base + "Dark2", UIColorType.AccentDark2), (Base + "Dark3", UIColorType.AccentDark3),
    };

    // Kept alive for ColorValuesChanged.
    private static readonly UISettings s_uiSettings = new();
    private static Color s_lastSystemAccent = s_uiSettings.GetColorValue(UIColorType.Accent);

    static AccentColors() => s_uiSettings.ColorValuesChanged += (_, _) =>
    {
        // Also raised for light/dark changes; only pass on a new accent colour.
        var accent = s_uiSettings.GetColorValue(UIColorType.Accent);
        if (accent == s_lastSystemAccent) return;
        s_lastSystemAccent = accent;
        SystemAccentChanged?.Invoke(null, EventArgs.Empty);
    };

    /// <summary>The Windows accent colour changed. Raised on a background thread.</summary>
    public static event EventHandler? SystemAccentChanged;

    /// <summary>True while the app follows the Windows accent colour (no custom colour applied).</summary>
    public static bool UsesSystemAccent { get; private set; } = true;

    /// <summary>"#RRGGBB" to a colour, or null for empty/invalid text (meaning "use the Windows accent colour").</summary>
    public static Color? Parse(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex)) return null;
        var text = hex.Trim().TrimStart('#');
        if (text.Length != 6 || !uint.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)) return null;
        return Color.FromArgb(0xFF, (byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
    }

    /// <summary>
    /// Sets the accent resources to shades of <paramref name="hex"/>, or to the Windows accent colours when it is
    /// empty. Call at startup before the window is created, and after that follow with <see cref="Refresh"/>.
    /// </summary>
    public static void Apply(string? hex)
    {
        var resources = Application.Current.Resources;
        if (Parse(hex) is { } color)
        {
            UsesSystemAccent = false;
            foreach (var (key, amount) in s_shades) resources[key] = Blend(color, amount);
        }
        else
        {
            UsesSystemAccent = true;
            foreach (var (key, type) in s_systemShades) resources[key] = s_uiSettings.GetColorValue(type);
        }
    }

    /// <summary>
    /// Theme resources are only looked up again when the theme changes, so flip the element's theme and back to
    /// make the accent brushes pick up the new colour.
    /// </summary>
    public static void Refresh(FrameworkElement root)
    {
        var requested = root.RequestedTheme;
        root.RequestedTheme = root.ActualTheme == ElementTheme.Dark ? ElementTheme.Light : ElementTheme.Dark;
        // Switching straight back is coalesced into no change, so restore once the flip has been applied.
        root.DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () => root.RequestedTheme = requested);
    }

    private static Color Blend(Color color, double amount)
    {
        if (amount == 0) return color;
        var target = amount > 0 ? Colors.White : Colors.Black;
        double t = Math.Abs(amount);
        return Color.FromArgb(0xFF, Mix(color.R, target.R), Mix(color.G, target.G), Mix(color.B, target.B));

        byte Mix(byte from, byte to) => (byte)Math.Round(from + (to - from) * t);
    }
}
