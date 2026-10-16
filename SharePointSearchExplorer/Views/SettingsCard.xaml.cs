using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Automation.Peers;
using Microsoft.UI.Xaml.Automation.Provider;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using SharePointSearchExplorer.Helpers;

namespace SharePointSearchExplorer.Views;

public sealed partial class SettingsCard : UserControl
{
    public static readonly DependencyProperty HeaderProperty =
        DependencyProperty.Register(nameof(Header), typeof(string), typeof(SettingsCard), new PropertyMetadata(""));

    public static readonly DependencyProperty DescriptionProperty =
        DependencyProperty.Register(nameof(Description), typeof(string), typeof(SettingsCard), new PropertyMetadata(""));

    public static readonly DependencyProperty IconProperty =
        DependencyProperty.Register(nameof(Icon), typeof(string), typeof(SettingsCard),
            new PropertyMetadata("", (d, e) => ((SettingsCard)d).Glyph = FluentIcons.Glyph((string)e.NewValue)));

    public static readonly DependencyProperty GlyphProperty =
        DependencyProperty.Register(nameof(Glyph), typeof(string), typeof(SettingsCard), new PropertyMetadata(""));

    public static readonly DependencyProperty ActionProperty =
        DependencyProperty.Register(nameof(Action), typeof(object), typeof(SettingsCard),
            new PropertyMetadata(null, (d, e) => ((SettingsCard)d).OnActionChanged(e.NewValue)));

    /// <summary>Below this much room for the header and description, the control moves under them.</summary>
    private const double MinTextWidth = 240;

    private bool _stacked;
    private ToggleSwitch? _toggle;
    private long _toggleCallback;
    private string _onText = "On";
    private string _offText = "Off";

    public static readonly DependencyProperty IsClickEnabledProperty =
        DependencyProperty.Register(nameof(IsClickEnabled), typeof(bool), typeof(SettingsCard),
            new PropertyMetadata(false, (d, e) => ((SettingsCard)d).OnIsClickEnabledChanged((bool)e.NewValue)));

    public SettingsCard()
    {
        InitializeComponent();
        // Disabling the card also disables its control (IsEnabled is inherited); the state greys the rest.
        IsEnabledChanged += (_, _) => VisualStateManager.GoToState(this, IsEnabled ? "Normal" : "Disabled", false);
    }

    /// <summary>Raised when a clickable card (see <see cref="IsClickEnabled"/>) is clicked, tapped or activated with Enter or Space.</summary>
    public event RoutedEventHandler? Click;

    /// <summary>
    /// The whole card is a button that opens a sub-page, with a chevron on its right edge, like the rows of
    /// Windows Settings that lead to another page.
    /// </summary>
    public bool IsClickEnabled
    {
        get => (bool)GetValue(IsClickEnabledProperty);
        set => SetValue(IsClickEnabledProperty, value);
    }

    private void OnIsClickEnabledChanged(bool enabled)
    {
        IsTabStop = enabled;
        UseSystemFocusVisuals = enabled;
        // Selectable text takes the pointer, so on a clickable card the description must be clickable too.
        DescriptionText.IsTextSelectionEnabled = !enabled;
        if (enabled) AutomationProperties.SetName(this, Header);
        VisualStateManager.GoToState(this, "ClickNormal", false);
    }

    private void GoToClickState(string state)
    {
        if (IsClickEnabled) VisualStateManager.GoToState(this, state, false);
    }

    protected override void OnPointerEntered(PointerRoutedEventArgs e) => GoToClickState("ClickPointerOver");

    protected override void OnPointerExited(PointerRoutedEventArgs e) => GoToClickState("ClickNormal");

    protected override void OnPointerCanceled(PointerRoutedEventArgs e) => GoToClickState("ClickNormal");

    protected override void OnPointerPressed(PointerRoutedEventArgs e) => GoToClickState("ClickPressed");

    protected override void OnPointerReleased(PointerRoutedEventArgs e) => GoToClickState("ClickPointerOver");

    protected override void OnTapped(TappedRoutedEventArgs e)
    {
        if (!IsClickEnabled) return;
        e.Handled = true;
        Click?.Invoke(this, new RoutedEventArgs());
    }

    protected override AutomationPeer OnCreateAutomationPeer() => new CardAutomationPeer(this);

    /// <summary>Presents a clickable card to screen readers as a button that can be invoked.</summary>
    private sealed partial class CardAutomationPeer(SettingsCard owner) : FrameworkElementAutomationPeer(owner), IInvokeProvider
    {
        protected override AutomationControlType GetAutomationControlTypeCore() =>
            owner.IsClickEnabled ? AutomationControlType.Button : base.GetAutomationControlTypeCore();

        protected override object GetPatternCore(PatternInterface patternInterface) =>
            owner.IsClickEnabled && patternInterface == PatternInterface.Invoke ? this : base.GetPatternCore(patternInterface);

        public void Invoke() => owner.Click?.Invoke(owner, new RoutedEventArgs());
    }

    protected override void OnKeyDown(KeyRoutedEventArgs e)
    {
        if (!IsClickEnabled || e.Key is not (Windows.System.VirtualKey.Enter or Windows.System.VirtualKey.Space)) return;
        e.Handled = true;
        Click?.Invoke(this, new RoutedEventArgs());
    }

    public string Header
    {
        get => (string)GetValue(HeaderProperty);
        set => SetValue(HeaderProperty, value);
    }

    public string Description
    {
        get => (string)GetValue(DescriptionProperty);
        set => SetValue(DescriptionProperty, value);
    }

    /// <summary>Fluent icon name, e.g. "Search" (see <see cref="FluentIcons"/>).</summary>
    public string Icon
    {
        get => (string)GetValue(IconProperty);
        set => SetValue(IconProperty, value);
    }

    public string Glyph
    {
        get => (string)GetValue(GlyphProperty);
        private set => SetValue(GlyphProperty, value);
    }

    /// <summary>The control shown on the right (toggle, box, button…).</summary>
    public object? Action
    {
        get => GetValue(ActionProperty);
        set => SetValue(ActionProperty, value);
    }

    public Visibility HasDescription(string? text) => string.IsNullOrEmpty(text) ? Visibility.Collapsed : Visibility.Visible;

    /// <summary>
    /// Like Windows Settings, a toggle switch's On/Off text goes to the left of the switch so the switch sits on the
    /// card's right edge, lined up with the other controls. The switch's own text (and the width kept for it) is removed.
    /// </summary>
    private void OnActionChanged(object? action)
    {
        if (_toggle != null)
        {
            _toggle.UnregisterPropertyChangedCallback(ToggleSwitch.IsOnProperty, _toggleCallback);
            _toggle = null;
        }
        if (action is ToggleSwitch toggle)
        {
            _onText = toggle.OnContent as string ?? "On";
            _offText = toggle.OffContent as string ?? "Off";
            toggle.OnContent = null;
            toggle.OffContent = null;
            toggle.MinWidth = 0;
            _toggle = toggle;
            _toggleCallback = toggle.RegisterPropertyChangedCallback(ToggleSwitch.IsOnProperty, (_, _) => UpdateToggleText());
        }
        UpdateToggleText();
    }

    private void UpdateToggleText()
    {
        ToggleStateText.Visibility = _toggle == null ? Visibility.Collapsed : Visibility.Visible;
        if (_toggle != null) ToggleStateText.Text = _toggle.IsOn ? _onText : _offText;
    }

    private void CardRoot_SizeChanged(object sender, SizeChangedEventArgs e) => UpdateLayoutMode(e.NewSize.Width);

    /// <summary>
    /// Side by side while the header and description keep at least <see cref="MinTextWidth"/>; otherwise the control
    /// moves under the text, left-aligned, as Windows Settings does in a narrow window.
    /// </summary>
    private void UpdateLayoutMode(double cardWidth)
    {
        double actionWidth = ActionPanel.DesiredSize.Width;
        var padding = CardRoot.Padding;
        var border = CardRoot.BorderThickness;
        double chrome = padding.Left + padding.Right + border.Left + border.Right + CardIcon.ActualWidth + 2 * CardRoot.ColumnSpacing;
        bool stack = actionWidth > 0 && cardWidth - chrome - actionWidth < MinTextWidth;
        if (stack == _stacked) return;

        _stacked = stack;
        Grid.SetRow(ActionPanel, stack ? 1 : 0);
        Grid.SetColumn(ActionPanel, stack ? 1 : 2);
        ActionPanel.HorizontalAlignment = stack ? HorizontalAlignment.Left : HorizontalAlignment.Right;
        ActionPanel.Margin = stack ? new Thickness(0, 12, 0, 0) : new Thickness(0);
    }
}
