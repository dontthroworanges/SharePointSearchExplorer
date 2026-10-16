using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Windows.System;
using Windows.UI.Core;

namespace SharePointSearchExplorer.Views;

/// <summary>Text search for a <see cref="CodeBlock"/>: a search box, the match count and previous / next buttons.</summary>
public sealed partial class FindBox : UserControl
{
    private CodeBlock? _target;

    public FindBox() => InitializeComponent();

    /// <summary>The code block to search. The search is run again whenever its text changes (e.g. another tab).</summary>
    public CodeBlock? Target
    {
        get => _target;
        set
        {
            if (_target != null) _target.TextChanged -= Target_TextChanged;
            _target = value;
            if (_target != null) _target.TextChanged += Target_TextChanged;
            Find();
        }
    }

    private void Target_TextChanged(object? sender, EventArgs e) => Find();

    private void QueryBox_TextChanged(object sender, TextChangedEventArgs e) => Find();

    private void Find()
    {
        _target?.Find(QueryBox.Text);
        UpdateCount();
    }

    private void FindNext(bool forward)
    {
        _target?.FindNext(forward);
        UpdateCount();
    }

    private void UpdateCount()
    {
        int count = _target?.MatchCount ?? 0;
        CountText.Text = QueryBox.Text.Length == 0 || _target is not { Text.Length: > 0 }
            ? ""
            : count == 0
                ? "No results"
                : $"{_target.CurrentMatch + 1:N0} of {count:N0}";
        PreviousButton.IsEnabled = NextButton.IsEnabled = count > 0;
    }

    private void QueryBox_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key != VirtualKey.Enter) return;
        bool shift = InputKeyboardSource.GetKeyStateForCurrentThread(VirtualKey.Shift).HasFlag(CoreVirtualKeyStates.Down);
        FindNext(forward: !shift);
        e.Handled = true;
    }

    private void Previous_Click(object sender, RoutedEventArgs e) => FindNext(forward: false);

    private void Next_Click(object sender, RoutedEventArgs e) => FindNext(forward: true);

    private void FocusAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        QueryBox.Focus(FocusState.Keyboard);
        QueryBox.SelectAll();
        args.Handled = true;
    }
}
