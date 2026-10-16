using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointSearchExplorer.Models;

/// <summary>
/// One refiner entry that can be ticked in the filter flyout: a file type, or the content class of results
/// that have no file type (e.g. STS_ListItem_GenericList).
/// </summary>
public partial class RefinerOption : ObservableObject
{
    public RefinerOption(string name, string token, long count, string refiner = "FileType")
    {
        Name = name;
        Token = token;
        Count = count;
        Refiner = refiner;
    }

    /// <summary>The managed property the entry came from (FileType or ContentClass).</summary>
    public string Refiner { get; }
    public string Name { get; }
    public string Token { get; }
    public long Count { get; }
    public string Label => $"{(string.IsNullOrEmpty(Name) ? "(none)" : Name)} ({Count:N0})";

    [ObservableProperty]
    private bool _isSelected;
}
