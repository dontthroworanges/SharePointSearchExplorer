using System.Globalization;
using System.Text.RegularExpressions;

namespace SharePointSearchExplorer.Models;

/// <summary>
/// The managed properties that can be shown in the Result Details summary (Settings › Search › Favorite
/// properties), and how their values are shown there.
/// </summary>
public static partial class FavoritePropertyCatalog
{
    /// <summary>The number of slots in the summary.</summary>
    public const int MaxCount = 12;

    /// <summary>The favorites of a new installation, and of "Reset to Default".</summary>
    public static readonly string[] Defaults =
    [
        "FileType", "Author", "ModifiedBy", "LastModifiedTime", "Created", "Size",
        "SiteTitle", "ContentClass", "WorkId", "Rank",
    ];

    /// <summary>Standard SharePoint managed properties offered as favorites, sorted by name.</summary>
    public static readonly IReadOnlyList<string> Standard = new[]
    {
        // Documents and items
        "Title", "Path", "OriginalPath", "Filename", "FileType", "FileExtension", "SecondaryFileExtension",
        "Author", "DisplayAuthor", "AuthorOWSUSER", "CreatedBy", "ModifiedBy", "EditorOWSUSER", "Created",
        "LastModifiedTime", "Write", "Size", "Description", "HitHighlightedSummary", "HitHighlightedProperties",
        "ContentClass", "ContentType", "ContentTypeId", "ContentSource", "IsDocument", "IsContainer",
        "IsExternalContent", "IsListItem", "language", "DetectedLanguage", "ParentLink", "LinkingUrl",
        "ServerRedirectedURL", "ServerRedirectedEmbedURL", "ServerRedirectedPreviewURL", "PictureThumbnailURL",
        "DefaultEncodingURL", "UniqueId", "ListId", "ListItemID", "ListUrl", "WebId", "SiteId", "DocId", "WorkId",
        "Rank", "ViewsLifeTime", "ViewsRecent", "ViewsLifeTimeUniqueUsers", "ViewsRecentUniqueUsers",
        "LastSharedByUser", "SharedWithUsersOWSUSER", "CheckoutUserOWSUSER", "ComplianceTag",
        "InformationProtectionLabelId", "SensitivityLabel", "FirstPublishedDate", "PromotedState", "Tags",
        "owstaxIdMetadataAllTagsInfo", "SectionNames", "MediaDuration", "PeopleInMedia", "importance",
        // Sites
        "SiteName", "SiteTitle", "SiteDescription", "SiteLogo", "SiteTemplate", "SiteTemplateId", "WebTemplate",
        "SPSiteUrl", "SPWebUrl", "DepartmentId", "IsHubSite", "RelatedHubSites", "GroupId", "SiteClosed",
        "UrlZone", "Culture", "GeoLocationSource",
        // Search internals
        "PartitionId", "ResultTypeId", "RenderTemplateId", "piSearchResultId", "CollapsingStatus",
        "NormSiteID", "NormWebID", "NormListID", "NormUniqueID", "DocumentSignature", "contentclass",
        // People
        "AccountName", "PreferredName", "FirstName", "LastName", "JobTitle", "Department", "WorkEmail",
        "WorkPhone", "MobilePhone", "OfficeNumber", "BaseOfficeLocation", "AboutMe", "PictureURL", "Skills",
        "Interests", "PastProjects", "Responsibilities", "Schools", "SipAddress", "UserProfile_GUID",
    }
    .Distinct(StringComparer.OrdinalIgnoreCase)
    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
    .ToList();

    /// <summary>
    /// A favorite's value as the summary shows it: timestamps in local time, Size in KB/MB, and a space after
    /// each ";" between authors. Anything else is shown as SharePoint returned it.
    /// </summary>
    public static string Format(string name, string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        if (IsoTimestamp().IsMatch(value)
            && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date))
            return date.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
        if (name.Equals("Size", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var size) && size > 0)
            return FormatSize(size);
        if (name.Equals("Author", StringComparison.OrdinalIgnoreCase)) return value.Replace(";", "; ");
        return value;
    }

    private static string FormatSize(double bytes)
    {
        string[] units = ["bytes", "KB", "MB", "GB", "TB"];
        int unit = 0;
        while (bytes >= 1024 && unit < units.Length - 1)
        {
            bytes /= 1024;
            unit++;
        }
        return unit == 0 ? $"{bytes:N0} bytes" : $"{bytes:0.#} {units[unit]}";
    }

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}")]
    private static partial Regex IsoTimestamp();
}
