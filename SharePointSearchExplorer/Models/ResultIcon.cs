using SharePointSearchExplorer.Helpers;

namespace SharePointSearchExplorer.Models;

/// <summary>
/// How a search result is drawn: an outline Fluent glyph, a short type label, and the name of the matching
/// colourful Fluent file type icon (Assets/FileTypes/{name}.svg), or null when there is none (e.g. people).
/// </summary>
public sealed record ResultIcon(string Glyph, string Label, string? FileTypeIcon = "genericfile")
{
    private static readonly Dictionary<string, ResultIcon> s_byExtension = new(StringComparer.OrdinalIgnoreCase);

    static ResultIcon()
    {
        void Map(string glyph, string label, string fileTypeIcon, params string[] extensions)
        {
            foreach (var e in extensions) s_byExtension[e] = new(glyph, label, fileTypeIcon);
        }

        Map(FluentIcons.DocumentText, "Word document", "docx", "doc", "docx", "docm", "odt");
        Map(FluentIcons.DocumentText, "Word template", "dotx", "dot", "dotx", "dotm");
        Map(FluentIcons.DocumentText, "Rich text document", "rtf", "rtf");
        Map(FluentIcons.DocumentTable, "Excel workbook", "xlsx", "xls", "xlsx", "xlsm", "xlsb", "ods");
        Map(FluentIcons.DocumentTable, "Excel template", "xltx", "xlt", "xltx", "xltm");
        Map(FluentIcons.Table, "CSV file", "csv", "csv", "tsv");
        Map(FluentIcons.SlideText, "PowerPoint presentation", "pptx", "ppt", "pptx", "pptm", "odp");
        Map(FluentIcons.SlideText, "PowerPoint slide show", "ppsx", "pps", "ppsx");
        Map(FluentIcons.SlideText, "PowerPoint template", "potx", "pot", "potx");
        Map(FluentIcons.Notebook, "OneNote", "one", "one", "onepkg");
        Map(FluentIcons.Notebook, "OneNote notebook", "onetoc", "onetoc2");
        Map(FluentIcons.DocumentPdf, "PDF document", "pdf", "pdf");
        Map(FluentIcons.Diagram, "Visio drawing", "vsdx", "vsd", "vsdx", "vsdm");
        Map(FluentIcons.Diagram, "Visio stencil", "vssx", "vss", "vssx");
        Map(FluentIcons.Diagram, "Visio template", "vstx", "vst", "vstx");
        Map(FluentIcons.GanttChart, "Project plan", "mpp", "mpp", "mpt");
        Map(FluentIcons.Form, "InfoPath form", "xsn", "xsn");
        Map(FluentIcons.Database, "Access database", "accdb", "accdb", "mdb");
        Map(FluentIcons.Document, "Publisher document", "pub", "pub");
        Map(FluentIcons.Notepad, "Text file", "txt", "txt", "log", "md", "markdown");
        Map(FluentIcons.Mail, "Email message", "email", "msg", "eml", "oft");
        Map(FluentIcons.DocumentOnePage, "Page", "html", "aspx", "asp");
        Map(FluentIcons.DocumentGlobe, "Web page", "html", "html", "htm", "mht", "mhtml");
        Map(FluentIcons.Image, "Image", "photo", "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "heic", "webp", "ico");
        Map(FluentIcons.Image, "Vector image", "vector", "svg", "ai", "eps", "emf", "wmf");
        Map(FluentIcons.Video, "Video", "video", "mp4", "mov", "avi", "wmv", "mkv", "webm", "m4v", "mpg", "mpeg");
        Map(FluentIcons.MusicNote2, "Audio", "audio", "mp3", "wav", "m4a", "wma", "aac", "flac", "ogg");
        Map(FluentIcons.FolderZip, "Compressed archive", "zip", "zip", "7z", "rar", "cab", "gz", "tar");
        Map(FluentIcons.Braces, "Data file", "xml", "xml", "config", "resx", "xaml");
        Map(FluentIcons.Braces, "Data file", "code", "json", "yaml", "yml");
        Map(FluentIcons.Code, "Code file", "code", "cs", "js", "ts", "jsx", "tsx", "ps1", "psm1", "py", "java", "css", "sql", "sh", "bat", "cmd", "vbs");
        Map(FluentIcons.AppGeneric, "Application", "exe", "exe", "msi", "appx", "msix");
        Map(FluentIcons.AppGeneric, "System file", "sysfile", "dll", "sys", "ini", "inf");
        Map(FluentIcons.TextDescription, "Font", "font", "ttf", "otf", "woff", "woff2");
        Map(FluentIcons.Diagram, "3D model", "model", "stl", "obj", "fbx", "3mf", "glb", "gltf");
        Map(FluentIcons.Whiteboard, "Whiteboard", "whiteboard", "whiteboard");
        Map(FluentIcons.DocumentLandscape, "Loop component", "loop", "loop", "fluid");
    }

    public static ResultIcon Default { get; } = new(FluentIcons.Document, "Document");

    /// <summary>Picks an icon from the result's container flag, file type and SharePoint content class.</summary>
    public static ResultIcon For(string fileType, string contentClass, bool isContainer, bool isDocument)
    {
        var cc = contentClass.ToLowerInvariant();

        if (cc.Contains("sppeople") || cc.Contains("person")) return new(FluentIcons.Person, "Person", FileTypeIcon: null);
        if (isContainer || (cc == "sts_list_documentlibrary" && !isDocument)) return new(FluentIcons.Folder, "Folder", "folder");
        if (s_byExtension.TryGetValue(fileType.Trim('.'), out var byType)) return byType;

        return cc switch
        {
            "sts_site" => new(FluentIcons.Globe, "Site", "spo"),
            "sts_web" => new(FluentIcons.Globe, "Subsite", "spo"),
            _ when cc.StartsWith("sts_list_") && cc.Contains("library") => new(FluentIcons.Library, "Library", "docset"),
            _ when cc.StartsWith("sts_list_") => new(FluentIcons.TextBulletListLtr, "List", "splist"),
            _ when cc.Contains("events") => new(FluentIcons.CalendarLtr, "Event", "calendar"),
            _ when cc.Contains("tasks") => new(FluentIcons.TaskListLtr, "Task", "todoitem"),
            _ when cc.Contains("announcements") => new(FluentIcons.Megaphone, "Announcement", "sponews"),
            _ when cc.Contains("contacts") => new(FluentIcons.ContactCard, "Contact", "contact"),
            _ when cc.Contains("discussion") => new(FluentIcons.Chat, "Discussion", "listitem"),
            _ when cc.Contains("links") => new(FluentIcons.Link, "Link", "link"),
            _ when cc.Contains("survey") => new(FluentIcons.Poll, "Survey", "form"),
            _ when cc.Contains("picturelibrary") => new(FluentIcons.ImageMultiple, "Picture", "photo"),
            _ when cc.Contains("webpagelibrary") || cc.Contains("publishingpages") => new(FluentIcons.DocumentOnePage, "Page", "html"),
            _ when cc.StartsWith("sts_listitem") && !isDocument => new(FluentIcons.TextBulletListSquare, "List item", "listitem"),
            _ => Default,
        };
    }
}
