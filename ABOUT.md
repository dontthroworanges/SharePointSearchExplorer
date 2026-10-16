# SharePoint Search Explorer

A WinUI 3 app inspired by the original PnP [SharePoint Search Query Tool](https://github.com/pnp/PnP-Tools/tree/master/Solutions/SharePoint.Search.QueryTool). Type a query, browse the primary results, and inspect each result's details, managed properties and crawled properties.

## What it does

| Area | Behaviour |
|---|---|
| **Query bar** (title bar) | KQL query box with recent-query suggestions. Point at a suggestion and click its **X** to remove it from the history, with a confirmation first (both can be turned off under **Settings › Search › Query box**). Clicking the X at the end of the query box clears the text and Primary Results. The accent button runs the search and becomes **Cancel** while a search is running. An empty query searches `*`. |
| **Primary Results** | Type icon, title and last-modified date. Icons cover Office files, PDFs, pages, images, video, audio, archives, code, email, folders, sites, lists, events, tasks, contacts and people; hover an icon for its type. The status line shows the total hit count and elapsed time. **Load more results** pages through the results. |
| **Filter** | File-type refiner (`FileType`) with counts. Tick one or more types and click **Apply**. A dot on the icon means a filter is active. |
| **Sort** | Relevance, last modified (newest or oldest), or size (largest or smallest). |
| **Result Details** | A summary card with the title, **Result Options** and your favorite properties (chosen under **Settings › Search › Favorite properties**; right-click one to copy it), then one table of every managed and crawled property found on the item, with a text filter, **Show/Hide empty**, **Copy all**, **Reload** and **Pop-Out**. Click the **Property** header to reverse the sort (A to Z or Z to A), and the **Source** header to show only managed or only crawled properties. Managed rows show their value; crawled rows show their property set, because search doesn't return crawled property values. Right-click a row to copy its name, its value, or both. The lists come from the same `ManagedProperties` and `CrawledProperties` refiner technique as the original tool. **Result Options** (the button beside the title in Result Details, and the right-click menu of a result in Primary Results) has **Open Result** (the SharePoint viewer: `SPWebUrl/_layouts/15/viewer.aspx?sourcedoc=UniqueId`, or the item's `Path` when a result has no `SPWebUrl` or `UniqueId`), **Open Parent Location** (the `ParentLink` managed property; unavailable when it's empty) and **Copy link** (the `Path`). |
| **Advanced Query Rules** (left rail) | The original tool's Query and Suggest settings, side by side, under the request URL (the **Query URL** / **Suggest URL** tabs pick which request it shows). The query text is the query box in the title bar, and its search button runs the request while this page is open. The true/false settings are one multi-select **Options** drop-down: checked sends `true`, cleared sends `false`, a dash leaves the setting out. Fields with an example or a list of known values (select properties, refiners, culture, result source, client type…) offer them as suggestions. **Save query** adds the request to the Query History page (stored without the site, as `_api/search/query?…`), where **Run again** opens it on this page and runs it. The request URL can be edited directly; the fields and the URL update each other. **Run** sends the URL as shown: query results open on the Search Results page, suggestions are listed on the page, and the full response is on the Server Response page. |
| **Profile** (left rail) | Switch between saved site profiles; add, edit or remove profiles; sign in with a different account; sign out. A green check on the icon means the profile is signed in; a yellow alert means it isn't yet. New profiles are named after the tenant in the URL (e.g. `contoso`), which you can change. |
| **Settings** (left rail) | Laid out like Windows 11 Settings: categories on the left (Search, Accounts, Appearance, Query History, About), setting cards on the right. **Search** starts with **Advanced pages** (off by default), which shows or hides the Advanced Query Rules and Server Response pages in the left rail and enables the settings that only apply to them, and includes **Favorite properties** (a sub-page for choosing up to 12 standard managed properties, in order, for the summary at the top of Result Details; searches request them alongside the built-in properties), the **Result Details Table Default Property Filter** (whether the Result Details table starts with all, only managed or only crawled properties) and the delimiters used by right-click **Copy name and value** and by **Copy all** (comma, semicolon, colon, pipe or space; comma/semicolon output is CSV-quoted). **Appearance** has the theme, the **Accent colour** (your Windows accent colour, or **Custom** with the 48 standard Windows 11 personalization colours), two **Word Wrap** switches (one for the Search Results page, one for the Server Response code blocks), and **Result icons**: **Outline** (Fluent System Icons) or **Colourful** (Microsoft's file type icons). |
| **Server Response** (left rail) | The HTTP request and response of the last search (the original tool's Status, Raw and Debug tabs). **HTTP Details** shows the request, status and query statistics; **Full API Response** shows the headers and the indented JSON body; **Debug** shows the query parameters sent, the query as SharePoint rewrote it, the correlation id and the other properties and diagnostic headers SharePoint returned. The search box (Ctrl+F) highlights every match in the current tab; Enter and Shift+Enter move between them. **Copy** copies the text and **Pop-Out** opens it in its own window, which has its own search box. The access token is redacted. |
| **Query History** (left rail) | Recent queries with **Run again**, **Export** and **Remove** on each. **Clear history** always asks for confirmation. **Export all…** saves the whole history to an XML file and **Import…** adds queries from one (see [Saved query files](#saved-query-files)). |

The left rail is a standard WinUI `NavigationView` in its default (Auto) display mode: expanded when the window is wide, a compact icon rail when it is narrower. The property tables use WinUI's `TableView` (`Microsoft.UI.Xaml.Controls.Tabular`), which is currently **experimental**: the app references Windows App SDK `2.5.4-experimental` for it, and should move to the stable release that includes TableView once it ships. It replaced the Windows Community Toolkit DataGrid, which scrolled poorly.

All icons come from Microsoft's [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) (MIT), bundled as fonts in `Assets/Fonts`. The optional colourful result icons are Microsoft's Fluent UI file type icons, bundled unmodified in `Assets/FileTypes` under the [Microsoft Fabric Assets License](https://aka.ms/fluentui-assets-license) (see the LICENSE.txt in that folder). The only other exceptions are built-in control parts such as drop-down chevrons, checkbox ticks and the window's minimize/maximize/close buttons, which Windows draws itself.

What was left out of the original tool on purpose: the raw request/response viewer, GET/POST switching, most query parameters, query rules, suggestions, rank-detail/explain views, the freshness-boost generator, and the forms, anonymous and explicit Windows-credential sign-in options.

## Signing in

### SharePoint Online (default): no app registration needed

This is the same method as the original tool's **"Authenticate using specific user account" → "SharePoint Online"** option. It signs in with Microsoft's public *SharePoint Online Management Shell* client (`9bc3ab49-b65d-410a-85ad-de819febfddc`), so you don't need to register an app in your tenant.

1. The first time a profile needs a token, a **Sign in to SharePoint Online** window opens with the Microsoft sign-in page. This window is a WinUI WebView2 control, the same kind of embedded window the original tool used.
2. The window keeps its cookies in `%LOCALAPPDATA%\SharePointSearchExplorer\WebView2`, so your sign-in session carries over between runs. It also allows Windows single sign-on, so a Microsoft 365 work account already signed in to Windows can be used without typing a password.
3. The token is saved in an encrypted (DPAPI, per Windows user) MSAL cache at `%LOCALAPPDATA%\SharePointSearchExplorer\msal_token_cache.bin`. Later searches and app restarts refresh it silently, with no window.

Each **profile** stores a site URL, an optional tenant and the account last used. Switching profile switches the site, tenant and account. Use **Profile → Sign in with a different account…** to pick another account for the same site.

The tenant is worked out from the site URL (`contoso.sharepoint.com` → `contoso.onmicrosoft.com`). Enter a tenant ID or domain under **Advanced** only for vanity domains or multi-geo setups.

### Optional: your own app registration

If your organization provides its own Entra app registration, enter its client ID under **Advanced → App (client) ID**. The registration must be a public client with redirect URI `http://localhost` and the SharePoint **AllSites.Read** delegated permission. Those profiles sign in through your default web browser instead of the embedded window.

### SharePoint Server (on-premises)

Choose **Windows account** as the sign-in method. Requests then use the current Windows user's credentials (Negotiate/NTLM).

## Build and run

Requirements: Windows 10 1809 or later, the .NET 8 SDK, and optionally Visual Studio 2022 with the *WinUI application development* workload.

```powershell
dotnet build SharePointSearchExplorer\SharePointSearchExplorer.csproj -c Release -p:Platform=x64
.\SharePointSearchExplorer\bin\x64\Release\net8.0-windows10.0.19041.0\win-x64\SharePointSearchExplorer.exe
```

The app is **unpackaged** and **self-contained with respect to the Windows App SDK**, so the build output runs directly and no MSIX or Windows App Runtime install is needed. Use `-p:Platform=ARM64` or `x86` to target other architectures. You can also open `SharePointSearchExplorer.sln` in Visual Studio and press F5.

Settings, profiles and query history are stored in `%LOCALAPPDATA%\SharePointSearchExplorer\settings.json`. If a `%LOCALAPPDATA%\SharePointSearchQueryTool` or `%LOCALAPPDATA%\SimpleSearchQueryTool` folder from an earlier name of the app exists, it is moved there on first launch, so settings and sign-ins carry over.

The project, solution, code namespaces and built executable (`SharePointSearchExplorer.exe`) all use the name `SharePointSearchExplorer`.

## Saved query files

**Query History → Export** writes queries to an XML file that you can back up or share:

```xml
<?xml version="1.0" encoding="utf-8"?>
<SavedQueries version="1" exported="2026-09-30T12:00:00.0000000Z">
  <Query>
    <QueryText>ContentType:"Document" IsContainer="false"</QueryText>
  </Query>
</SavedQueries>
```

**Import…** also accepts the shorter `<Query>…</Query>` form. Imported queries go to the top of the history in file order. Queries already in the history (ignoring case) are skipped. History holds 25 queries, so the oldest entries are dropped to make room.

## Project layout

```
SharePointSearchExplorer/
├── App.xaml(.cs)              App startup, shared services
├── MainWindow.xaml(.cs)       Title bar + query box, navigation rail, profile menu
├── Views/
│   ├── SearchPage             Results list, filter/sort, result details
│   ├── SettingsPage           App settings (Windows 11 Settings layout)
│   ├── HttpDetailsPage        Server Response: last search's HTTP details, full API response, debug
│   ├── QueryHistoryPage       Recent queries, XML import/export
│   ├── SettingsCard           One settings row: icon, header, description, control
│   └── ProfileDialog          Add / edit a site profile
├── ViewModels/
│   └── SearchViewModel.cs     Search, paging, refiners, property loading
├── Services/
│   ├── AuthService.cs         MSAL sign-in, token cache, per-profile accounts
│   ├── WebViewSignIn.cs       Embedded WebView2 sign-in window (SharePoint Online method)
│   ├── SearchService.cs       SharePoint Search REST calls and response parsing
│   ├── SettingsStore.cs       JSON persistence
│   └── QueryFile.cs           Saved queries XML format
├── Helpers/
│   ├── AccentColors.cs        Windows colour palette, accent colour overrides
│   ├── FluentIcons.cs         Icon fonts + {h:FluentIcon} markup extensions
│   └── FluentIcons.g.cs       Glyph table generated from the Fluent icon font map
├── Assets/AppIcon.ico/.png    App icon
├── Assets/Fonts/              Fluent UI System Icons fonts + license
└── Models/                    Profiles, results, result icons, property rows, settings
```

## How the property table works

The Search REST API has no "all properties" switch, so the tool uses the same approach as the original:

1. **Managed properties:** query for the item alone (`WorkId:"<id>"` as a hidden constraint) with `refiners=ManagedProperties(filter=5000/0/*)`. The refiner entries are the names of the item's managed properties. The tool then re-queries the item with those names as `SelectProperties`, sent via `POST /_api/search/postquery` because the list is too long for a URL.
2. **Crawled properties:** the same item query with `refiners=CrawledProperties(filter=5000/0/*)`. Search returns crawled property names but not their values.

Both lists are loaded when a result is selected. If the table reports no managed property names, the item was indexed before the `ManagedProperties` refiner was available. Re-indexing the library or site fixes this.

## Credits

Inspired by the PnP [SharePoint Search Query Tool](https://github.com/pnp/PnP-Tools/tree/master/Solutions/SharePoint.Search.QueryTool) by Mikael Svenson and the PnP community. The app's icons are Microsoft's [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) (MIT License), and its file type icons are Microsoft's Fluent file type icons, used under the [Microsoft Fabric Assets License](https://aka.ms/fluentui-assets-license). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) for details. This project is not affiliated with or endorsed by Microsoft or PnP.
