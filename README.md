# SharePoint Search Explorer
Reimplementation of the PnP SharePoint Search Query tool in C# and WinUI 3. 

> [!IMPORTANT]
> This application was Vibe-Coded. If this is disagreeable to you, please proceed no further. 

## About
In my day-to-day work with SharePoint, I very frequently use the venerable [SharePoint Search Query Tool (SSQT)](https://github.com/pnp/PnP-Tools/tree/master/Solutions/SharePoint.Search.QueryTool). However, I started to have issues that caused friction when using the application. Namely, I had to sign-in each time I needed to run a query, and the properties windows would often quit unexpexctedly. I've been dabling with vibecoding certain utilities for my document management needs, so I figured I would try supervising a rebuild of a new version of the SSQT. 

### Project Goals
- Modernize the UI to use the WinUI 3 toolkit. This would be done to bring the application into cadence with Windows 11 look and feel. This also offers additional customization and theming which is always nice to have. 
- Focus the user experience around result properties. While I'm sure that there are a lot of users who focus on diving into payload code and writing their own custom advanced queries, this is not my focus for using the app. Therefore I have prioritized quickly searching for items and displaying their crawled and managed properties.
- Create user experience improvements such as easy filtering of search results, searching for properties, and build-in often used links like navigating to the result parent location, etc. 
- Keeping my sign-in persistent between sessions.
