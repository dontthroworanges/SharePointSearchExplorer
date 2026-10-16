# SharePoint Search Explorer
This application is a reimplementation of the [SharePoint Search Query Tool (SSQT)](https://github.com/pnp/PnP-Tools/tree/master/Solutions/SharePoint.Search.QueryTool) in C# using the WinUI 3. <BR>

> [!IMPORTANT]
> **Disclaimer:** This application was built with AI-generated code. If you prefer not to use AI-generated software, please do not proceed. 

## About
In my day-to-day work with SharePoint, I very frequently use the venerable [SharePoint Search Query Tool (SSQT)](https://github.com/pnp/PnP-Tools/tree/master/Solutions/SharePoint.Search.QueryTool). However, I recently started to have issues that caused friction when using the application. Namely, I had to sign-in each time I needed to run a query, and the properties windows would often quit unexpexctedly. I've been dabling with vibecoding certain utilities for my document management needs, so I figured I would try supervising a rebuild of a new version of the SSQT. SharePoint Search Explorer is the result of that effort. 

### Project Goals
- Focus the user experience around result properties. While I'm sure that there are a lot of users who focus on diving into payload code and writing their own custom advanced queries, this is not my focus for using the app. Therefore I have prioritized quickly searching for items and displaying their crawled and managed properties. As such, I have hidden the majority of the Power User/Advanced type features by default. 
- Modernize the UI to use the WinUI 3 framework. This would be done to bring the application into cadence with current Windows 11 look and feel. This also offers additional customization and themeing which is always nice to have. 
- Create user experience improvements such as easy filtering of search results, searching for properties, and build-in often used links like navigating to the result parent location, etc. 
- Allow users to create multiple tenant profiles. 
- Keeping sign-in persistent between sessions.
- Application should be portable without any installation required. 
<BR><BR>
![Screenshot](/docs/screenshots/profiles-dark.png)