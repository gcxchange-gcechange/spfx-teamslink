# spfx-teamslink

## Summary

- spfx-teamslink is a SPFX Application Customizer Extension that add a Conversations/Become a member link to community pages on gcxchange based on user membership to group. It also hide the teams icon next to the title.

- You need to update the hubSiteIds property with comma seperated GUIDs for your valid hubs. You can do so in the serve.json file or from the tenant wide exstensions list when deployed.

- You need to update the teamslinkId in TeamsLinkApplicationCustomizer.ts to avoid duplication.

- You need to maitain a SharePoint list and store the value of teams id and teams link of each teams site where you want to apply this extension.


## Prerequisites
- spfx-teamslink is intended to be deployed tenant wide.
- A SharePoint list should created with the following fields and store the value of teams id and teams link of each teams site where you want to apply this extension. 

Column Name|Field Type
-------|----
`Teamslink`|Single line of text
`TeamsID`|Single line of text

- Created List's URL should given as value of TeamsListUrl in TeamsLinkApplicationCustomizer.ts

## API permission
- Microsoft Graph, User.ReadBasic.All
- Microsoft Graph, Team.ReadBasic.All
- Microsoft Graph, Channel.ReadBasic.All

## Version 
![SPFX](https://img.shields.io/badge/SPFX-1.17.4-green.svg)
![Node.js](https://img.shields.io/badge/Node.js-v16.3+-green.svg)

## Applies to
- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Version history
Version|Date|Comments
-------|----|--------
1.0|Dec 9, 2021|Initial release
1.1|March 25, 2022|Next release
1.3|October 25, 2023| Upgraded to SPFX 1.17.4

## Minimal Path to Awesome
- Clone this repository
- Ensure that you are at the solution folder
- Ensure the current version of the Node.js (16.3+)
  - **in the command-line run:**
    - **npm install**
- Create a SharePoint list with the following fields and store the value of teams id and teams link of each teams site where you want to apply this extension. 

Column Name|Field Type
-------|----
`Teamslink`|Single line of text
`TeamsID`|Single line of text

- Give created list's url as the value of the TeamsListUrl in TeamsLinkApplicationCustomizer.ts. 

- To debug
  - go to the `config\serve.json` file and update `pageUrl` to url of any teams site
  - **in the command-line run:**
    - **gulp clean**
    - **gulp serve**
- To deploy: 
  - **in the command-line run:**
    - **gulp clean**
    - **gulp bundle --ship**
    - **gulp package-solution --ship**

- Upload the extension from `\sharepoint\solution` to your tenant's app store
- To add or modify extension properties
  - **Go to Modern Appcatalog**
  - **Click ...More features in the left side**
  - **Open the tenant wide Extension**
  - **Edit the hubSiteIds under the title called TeamsLink**

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**