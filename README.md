# spfx-teamslink

## Summary
- spfx-teamslink is a SPFX Application Customizer Extension that hides the join teams button for teams sites.

## Prerequisites
- spfx-teamslink is intended to be deployed tenant wide.

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
1.4|July 10, 2026|Disable teamslink

## Minimal Path to Awesome
- Clone this repository
- Ensure that you are at the solution folder
- Ensure the current version of the Node.js (16.3+)
  - **in the command-line run:**
    - **npm install**

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

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**