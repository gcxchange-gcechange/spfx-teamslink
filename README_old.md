# spfx-teamslink

## Summary

Use graph API to add a Conversations/Become a member link to community pages on gcxchange based on user membership to group.

**_You need to update the hubSiteIds property with comma seperated GUIDs for your valid hubs. You can do so in the serve.json file or from the tenant wide exstensions list when deployed._**

**_You need to update the teamslinkId in TeamsLinkApplicationCustomizer.ts to avoid duplication_**

## Deployment

spfx-teamslink is intended to be deployed tenant wide.

## Required API access

These Graph permissions are required for spfx-teamslink to run properly
- User.ReadBasic.All
- Team.ReadBasic.All
- Channel.ReadBasic.All

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**
