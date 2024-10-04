/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'TeamsLinkApplicationCustomizerStrings';

import { graph } from "@pnp/graph";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import "@pnp/graph/groups";

import styles from './components/TeamsLink.module.scss';

import { SPHttpClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';


export interface ITeamsLinkApplicationCustomizerProperties {
  TeamsListUrl: string;
  hubSiteIds: string;
  noTeamsLink: string;
}



export default class TeamsLinkApplicationCustomizer
  extends BaseApplicationCustomizer<ITeamsLinkApplicationCustomizerProperties> {

  teamslinkId: string = "f3f79be5-ebc1-4ce5-8435-96db86a4eb20";


  @override
  public async onInit(): Promise<void> {

    await super.onInit();

    this.context.application.navigatedEvent.add(this, this.initialize);
    this.context.application.navigatedEvent.add(this, this.removeTeamsLink);


    window.addEventListener('click', (event) => {
      const el = event.target as HTMLElement;

      if (el.innerHTML === "Republish" || el.className.includes("ms-Icon ms-Button-icon") || el.className.includes("ms-Icon--ChromeClose")) {
        const interval = window.setInterval(() => {
        const teamsChannelButton = document.querySelector('button[title="Go to the Microsoft Teams channel"]');
        const customTeamsButton = document.querySelector(`button[title="${strings.conversations}"]`);

          if (teamsChannelButton !== null) {
            teamsChannelButton.remove();
            clearInterval(interval);
          }

          if (customTeamsButton === null) {
            console.log("Button NULL? - ",customTeamsButton)
            this.initialize();
          }

        }, 1000)
      }
    })

    this.removeTeamsLink();

    console.log("onInit", this.context);

    return Promise.resolve();
  }



  public removeTeamsLink():void {
    const teamsChannelButton = document.querySelector('button[title="Go to the Microsoft Teams channel"]');
    if(teamsChannelButton) {
      teamsChannelButton.remove();
    }
  }


  public async initialize():Promise<string|void> {
    graph.setup({
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      spfxContext: this.context as any
    });
    // Check if a community site

    if(!this.context.pageContext.legacyPageContext.isHubSite && this.checkHubSiteIds()){

      const teamsUrl =  await this.getTeamURL();
      const isMember = await this.isMember();

      // Add conversations
      this.render(teamsUrl, isMember);

      const siteHeader = document.querySelector('[data-automationid="SiteHeader"]');

      // eslint-disable-next-line @typescript-eslint/no-this-alias
      const context = this;
      // Watch to see if elements change based on window size
      const observer = new MutationObserver(function(mutations_list) {
        mutations_list.forEach(function(mutation) {
          mutation.addedNodes.forEach(function(added_node) {

            // Desktop size
            if(added_node.isSameNode(siteHeader.querySelector('[class^="actionsWrapper-"]'))){
              if(!context.linkExists()) {
                const actionLink = context.createLink(teamsUrl);
                actionLink.className = styles.actionLinkBox;

                if(isMember){
                  actionLink.innerText = strings.conversations;
                  actionLink.setAttribute("aria-label", strings.conversations);
                } else {
                  actionLink.innerText = strings.become;
                  actionLink.setAttribute("aria-label", strings.become);
                }

                siteHeader.querySelector('[class^="actionsWrapper-"]').prepend(actionLink);
              }
            // Mobile size
            } else if(added_node.isSameNode(siteHeader.querySelector('[class^="sideActionsWrapper-"]'))) {
              if(!context.linkExists()) {
                const actionLink = context.createLink(teamsUrl);

                if(isMember){
                  actionLink.innerText = strings.conversations;
                  actionLink.setAttribute("aria-label", strings.conversations);
                  actionLink.setAttribute("name", strings.conversations);
                } else {
                  actionLink.innerText = strings.become;
                  actionLink.setAttribute("aria-label", strings.become);

                }

                context.applyMobileStyle();
                siteHeader.querySelector('[class^="sideActionsWrapper-"]').prepend(actionLink);
              }
            }
          });
        });
      });

      const config = {
        childList: true
      };

      observer.observe(siteHeader, config);
    }


  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public render(teamsUrl:any, isMember:boolean):any {
    if(this.linkExists())
      return;

    const actionLink = this.createLink(teamsUrl);
    actionLink.className = styles.actionLinkBox;

    if(isMember){
      actionLink.innerText = strings.conversations;
      actionLink.setAttribute("aria-label", strings.conversations);
      actionLink.setAttribute("title", strings.conversations);
    } else {
      actionLink.innerText = strings.become;
      actionLink.setAttribute("aria-label", strings.become);
      actionLink.setAttribute("title", strings.become);
    }

    const actionsBar = document.querySelector('[class^="actionsWrapper-"]');
    if(actionsBar){
      //actionsBar.prepend(spacer);
      actionsBar.prepend(actionLink);
    } else {
      this.applyMobileStyle();
      document.querySelector('[class^="sideActionsWrapper-"]').prepend(actionLink)
    }
    //document.querySelector<HTMLElement>('[class^="ms-Button ms-Button--icon teamsChannelLink-"]').style.display = "none";
    console.log("render")
  }

  public async getTeamURL():Promise<string|void> {
    const TeamsListUrl = "https://devgcx.sharepoint.com/sites/app-reference/_api/lists/GetByTitle('TeamsLink')/items?$top=4000";
    const noTeamsLink = "NOTEAMSLINK";
    const groupid = this.context.pageContext.site.group.id._guid;
    // Get teams link sharepoint list
    const url = await this.context.spHttpClient.get(TeamsListUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        "Accept": "application/json;odata=nometadata",
        "odata-version": "3.0"
      }
    })
    .then(function (response) {
      if (!response.ok){
          throw Error(`Error, status code ${response.status}`);
      }else{
          return response.json();
      }
    }).then(function(jsonObject){
      const TeamsLinksList = jsonObject.value;
      //Get all item and check if matching groupid
      let teamslink = noTeamsLink
      for (const item of TeamsLinksList) {
        if(item.TeamsID === groupid){
          teamslink = item.Teamslink
          break;
        }
      }
      return teamslink
    }).catch(function (error) {
      console.log(`Error:${error}`);
    });
    return Promise.resolve(url);
  }

  public async isMember():Promise<boolean>{
    const groupid = this.context.pageContext.site.group.id._guid;
    let isMember = false;

    await graph.me.checkMemberGroups([groupid]).then(res => {
      if(res[0] === groupid){
        isMember = true;
      }
    });

    return isMember;
  }

  private createLink(teamsUrl): HTMLAnchorElement {
    const actionLink = document.createElement("a");

    actionLink.href = teamsUrl;
    actionLink.className = styles.actionsLink;
    actionLink.target = "_blank";
    actionLink.id = this.teamslinkId;

    return actionLink;
  }

  private linkExists(): boolean {
    const link = document.getElementById(this.teamslinkId);
    return link !== null;
  }

  // Make sure the hub we're in is one of the approved hubs
  private checkHubSiteIds(): boolean {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const context = this;
    const hubSiteIds = `${this.properties.hubSiteIds}`.replace(/\s/g, '').split(',');
    for(let i = 0; i < hubSiteIds.length; i++) {
      if(context.context.pageContext.legacyPageContext.hubSiteId === hubSiteIds[i])
        return true;
    }

    return false;
  }

  private applyMobileStyle(): void {
    const actionWrapper = document.querySelector('[data-automationid="SiteHeader"]').querySelector('[class^="sideActionsWrapper-"]') as HTMLElement;
    actionWrapper.style.display = "inline";

    const moreActions = actionWrapper.querySelector('[class^="moreActionsButton-"]') as HTMLElement;
    moreActions.style.display = "inline";
  }



}



