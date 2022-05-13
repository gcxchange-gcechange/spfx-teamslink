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

export interface ITeamsLinkApplicationCustomizerProperties {
  TeamsListUrl: string;
  hubSiteIds: string;
  noTeamsLink: string;
}

export default class TeamsLinkApplicationCustomizer
  extends BaseApplicationCustomizer<ITeamsLinkApplicationCustomizerProperties> {

  teamslinkId: string = "";

  @override
  public async onInit(): Promise<void> {

    graph.setup({
      spfxContext: this.context
    });
    // Check if a community site
    if(!this.context.pageContext.legacyPageContext.isHubSite && this.checkHubSiteIds()){

      var teamsUrl =  await this.getTeamURL();
      let isMember = await this.isMember();

      // Add conversations
      this.render(teamsUrl, isMember);

      var siteHeader = document.querySelector('[data-automationid="SiteHeader"]');

      var context = this;
      // Watch to see if elements change based on window size
      const observer = new MutationObserver(function(mutations_list) {
        mutations_list.forEach(function(mutation) {
          mutation.addedNodes.forEach(function(added_node) {

            // Desktop size
            if(added_node.isSameNode(siteHeader.querySelector('[class^="actionsWrapper-"]'))){
              if(!context.linkExists()) {
                let actionLink = context.createLink(teamsUrl);

                let spacer = document.createElement("span");
                spacer.className = styles.spacer;
                spacer.innerText = "|"

                if(isMember){
                  actionLink.innerText = strings.conversations;
                  actionLink.setAttribute("aria-label", strings.conversations);
                } else {
                  actionLink.innerText = strings.become;
                  actionLink.setAttribute("aria-label", strings.become);
                }

                siteHeader.querySelector('[class^="actionsWrapper-"]').prepend(spacer);
                siteHeader.querySelector('[class^="actionsWrapper-"]').prepend(actionLink);
              }
            // Mobile size
            } else if(added_node.isSameNode(siteHeader.querySelector('[class^="sideActionsWrapper-"]'))) {
              if(!context.linkExists()) {
                let actionLink = context.createLink(teamsUrl);

                if(isMember){
                  actionLink.innerText = strings.conversations;
                  actionLink.setAttribute("aria-label", strings.conversations);
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

    return Promise.resolve();
  }

  public render(teamsUrl, isMember) {
    if(this.linkExists())
      return;

    let actionLink = this.createLink(teamsUrl);

    let spacer = document.createElement("span");
    spacer.className = styles.spacer;
    spacer.innerText = "|"

    if(isMember){
      actionLink.innerText = strings.conversations;
      actionLink.setAttribute("aria-label", strings.conversations);
    } else {
      actionLink.innerText = strings.become;
      actionLink.setAttribute("aria-label", strings.become);
    }

    let actionsBar = document.querySelector('[class^="actionsWrapper-"]');
    if(actionsBar){
      actionsBar.prepend(spacer);
      actionsBar.prepend(actionLink);
    } else {
      this.applyMobileStyle();
      document.querySelector('[class^="sideActionsWrapper-"]').prepend(actionLink)
    }
  }

  public async getTeamURL() {
    var groupid = this.context.pageContext.site.group.id._guid;
    // Get teams link sharepoint list
    var url = await this.context.spHttpClient.get(this.properties.TeamsListUrl,
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
      var TeamsLinksList = jsonObject.value;
      //Get all item and check if matching groupid
      for (const item of TeamsLinksList) {
        var teamslink = ""
        if(item.TeamsID == groupid){
          teamslink = item.Teamslink
        } else{
          teamslink = this.properties.noTeamsLink
        }
        return teamslink
      }
    }).catch(function (error) {
      console.log(`Error:${error}`);
    });
    return Promise.resolve(url);
  }

  public async isMember(){
    var groupid = this.context.pageContext.site.group.id._guid;
    let isMember = false;

    await graph.me.checkMemberGroups([groupid]).then(res => {
      if(res[0] == groupid){
        isMember = true;
      }
    });

    return isMember;
  }

  private createLink(teamsUrl): HTMLAnchorElement {
    let actionLink = document.createElement("a");

    actionLink.href = teamsUrl;
    actionLink.className = styles.actionsLink;
    actionLink.target = "_blank";
    actionLink.id = this.teamslinkId;

    return actionLink;
  }

  private linkExists(): boolean {
    let link = document.getElementById(this.teamslinkId);
    return link !== null;
  }

  // Make sure the hub we're in is one of the approved hubs
  private checkHubSiteIds(): boolean {
    let context = this;
    let hubSiteIds = `${this.properties.hubSiteIds}`.replace(/\s/g, '').split(',');

    for(let i = 0; i < hubSiteIds.length; i++) {
      if(context.context.pageContext.legacyPageContext.hubSiteId == hubSiteIds[i])
        return true;
    }

    return false;
  }

  private applyMobileStyle(): void {
    var actionWrapper = document.querySelector('[data-automationid="SiteHeader"]').querySelector('[class^="sideActionsWrapper-"]') as HTMLElement;
    actionWrapper.style.display = "inline";

    var moreActions = actionWrapper.querySelector('[class^="moreActionsButton-"]') as HTMLElement;
    moreActions.style.display = "inline";
  }
}
