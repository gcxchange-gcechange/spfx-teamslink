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

export interface ITeamsLinkApplicationCustomizerProperties {
  hubSiteIds: string
}

export default class TeamsLinkApplicationCustomizer
  extends BaseApplicationCustomizer<ITeamsLinkApplicationCustomizerProperties> {

  teamslinkId: string = "spfx-teamslink";

  @override
  public async onInit(): Promise<void> {
    console.log("SPFx - Teamslink");

    graph.setup({
      spfxContext: this.context
    });

    // Check if a community site
    if(!this.context.pageContext.legacyPageContext.isHubSite && this.checkHubSiteIds()){

      let teamsUrl = await this.getTeamURL();
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
                  actionLink.innerText = "Conversations";
                } else {
                  actionLink.innerText = strings.become;
                }

                siteHeader.querySelector('[class^="actionsWrapper-"]').prepend(spacer);
                siteHeader.querySelector('[class^="actionsWrapper-"]').prepend(actionLink);
              }
            // Mobile size
            } else if(added_node.isSameNode(siteHeader.querySelector('[class^="sideActionsWrapper-"]'))) {
              if(!context.linkExists()) {
                let actionLink = context.createLink(teamsUrl);

                if(isMember){
                  actionLink.innerText = "Conversations";
                } else {
                  actionLink.innerText = strings.become;
                }

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
      actionLink.innerText = "Conversations";
    } else {
      actionLink.innerText = strings.become;
    }

    let actionsBar = document.querySelector('[class^="actionsWrapper-"]');
    if(actionsBar){
      actionsBar.prepend(spacer);
      actionsBar.prepend(actionLink);
    } else {
      document.querySelector('[class^="sideActionsWrapper-"]').append(actionLink)
    }
  }

  public async getTeamURL() {

    var groupid = this.context.pageContext.site.group.id._guid;
    var url = "";

    await this.callTeamsAPI(groupid).then(res => {
      res.forEach((channel) => {
        if(channel.displayName == "General") {
          url = "https://teams.microsoft.com/_#/conversations/General?threadId=" + channel.id;
        }
      });

      // If no General channel found, take first channel
      if(url == ""){
        url = "https://teams.microsoft.com/_#/conversations/" + res[0].displayName + "?threadId=" + res[0].id;
      }
    });

    return url;
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

  private async callTeamsAPI(groupid){

    const res = await graph.teams.getById(groupid).channels();

    return res;
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

  // TODO: Rethink this
  // Original hubSiteIds: "688cb2b9-e071-4b25-ad9c-2b0dca2b06ba" "903ef314-6346-4d28-a135-07cd7a9f5c38"
  private checkHubSiteIds(): boolean {
    console.log("hubSiteId: " + this.context.pageContext.legacyPageContext.hubSiteId);
    return true;
    let context = this;
    let hubSiteIds = `${this.properties.hubSiteIds}`.replace(/\s/g, '').split(',');

    for(let i = 0; i < hubSiteIds.length; i++) {
      if(context.context.pageContext.legacyPageContext.hubSiteId == hubSiteIds[i])
        return true;
    }

    return false;
  }
}
