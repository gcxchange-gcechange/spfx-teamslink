import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'TeamsLinkApplicationCustomizerStrings';

const LOG_SOURCE: string = 'TeamsLinkApplicationCustomizer';

import { graph } from "@pnp/graph";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import "@pnp/graph/groups";

import styles from './components/TeamsLink.module.scss';

export interface ITeamsLinkApplicationCustomizerProperties {}

export default class TeamsLinkApplicationCustomizer
  extends BaseApplicationCustomizer<ITeamsLinkApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {

    graph.setup({
      spfxContext: this.context
    });

    // Check if a community site
    if(!this.context.pageContext.legacyPageContext.isHubSite && (this.context.pageContext.legacyPageContext.hubSiteId == "688cb2b9-e071-4b25-ad9c-2b0dca2b06ba" || this.context.pageContext.legacyPageContext.hubSiteId == "ab266fda-839c-4854-a682-72b52934ce40")){
      let teamsUrl = await this.getTeamURL();
      let isMember = await this.isMember();

      // Add conversations
      this.render(teamsUrl, isMember);

      var siteHeader = document.querySelector('[data-automationid="SiteHeader"]')

      // Watch to see if elements change based on window size
      const observer = new MutationObserver(function(mutations_list) {
        mutations_list.forEach(function(mutation) {
          mutation.addedNodes.forEach(function(added_node) {

            // Desktop size
            if(added_node.isSameNode(siteHeader.querySelector('[class^="actionsWrapper-"]'))){

              let actionLink = document.createElement("a");
              actionLink.href = teamsUrl;
              actionLink.className = styles.actionsLink;

              if(isMember){
                actionLink.innerText = "Conversations";
              } else {
                actionLink.innerText = strings.become;
              }

              siteHeader.querySelector('[class^="actionsWrapper-"]').append(actionLink);

            // Mobile size
            } else if(added_node.isSameNode(siteHeader.querySelector('[class^="sideActionsWrapper-"]'))) {

              let actionLink = document.createElement("a");
              actionLink.href = teamsUrl;
              actionLink.className = styles.actionsLink;

              if(isMember){
                actionLink.innerText = "Conversations";
              } else {
                actionLink.innerText = strings.become;
              }

              siteHeader.querySelector('[class^="sideActionsWrapper-"]').prepend(actionLink);

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

  public async render(teamsUrl, isMember) {

    let actionLink = document.createElement("a");
    actionLink.href = teamsUrl;
    actionLink.className = styles.actionsLink;

    if(isMember){
      actionLink.innerText = "Conversations";
    } else {
      actionLink.innerText = strings.become;
    }


    let actionsBar = document.querySelector('[class^="actionsWrapper-"]');
    if(actionsBar){
      actionsBar.append(actionLink);
    } else {
      document.querySelector('[class^="sideActionsWrapper-"]').append(actionLink)
    }

  }

  public async getTeamURL() {

    var groupid = this.context.pageContext.site.group.id._guid;
    var url = "";

    await this.callTeamsAPI(groupid).then(res => {
      res.forEach((channel) => {
        if(channel.displayName == "General"){
          url = channel.webUrl;
        }
      });

      // If no General channel found, take first channel
      if(url == ""){
        url = res[0].webUrl;
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

    let response = await res;

    return response;
  }
}
