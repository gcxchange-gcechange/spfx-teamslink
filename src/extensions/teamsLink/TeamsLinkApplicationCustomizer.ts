/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

export interface ITeamsLinkApplicationCustomizerProperties {

}

export default class TeamsLinkApplicationCustomizer
  extends BaseApplicationCustomizer<ITeamsLinkApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {

    await super.onInit();

    this.context.application.navigatedEvent.add(this, this.removeTeamsLink);

    window.addEventListener('click', (event) => {
      const el = event.target as HTMLElement;

      if (el.innerHTML === "Republish" || el.className.includes("ms-Icon ms-Button-icon") || el.className.includes("ms-Icon--ChromeClose")) {
        const interval = window.setInterval(() => {
        const teamsChannelButton = document.querySelector('button[title="Go to the Microsoft Teams channel"]');

          if (teamsChannelButton !== null) {
            teamsChannelButton.remove();
            clearInterval(interval);
          }
        }, 1000);
      }
    });

    let resizeTimeout: number;

    window.addEventListener('resize', () => {
      clearTimeout(resizeTimeout);

      resizeTimeout = window.setTimeout(() => {
        const maxTime = 5000;
        const intervalTime = 1000;
        let timeElapsed = 0;

        const interval = window.setInterval(() => {
          const teamsChannelButton = document.querySelector(
            'button[title="Go to the Microsoft Teams channel"]'
          );

          if (teamsChannelButton) {
            teamsChannelButton.remove();
            clearInterval(interval);
            return;
          }

          timeElapsed += intervalTime;

          if (timeElapsed >= maxTime) {
            clearInterval(interval);
          }
        }, intervalTime);
      }, 200);
    });

    this.removeTeamsLink();

    return Promise.resolve();
  }

  public removeTeamsLink():void {
    const teamsChannelButton = document.querySelector('button[title="Go to the Microsoft Teams channel"]');
    if(teamsChannelButton) {
      teamsChannelButton.remove();
    }
  }
}