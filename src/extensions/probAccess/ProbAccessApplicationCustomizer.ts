import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/security";
import "@pnp/sp/sites";
import { IWebInfo } from '@pnp/sp/webs';
import "@pnp/sp/site-users/web";

// Initialize PnPjs
sp.setup({
  sp: {
    baseUrl: "https://devgcx.sharepoint.com" // need to update this link in Prod
  }
});

const LOG_SOURCE: string = 'ProbAccessApplicationCustomizer';

export interface IProbAccessApplicationCustomizerProperties {
}

export default class ProbAccessApplicationCustomizer extends BaseApplicationCustomizer<IProbAccessApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ProbAccessApplicationCustomizer`);
    console.log('Initialized ProbAccessApplicationCustomizer');

    try {
      // Check if the current URL is the app catalog page
      if (window.location.href.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) { // need to update this link in Prod
        console.log('App catalog page detected, skipping redirection...');
        return Promise.resolve();
      }

      console.log('Fetching current web...');
      const currentWeb = await sp.web();
      const siteUrl = currentWeb.Url;
      console.log('Site URL:', siteUrl);

      // Check if the site URL includes "/teams/b" for Protected B sites
      const isProtectedB = siteUrl.includes("/teams/b");
      console.log('Is Protected B:', isProtectedB);

      if (isProtectedB) {
        interface IWebInfoWithPrivacy extends IWebInfo {
          PrivacyComplianceLevel: string;
        }

        console.log('Fetching privacy settings...');
        const privacySetting = await sp.web.select("Title", "PrivacyComplianceLevel").get() as IWebInfoWithPrivacy;
        console.log('Privacy Setting:', privacySetting);

        // Check if the site is public
        const isPublic = privacySetting.PrivacyComplianceLevel === "Public";
        console.log('Is Public:', isPublic);

        if (isPublic) {
          console.log('Fetching user groups...');
          const userGroups = await sp.web.currentUser.groups.get();
          console.log('User Groups:', userGroups);

          // Check if the user is a member or owner
          const isMemberOrOwner = userGroups.some((group: { Title: string; }) => group.Title.includes("Members") || group.Title.includes("Owners"));
          console.log('Is Member or Owner:', isMemberOrOwner);

          if (!isMemberOrOwner) {
            console.log('User is not a member or owner, redirecting...');
            sessionStorage.setItem('redirected', 'true');
            sessionStorage.setItem('removedFromCommunity', 'true');
            window.location.href = "https://devgcx.sharepoint.com"; // need to update this in Prod
            return Promise.resolve();
          } else {
            console.log('User is a member or owner, no redirection needed.');
            sessionStorage.setItem('redirected', 'true');
          }
        } else {
          console.log('Privacy setting is not public, redirecting...');
          sessionStorage.setItem('redirected', 'true');
          sessionStorage.setItem('removedFromCommunity', 'true');
          window.location.href = "https://devgcx.sharepoint.com"; // need to update this in Prod
          return Promise.resolve();
        }
      }
    } catch (error) {
      Log.error(LOG_SOURCE, error);
      console.error('Error:', error);
      sessionStorage.setItem('redirected', 'true');
      sessionStorage.setItem('removedFromCommunity', 'true');
      window.location.href = "https://devgcx.sharepoint.com"; // need to update this in Prod
      return Promise.resolve();
    }

    console.log('User has the necessary access, no redirection needed.');
    sessionStorage.setItem('redirected', 'true');

    return Promise.resolve();
  }
}