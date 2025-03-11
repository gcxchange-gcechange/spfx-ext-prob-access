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
    baseUrl: "https://devgcx.sharepoint.com"
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

    // Clear session storage item for testing
    sessionStorage.removeItem('redirected');
    sessionStorage.removeItem('removedFromCommunity');

    // Check if redirection has already occurred
    if (sessionStorage.getItem('redirected') === 'true') {
      console.log('Redirection has already occurred, skipping...');
      return Promise.resolve();
    }

    // Check if the user has been removed from the community
    if (sessionStorage.getItem('removedFromCommunity') === 'true') {
      console.log('User has been removed from the community, redirecting...');
      window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
      return Promise.resolve();
    }

    // Check if the current URL is the app catalog page
    if (window.location.href.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) { // need to update this link in Prod
      console.log('App catalog page detected, skipping redirection...');
      return Promise.resolve();
    }

    // Check if the current URL is a Protected B site
    const currentUrl = window.location.href;
    const isProtectedB = currentUrl.includes("/teams/b"); // pro b sites only

    if (!isProtectedB) {
      console.log('Current URL is not a Protected B site, skipping redirection...');
      return Promise.resolve();
    }

    try {
      console.log('Fetching current web...');
      const currentWeb = await sp.web();
      const siteUrl = currentWeb.Url;
      console.log('Site URL:', siteUrl);

      const isProtectedB = siteUrl.includes("/teams/b"); // pro b sites only
      console.log('Is Protected B:', isProtectedB);

      if (isProtectedB) {
        interface IWebInfoWithPrivacy extends IWebInfo {
          PrivacyComplianceLevel: string;
        }

        console.log('Fetching privacy settings...');
        const privacySetting = await sp.web.select("Title", "PrivacyComplianceLevel").get() as IWebInfoWithPrivacy;
        console.log('Privacy Setting:', privacySetting);
        const isPublic = privacySetting.PrivacyComplianceLevel === "Public";
        console.log('Is Public:', isPublic);

        if (isPublic) {
          console.log('Fetching user groups...');
          const userGroups = await sp.web.currentUser.groups.get();
          console.log('User Groups:', userGroups);
          const isMemberOrOwner = userGroups.some((group: { Title: string; }) => group.Title.includes("Members") || group.Title.includes("Owners"));
          console.log('Is Member or Owner:', isMemberOrOwner);

          if (!isMemberOrOwner) {
            console.log('User is not a member or owner, redirecting...');
            sessionStorage.setItem('removedFromCommunity', 'true');
            window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
            return Promise.resolve();
          } 
          
          else {
            console.log('User is a member or owner, no redirection needed.');
            sessionStorage.removeItem('removedFromCommunity');
          }
        } 
        
        else {
          console.log('Privacy setting is not public, redirecting...');
          sessionStorage.setItem('removedFromCommunity', 'true');
          window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
          return Promise.resolve();
        }
      } 
      
      else {
        console.log('Site is Protected B, no redirection needed.');
      }
    } 
    
    catch (error) {
      Log.error(LOG_SOURCE, error);
      console.error('Error:', error);
      sessionStorage.setItem('removedFromCommunity', 'true');
      window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
      return Promise.resolve();
    }

    console.log('User has the necessary access, no redirection needed.');
    sessionStorage.setItem('redirected', 'true');

    return Promise.resolve();
  }
}