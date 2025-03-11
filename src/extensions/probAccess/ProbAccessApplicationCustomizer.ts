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

export interface IProbAccessApplicationCustomizerProperties {}

export default class ProbAccessApplicationCustomizer extends BaseApplicationCustomizer<IProbAccessApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ProbAccessApplicationCustomizer`);
    console.log('Initialized ProbAccessApplicationCustomizer');

    const checkAccess = async (): Promise<void> => {
      try {
        const currentUrl = window.location.href;
        console.log('Current URL:', currentUrl);

        // Check if the current URL is the app catalog page
        if (currentUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) { // need to update this link in Prod
          console.log('App catalog page detected, skipping redirection...');
          return;
        }

        // Check if the current URL is a Protected B site
        if (!currentUrl.includes('/teams/b')) {
          console.log('Not a Protected B site, skipping further checks...');
          return;
        }

        console.log('Fetching current web...');
        const currentWeb = await sp.web();
        const siteUrl = currentWeb.Url;
        console.log('Site URL:', siteUrl);

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
          const userGroups = await sp.web.currentUser.groups();
          console.log('User Groups:', userGroups);
          const isMemberOrOwner = userGroups.some((group: { Title: string; }) => group.Title.includes("Members") || group.Title.includes("Owners"));
          console.log('Is Member or Owner:', isMemberOrOwner);

          if (!isMemberOrOwner) {
            console.log('User is not a member or owner, redirecting...');
            sessionStorage.setItem('redirected', 'true');
            sessionStorage.setItem('removedFromCommunity', 'true');
            window.location.href = "https://devgcx.sharepoint.com"; // need to update this in Prod
            return;
          } 
          
          else {
            console.log('User is a member or owner, granting access.');
            sessionStorage.setItem('redirected', 'true');
          }
        } 
        
        else {
          console.log('Privacy setting is not public, redirecting...');
          sessionStorage.setItem('redirected', 'true');
          sessionStorage.setItem('removedFromCommunity', 'true');
          window.location.href = "https://devgcx.sharepoint.com"; // need to update this in Prod
          return;
        }
      } 
      
      catch (error) {
        Log.error(LOG_SOURCE, error);
        console.error('Error:', error);
        sessionStorage.setItem('redirected', 'true');
        sessionStorage.setItem('removedFromCommunity', 'true');
        window.location.href = "https://devgcx.sharepoint.com"; // need to update this in Prod
        return;
      }

      console.log('User has the necessary access, no redirection needed.');
      sessionStorage.setItem('redirected', 'true');
    };

    // Initial check when the application customizer is initialized
    await checkAccess();

    // Listen for URL changes
    window.addEventListener('popstate', checkAccess);
    window.addEventListener('pushstate', checkAccess);
    window.addEventListener('replacestate', checkAccess);

    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    history.pushState = function (...args) {
      const result = originalPushState.apply(this, args);
      window.dispatchEvent(new Event('pushstate'));
      window.dispatchEvent(new Event('popstate'));
      return result;
    };

    history.replaceState = function (...args) {
      const result = originalReplaceState.apply(this, args);
      window.dispatchEvent(new Event('replacestate'));
      window.dispatchEvent(new Event('popstate'));
      return result;
    };
  }
}