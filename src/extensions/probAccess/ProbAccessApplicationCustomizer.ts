/**
 * ProBAccessApplicationCustomizer -
 * Checks if a site is Protected B by looking for /teams/b in the URL
 * If Access level is Public:
    Check if the user is a member or owner.
    If not, remove and redirect to the home page.
 * If Access level is Private:
    Do Nothing.
 * Additional Use Cases:
    Ensure the app catalog is never redirected.
    No redirection for new tabs or search bar accesses, except for unauthorized access to public Protected B sites.
 */

    import { override } from '@microsoft/decorators';
    import { Log } from '@microsoft/sp-core-library';
    import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
    import { sp } from "@pnp/sp";
    import "@pnp/sp/webs";
    import "@pnp/sp/site-groups/web";
    import "@pnp/sp/security";
    import "@pnp/sp/sites";
    import "@pnp/sp/site-users/web";
    import { setup as pnpSetup } from "@pnp/common";
    
    // Initialize PnPjs
    pnpSetup({
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
          const siteUrl = window.location.href.toLowerCase();
          console.log('Site URL:', siteUrl);
          const isProtectedB = siteUrl.includes("/teams/b");
          console.log('Is Protected B:', isProtectedB);
    
          // Check if the current URL is the app catalog page
          if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) { // need to update this link in Prod
            console.log('App catalog page detected, skipping redirection...');
            return Promise.resolve();
          }
    
          if (isProtectedB) {
            console.log('Checking site access level from the DOM...');
            const groupInfoElement = document.querySelector('[data-automationid="SiteHeaderGroupType"]');
    
            if (groupInfoElement) {
              const groupInfoText = groupInfoElement.textContent || '';
              console.log('Group Info Text:', groupInfoText);
              const isPublic = groupInfoText.includes('Public group');
              const isPrivate = groupInfoText.includes('Private group');
              console.log('Is Public:', isPublic);
              console.log('Is Private:', isPrivate);
    
              if (isPublic) {
                console.log('Fetching user groups...');
                const userGroups = await sp.web.currentUser.groups.get();
                console.log('User Groups:', userGroups);
    
                // Check if the user is a member or owner of any group by membership
                const isMemberOrOwner = userGroups.some((group: { Id: number; Title: string; }) => {
                  console.log('Group ID:', group.Id);
                  console.log('Group Title:', group.Title);
                  return group.Title.includes("Members") || group.Title.includes("Owners");
                });
                console.log('Is Member or Owner:', isMemberOrOwner);
    
                if (!isMemberOrOwner) {
                  console.log('User is not a member or owner, redirecting...');
                  window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
                  return Promise.resolve();
                } else {
                  console.log('User is a member or owner, no redirection needed.');
                }
              } else {
                console.log('Site is private, no redirection needed.');
              }
            } else {
              console.error('Group info element not found.');
            }
          }
        } catch (error) {
          Log.error(LOG_SOURCE, error);
          console.error('Error:', error);
          window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
          return Promise.resolve();
        }
    
        console.log('User has the necessary access, no redirection needed.');
        return Promise.resolve();
      }
    }