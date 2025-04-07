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
            console.log('Checking site access level...');
    
            // Extract mailNickname from the URL
            const mailNicknameMatch = siteUrl.match(/\/teams\/(b\d+)/);
            
            if (!mailNicknameMatch) {
              console.error('Mail nickname not found in URL.');
              return Promise.resolve();
            }
            const mailNickname = mailNicknameMatch[1];
            console.log('Mail Nickname:', mailNickname);
    
            // Get the group details using mailNickname
            const groupResponse = await sp.web.siteGroups.filter(`MailNickname eq '${mailNickname}'`).get();
            if (groupResponse.length === 0) {
              console.error('Group not found for mailNickname:', mailNickname);
              return Promise.resolve();
            }
            const groupId = groupResponse[0].Id;
            console.log('Group ID:', groupId);
    
            // Check the group's visibility property
            const group = await sp.web.siteGroups.getById(groupId).get();
            const isPublic = group.AllowMembersEditMembership && group.AllowRequestToJoinLeave && group.AutoAcceptRequestToJoinLeave;
            const isPrivate = !isPublic;
            console.log('Is Public:', isPublic);
            console.log('Is Private:', isPrivate);
    
            if (isPublic) {
              console.log('Checking user membership using the SharePoint REST API...');
              const membersResponse = await sp.web.siteGroups.getById(groupId).users.get();
              const currentUser = await sp.web.currentUser.get();
              const isMemberOrOwner = membersResponse.some((member) => {
                return member.Email === currentUser.Email || member.Id === currentUser.Id;
              });
              console.log('Is Member or Owner:', isMemberOrOwner);
    
              if (!isMemberOrOwner) {
                console.log('User is not a member or owner, redirecting...');
                // Wait for 10 minutes before redirecting
                setTimeout(() => {
                  window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
                }, 10 * 60 * 1000); // 10 minutes in milliseconds
                return Promise.resolve();
              } 
              
              else {
                console.log('User is a member or owner, no redirection needed.');
              }
            } 
            
            else {
              console.log('Site is private, no redirection needed.');
            }
          }
        } 
        
        catch (error) {
          Log.error(LOG_SOURCE, error);
          console.error('Error:', error);
          // Wait for 10 minutes before redirecting
          setTimeout(() => {
            window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
          }, 10 * 60 * 1000); // 10 minutes in milliseconds
          return Promise.resolve();
        }
    
        console.log('User has the necessary access, no redirection needed.');
        return Promise.resolve();
      }
    }