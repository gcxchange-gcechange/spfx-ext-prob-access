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
    import { HttpClient } from '@microsoft/sp-http';
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
    
          // Check if a ProtectedB site (starts with lowercase 'b') and includes '/teams/b
          const isProtectedB = siteUrl.includes("/teams/b");
          console.log('Is Protected B:', isProtectedB);
    
          // Check if the current URL is the app catalog page
          if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) { // need to update this link in Prod
            console.log('App catalog page detected, skipping redirection...');
            return Promise.resolve();
          }
    
          if (isProtectedB) {
            const mailNickname = siteUrl.split('/teams/')[1];
            console.log('Mail Nickname:', mailNickname);
    
            // Get the groupId of the site using mailNickname
            const groupIdResponse = await this.context.httpClient.get(
              `https://graph.microsoft.com/v1.0/groups?$filter=mailNickname eq '${mailNickname}'`,
              HttpClient.configurations.v1
            );
            const groupIdJson = await groupIdResponse.json();
            const groupId = groupIdJson.value[0].id;
            console.log('Group ID:', groupId);
    
            // Check the visibility property of the group
            const groupVisibilityResponse = await this.context.httpClient.get(
              `https://graph.microsoft.com/v1.0/groups/${groupId}`,
              HttpClient.configurations.v1
            );
            const groupVisibilityJson = await groupVisibilityResponse.json();
            const isPublic = groupVisibilityJson.visibility === 'Public';
            console.log('Is Public:', isPublic);
    
            if (isPublic) {
              // Get the members of the group using groupId
              const membersResponse = await this.context.httpClient.get(
                `https://graph.microsoft.com/v1.0/groups/${groupId}/members`,
                HttpClient.configurations.v1
              );
              const membersJson = await membersResponse.json();
              const members = membersJson.value.map((member: any) => member.userPrincipalName);
              console.log('Group Members:', members);
    
              // Get the current user's userPrincipalName
              const currentUserResponse = await this.context.httpClient.get(
                `https://graph.microsoft.com/v1.0/me`,
                HttpClient.configurations.v1
              );
              const currentUserJson = await currentUserResponse.json();
              const currentUser = currentUserJson.userPrincipalName;
              console.log('Current User:', currentUser);
    
              // Check if the current user is a member of the group
              const isMember = members.includes(currentUser);
              console.log('Is Member:', isMember);
    
              if (!isMember) {
                console.log('User is not a member, redirecting...');
                window.location.href = "https://devgcx.sharepoint.com"; // need to update this link in Prod
                return Promise.resolve();
              } else {
                console.log('User is a member, no redirection needed.');
              }
            } else {
              console.log('Site is private, no redirection needed.');
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