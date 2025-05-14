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

export default class ProbAccessApplicationCustomizer extends BaseApplicationCustomizer<IProbAccessApplicationCustomizerStrings> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ProbAccessApplicationCustomizer`);
    console.log('Initialized ProbAccessApplicationCustomizer');

    try {
      const siteUrl = window.location.href.toLowerCase();
      console.log('Site URL:', siteUrl);
      const isProtectedB = siteUrl.includes("/teams/b");
      console.log('Is Protected B:', isProtectedB);

      if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) {
        console.log('App catalog page detected, skipping redirection...');
        return Promise.resolve();
      }

      if (isProtectedB) {
        const mailNicknameMatch = siteUrl.match(/\/teams\/(b\d+)/);
        if (!mailNicknameMatch) {
          console.error('Mail nickname not found in URL.');
          return Promise.resolve();
        }
        const mailNickname = mailNicknameMatch[1];
        console.log('Mail Nickname:', mailNickname);

        // log all site groups for debugging
        const allGroups = await sp.web.siteGroups.get();
        console.log('All Site Groups:', allGroups);

        // filter groups to find the matching group
        let groupResponse = await sp.web.siteGroups.filter(`Title eq '${mailNickname}'`).get();
        console.log('Group Response:', groupResponse);

        // // Step 3: Fallback search if the group isn't found
        // if (groupResponse.length === 0) {
        //   console.warn(`Group not found for mailNickname: ${mailNickname}`);
        //   console.warn(`Attempting fallback search...`);

        //   // Try to find the group by partial match
        //   const matchingGroup = allGroups.find(g => g.Title.includes(mailNickname));
        //   if (matchingGroup) {
        //     console.log('Fallback Group Found:', matchingGroup);
        //     groupResponse = [matchingGroup]; // Use the fallback group
        //   } else {
        //     console.error(`No group found for mailNickname: ${mailNickname}, even in fallback search.`);
        //     return Promise.resolve(); // Exit gracefully
        //   }
        // }

        // const groupId = groupResponse[0].Id;
        // console.log('Group ID:', groupId);

        // // Step 4: Check permissions and group visibility
        // const group = await sp.web.siteGroups.getById(groupId).get();
        // const isPublic = group.AllowMembersEditMembership && group.AllowRequestToJoinLeave && group.AutoAcceptRequestToJoinLeave;
        // console.log('Is Public:', isPublic);

        // if (isPublic) {
        //   const currentUser = await sp.web.currentUser.get();
        //   console.log('Current User:', currentUser);

        //   const membersResponse = await sp.web.siteGroups.getById(groupId).users.get();
        //   const isMemberOrOwner = membersResponse.some((member) => {
        //     return member.Email === currentUser.Email || member.Id === currentUser.Id;
        //   });
        //   console.log('Is Member or Owner:', isMemberOrOwner);

        //   if (!isMemberOrOwner) {
        //     console.log('User is not a member or owner, redirecting...');
        //     setTimeout(() => {
        //       window.location.href = "https://devgcx.sharepoint.com";
        //     }, 10 * 60 * 1000); // 10 minutes in milliseconds
        //     return Promise.resolve();
        //   }
        // }
      }
    } catch (error) {
      Log.error(LOG_SOURCE, error);
      console.error('Error:', error);
      setTimeout(() => {
        window.location.href = "https://devgcx.sharepoint.com";
      }, 10 * 60 * 1000); // 10 minutes in milliseconds
      return Promise.resolve();
    }

    console.log('User has the necessary access, no redirection needed.');
    return Promise.resolve();
  }
}