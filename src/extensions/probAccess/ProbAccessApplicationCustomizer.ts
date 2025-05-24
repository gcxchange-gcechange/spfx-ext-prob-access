/**
 * ProBAccessApplicationCustomizer -
 * Checks if a site is Protected B by looking for /teams/b in the URL
 * If Access level is Public:
 *    Check if the user is a member or owner via Graph.
 *    If not, remove and redirect to the home page.
 * If Access level is Private:
 *    Do Nothing.
 * Additional Use Cases:
 *    Ensure the app catalog is never redirected.
 *    No redirection for new tabs or search bar accesses, except for unauthorized access to public Protected B sites.
 */

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { setup as pnpSetup } from "@pnp/common";
import { MSGraphClientV3 } from '@microsoft/sp-http';

pnpSetup({
  sp: {
    baseUrl: "https://devgcx.sharepoint.com" // Update this link in Prod
  }
});

const LOG_SOURCE: string = 'ProBAccessApplicationCustomizer';

interface GraphUser {
  id: string;
  mail?: string;
  userPrincipalName?: string;
}

interface GraphGroup {
  id: string;
  displayName?: string;
  sharepointSiteUrl?: string;
}

export default class ProBAccessApplicationCustomizer extends BaseApplicationCustomizer<{}> {
  // Helper: Check if user is in array of Graph API users
  private isUserInGraphResults(results: GraphUser[], userEmail: string): boolean {
    if (!results || !results.length) return false;
    return results.some(user =>
      (user.mail && user.mail.toLowerCase() === userEmail) ||
      (user.userPrincipalName && user.userPrincipalName.toLowerCase() === userEmail)
    );
  }

  // Helper: Find group matching current site URL
  private async getGroupForCurrentSite(
    graphClient: MSGraphClientV3,
    siteUrl: string
  ): Promise<GraphGroup | null> {
    const url = encodeURIComponent(siteUrl);
    // Try to match group by sharepointSiteUrl
    const response: { value: GraphGroup[] } = await graphClient
      .api(`/groups?$filter=groupTypes/any(c:c eq 'Unified') and sharepointSiteUrl eq '${url}'`)
      .version('v1.0')
      .get();

    if (response.value && response.value.length > 0) {
      return response.value[0]; // Found the group
    }

    // Fallback: Try to match by site URL prefix (sometimes sharepointSiteUrl can have trailing slash or use lowercase)
    const allUnifiedGroups: { value: GraphGroup[] } = await graphClient
      .api(`/groups?$filter=groupTypes/any(c:c eq 'Unified')`)
      .version('v1.0')
      .select('id,displayName,sharepointSiteUrl')
      .top(999)
      .get();

    if (allUnifiedGroups.value && allUnifiedGroups.value.length > 0) {
      const found = allUnifiedGroups.value.find(g =>
        g.sharepointSiteUrl && siteUrl.startsWith(g.sharepointSiteUrl.toLowerCase())
      );
      return found || null;
    }

    return null;
  }

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ProBAccessApplicationCustomizer`);
    console.log('Initialized ProBAccessApplicationCustomizer');

    try {
      const siteUrl = window.location.href.toLowerCase();
      console.log('Site URL:', siteUrl);

      // check if the site is Protected B
      const isProtectedB = siteUrl.includes("/teams/b");
      console.log('Is Protected B:', isProtectedB);

      if (!isProtectedB) {
        console.log('Not a Protected B site, skipping checks...');
        return Promise.resolve();
      }

      // skip checks for the app catalog
      if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) {
        console.log('App catalog page detected, skipping redirection...');
        return Promise.resolve();
      }

      // check the site's privacy setting
      const siteProperties = await sp.site.get();
      const isPublic = siteProperties.Privacy !== "Private";
      console.log('Is Public:', isPublic);

      if (!isPublic) {
        console.log('Site is private, no redirection required.');
        return Promise.resolve();
      }

      // get current user info (including display name)
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email ? currentUser.Email.toLowerCase() : '';
      const currentUserName = currentUser.Title ? currentUser.Title.trim() : '';
      if (!currentUserName) {
        throw new Error('Unable to determine user name');
      }
      console.log('Current user email:', currentUserEmail);
      console.log('Current user name:', currentUserName);

      // === GRAPH API CHECK ===
      // Dynamically find the group for this site
      const graphClient = await this.context.msGraphClientFactory.getClient('3');
      const group = await this.getGroupForCurrentSite(graphClient, siteUrl);

      if (!group) {
        console.warn('No Microsoft 365 Group found for this site. Redirecting...');
        window.location.href = "https://devgcx.sharepoint.com";
        return Promise.resolve();
      }

      console.log(`Found group: ${group.displayName} (${group.id})`);

      // Owners
      const ownersResponse = await graphClient.api(`/groups/${group.id}/owners`).version('v1.0').get();
      const owners: GraphUser[] = ownersResponse.value || [];
      // Transitive Members
      const membersResponse = await graphClient.api(`/groups/${group.id}/transitiveMembers`).version('v1.0').get();
      const members: GraphUser[] = membersResponse.value || [];

      // Check if current user is owner or member
      const isOwner = this.isUserInGraphResults(owners, currentUserEmail);
      const isMember = this.isUserInGraphResults(members, currentUserEmail);

      console.log('Is user an Owner:', isOwner);
      console.log('Is user a Transitive Member:', isMember);

      if (isOwner || isMember) {
        // User is authorized
        console.log('User is authorized (Graph check passed).');
        return Promise.resolve();
      }

      // Not an owner/member: redirect
      console.warn('User is not owner/member of group, redirecting...');
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();

    } catch (error: any) {
      // handle unexpected errors with redirection
      Log.error(LOG_SOURCE, error.message || error);
      console.error('Error:', error);

      // fallback redirection to the home page
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();
    }
  }
}