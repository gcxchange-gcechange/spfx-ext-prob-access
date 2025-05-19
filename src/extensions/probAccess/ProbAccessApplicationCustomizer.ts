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
import "@pnp/sp/site-users";
import { setup as pnpSetup } from "@pnp/common";

// Initialize PnPjs
pnpSetup({
  sp: {
    baseUrl: "https://devgcx.sharepoint.com" // Update this link in Prod
  }
});

const LOG_SOURCE: string = 'ProBAccessApplicationCustomizer';

export default class ProBAccessApplicationCustomizer extends BaseApplicationCustomizer<{}> {

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

      // get current user email
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email ? currentUser.Email.toLowerCase() : '';
      console.log('Current user email:', currentUserEmail);

      // site groups
      const webUrl = sp.web.toUrl();
      const groupsResponse = await fetch(`${webUrl}/_api/web/sitegroups`, {
        headers: { 'Accept': 'application/json;odata=verbose' }
      });
      const groupsData = await groupsResponse.json();

      // find "Owners" and "Members" groups for the site
      const ownersGroup = groupsData.d.results.find((g: any) => g.Title.toLowerCase().includes('owners'));
      const membersGroup = groupsData.d.results.find((g: any) => g.Title.toLowerCase().includes('members'));

      // helper to fetch users in a group
      const getGroupUsers = async (groupId: number): Promise<string[]> => {
        const resp = await fetch(`${webUrl}/_api/web/sitegroups(${groupId})/users`, {
          headers: { 'Accept': 'application/json;odata=verbose' }
        });
        const data = await resp.json();
        return data.d.results.map((u: any) => u.Email?.toLowerCase()).filter((e: any) => !!e);
      };

      // get all emails in owners and members
      let allowedEmails: string[] = [];
      if (ownersGroup) {
        allowedEmails = allowedEmails.concat(await getGroupUsers(ownersGroup.Id));
      }
      if (membersGroup) {
        allowedEmails = allowedEmails.concat(await getGroupUsers(membersGroup.Id));
      }
      allowedEmails = Array.from(new Set(allowedEmails)); // unique

      console.log('Allowed emails for this community:', allowedEmails);

      if (!allowedEmails.includes(currentUserEmail)) {
        console.log('User is not in allowed group, redirecting...');
        window.location.href = "https://devgcx.sharepoint.com";
        return Promise.resolve();
      }
      
    } catch (error) {
      // handle unexpected errors with redirection
      Log.error(LOG_SOURCE, error.message || error);
      console.error('Error:', error);

      // fallback redirection to the home page
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();
    }

    console.log('User has the necessary access, no redirection needed.');
    return Promise.resolve();
  }
}