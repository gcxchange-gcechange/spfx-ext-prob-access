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

/**
 * ProBAccessApplicationCustomizer -
 * Ensures that Protected B sites are only accessible to members or owners when public.
 */

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/security";
import { setup as pnpSetup } from "@pnp/common";
import { PermissionKind } from "@pnp/sp/security";

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

      // Check if the site is Protected B
      const isProtectedB = siteUrl.includes("/teams/b");
      console.log('Is Protected B:', isProtectedB);

      if (!isProtectedB) {
        console.log('Not a Protected B site, skipping checks...');
        return Promise.resolve();
      }

      // Skip checks for the app catalog
      if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) {
        console.log('App catalog page detected, skipping redirection...');
        return Promise.resolve();
      }

      // Check the site's privacy setting
      const siteProperties = await sp.site.get();
      const isPublic = siteProperties.Privacy !== "Private";
      console.log('Is Public:', isPublic);

      if (!isPublic) {
        console.log('Site is private, no redirection required.');
        return Promise.resolve();
      }

      // Check if the user has read permissions
      const hasAccess = await sp.web.currentUserHasPermissions(PermissionKind.ViewListItems);
      console.log('Does User Have Access:', hasAccess);

      if (!hasAccess) {
        // Retrieve the current user's email
        const currentUser = await sp.web.currentUser.get();
        const userEmail = currentUser.Email;
        console.log('Current User Email:', userEmail);

        // Check if the user is in any of the site's groups
        const groups = await sp.web.siteGroups();
        const userGroups = await Promise.all(
          groups.map(async group => {
            const users = await sp.web.siteGroups.getById(group.Id).users();
            return users.some(user => user.Email.toLowerCase() === userEmail.toLowerCase());
          })
        );

        const isMemberOrOwner = userGroups.includes(true);
        console.log('Is User a Member or Owner:', isMemberOrOwner);

        // Redirect if the user is not a member or owner
        if (!isMemberOrOwner) {
          console.log('User is not a member or owner, redirecting...');
          window.location.href = "https://devgcx.sharepoint.com";
          return Promise.resolve();
        }
      }

    } catch (error) {
      // Handle unexpected errors with redirection
      Log.error(LOG_SOURCE, error.message || error);
      console.error('Error:', error);

      // Fallback redirection to the home page
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();
    }

    console.log('User has the necessary access, no redirection needed.');
    return Promise.resolve();
  }
}