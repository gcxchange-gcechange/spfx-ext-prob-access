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
import "@pnp/sp/site-users/web";
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

      // Step 1: Check if the site is Protected B
      const isProtectedB = siteUrl.includes("/teams/b");
      console.log('Is Protected B:', isProtectedB);

      if (!isProtectedB) {
        console.log('Not a Protected B site, skipping checks...');
        return Promise.resolve();
      }

      // Step 2: Skip checks for the app catalog
      if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) {
        console.log('App catalog page detected, skipping redirection...');
        return Promise.resolve();
      }

      // Step 3: Check the site's privacy setting
      const siteProperties = await sp.site.get();
      const isPublic = siteProperties.Privacy !== "Private";
      console.log('Site Privacy:', siteProperties.Privacy, 'Is Public:', isPublic);

      if (!isPublic) {
        console.log('Site is private, no redirection required.');
        return Promise.resolve();
      }

      // Step 4: Get the current user
      const currentUser = await sp.web.currentUser.get();
      console.log('Current User:', currentUser);

      // Step 5: Get the Owners and Members groups
      const ownersGroup = await sp.web.associatedOwnerGroup.get();
      const membersGroup = await sp.web.associatedMemberGroup.get();
      console.log('Owners Group:', ownersGroup);
      console.log('Members Group:', membersGroup);

      // Step 6: Fetch users in Owners and Members groups
      const [owners, members] = await Promise.all([
        sp.web.siteGroups.getById(ownersGroup.Id).users.get(),
        sp.web.siteGroups.getById(membersGroup.Id).users.get()
      ]);

      console.log('Owners:', owners.map(user => user.Email));
      console.log('Members:', members.map(user => user.Email));

      // Step 7: Check if the current user is a member or owner
      const isMemberOrOwner = [...owners, ...members].some(user => {
        return (
          user.Email?.toLowerCase() === currentUser.Email?.toLowerCase() ||
          user.Id === currentUser.Id
        );
      });

      console.log('Is Member or Owner:', isMemberOrOwner);

      // Step 8: Redirect if the user does not have access
      if (!isMemberOrOwner) {
        console.log('User is not a member or owner, redirecting...');
        window.location.href = "https://devgcx.sharepoint.com";
        return Promise.resolve();
      }

    } catch (error) {
      // Handle unexpected errors with redirection
      Log.error(LOG_SOURCE, error.message || error);
      console.error('Error:', error);
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();
    }

    console.log('User has the necessary access, no redirection needed.');
    return Promise.resolve();
  }
}