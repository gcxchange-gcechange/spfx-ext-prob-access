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
import "@pnp/sp/lists";
import "@pnp/sp/items";
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

      // Get the current user's email address
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email.toLowerCase();
      console.log('Current User Email:', currentUserEmail);

      // Retrieve the site address from the URL
      const siteAddress = this.getSiteAddressFromUrl(siteUrl); // Helper function
      console.log('Site Address:', siteAddress);

      if (!siteAddress) {
        console.error('Site address could not be determined. Redirecting user for safety.');
        window.location.href = "https://devgcx.sharepoint.com"; // Redirect user
        return Promise.resolve();
      }

      const siteUsers = await sp.web.lists.getByTitle(siteAddress).items.select("Email").get(); // Fetch users for the site
      const siteUserEmails = siteUsers.map(user => user.Email.toLowerCase());
      console.log('Site User Emails:', siteUserEmails);

      // Check if the current user's email exists in the site users list
      const isUserInSite = siteUserEmails.includes(currentUserEmail);
      console.log('Is User in Site:', isUserInSite);

      // Redirect if the user is not in the site users list
      if (!isUserInSite) {
        console.log('User is not in the site user list, redirecting...');
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

  // Helper function: Extracts site address from the URL
  private getSiteAddressFromUrl(url: string): string {
    // Extract the part of the URL after '/sites/' or '/teams/'
    const match = url.match(/\/(sites|teams)\/([^/?]+)/);
    if (match && match[2]) {
        const siteAddress = match[2].trim(); // Extract and trim the site address (e.g., b10001638)
        console.log('Extracted Site Address:', siteAddress);
        return siteAddress;
    }

    // If no match, log a warning and return an empty string to force redirection
    console.warn('Could not extract site address from URL:', url);
    return ""; // Return empty string to force redirection
  }
}