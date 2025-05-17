/**
 * ProBAccessApplicationCustomizer -
 * This script validates if a user is authorized to access a GCXchange community.
 * It checks membership in SharePoint Groups (e.g., Owners, Members).
 * If the user is not authorized, they are redirected to the home page.
 */

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users";
import { setup as pnpSetup } from "@pnp/common";

// Initialize PnPjs
pnpSetup({
  sp: {
    baseUrl: "https://devgcx.sharepoint.com", // Update this for production
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

      // Check if the site is a GCXchange community (e.g., contains /teams/b)
      const isGCXCommunity = siteUrl.includes("/teams/b");
      console.log('Is GCXchange Community:', isGCXCommunity);

      if (!isGCXCommunity) {
        console.log('Not a GCXchange community site, skipping checks...');
        return Promise.resolve();
      }

      // Skip checks for the app catalog
      if (siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) {
        console.log('App catalog page detected, skipping checks...');
        return Promise.resolve();
      }

      // Get the current user's email address
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email.toLowerCase();
      console.log('Current User Email:', currentUserEmail);

      // Extract the site address (community identifier) from the URL
      const siteAddress = this.getSiteAddressFromUrl(siteUrl);
      console.log('Site Address (Community Identifier):', siteAddress);

      if (!siteAddress) {
        console.error('Site address could not be determined. Redirecting user for safety.');
        window.location.href = "https://devgcx.sharepoint.com"; // Redirect to home page
        return Promise.resolve();
      }

      // Validate user membership in SharePoint Groups
      const isUserAuthorized = await this.isUserInAnyGroup(['Owners', 'Members'], currentUserEmail);

      if (!isUserAuthorized) {
        console.log('User is not authorized to access this community. Redirecting...');
        window.location.href = "https://devgcx.sharepoint.com"; // Redirect to home page
        return Promise.resolve();
      }

      console.log('User is authorized to access the community.');

    } catch (error) {
      // Handle errors gracefully
      Log.error(LOG_SOURCE, error.message || error);
      console.error('Error:', error);

      // Redirect to the home page as a fallback
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();
    }

    console.log('Access validation completed.');
    return Promise.resolve();
  }

  /**
   * Extracts the site address (community identifier) from the URL.
   * For example, in "https://devgcx.sharepoint.com/teams/b10001638", it extracts "b10001638".
   */
  private getSiteAddressFromUrl(url: string): string {
    const match = url.match(/\/(sites|teams)\/([^/?]+)/);
    if (match && match[2]) {
      const siteAddress = match[2].trim(); // Extract and trim the site address
      console.log('Extracted Site Address:', siteAddress);
      return siteAddress;
    }

    console.warn('Could not extract site address from URL:', url);
    return ""; // Return empty string if extraction fails
  }

  /**
   * Checks if the user is in any of the specified SharePoint Groups.
   * @param groupNames - An array of SharePoint group names to check.
   * @param userEmail - The email address of the current user.
   * @returns True if the user is in any of the specified groups, otherwise false.
   */
  private async isUserInAnyGroup(groupNames: string[], userEmail: string): Promise<boolean> {
    try {
      for (const groupName of groupNames) {
        const isInGroup = await this.isUserInGroup(groupName, userEmail);
        if (isInGroup) {
          console.log(`User is in group: ${groupName}`);
          return true; // User is authorized if found in any group
        }
      }
      return false; // User is not in any group
    } catch (error) {
      console.error('Error checking groups:', error);
      return false; // Default to unauthorized on error
    }
  }

  /**
   * Checks if the user is in a specific SharePoint Group.
   * @param groupName - The name of the SharePoint group to check.
   * @param userEmail - The email address of the current user.
   * @returns True if the user is in the group, otherwise false.
   */
  
  private async isUserInGroup(groupName: string, userEmail: string): Promise<boolean> {
    try {
      const groupUsers = await sp.web.siteGroups.getByName(groupName).users();
      const groupEmails = groupUsers.map(user => user.Email.toLowerCase());
      console.log(`Users in ${groupName} group:`, groupEmails);

      return groupEmails.includes(userEmail.toLowerCase());
    } catch (error) {
      console.error(`Error fetching group '${groupName}':`, error);
      return false; // Return false if the group is not found or an error occurs
    }
  }
}