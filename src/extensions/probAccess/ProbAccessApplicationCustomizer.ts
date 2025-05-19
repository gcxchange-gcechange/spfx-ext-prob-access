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

pnpSetup({
  sp: {
    baseUrl: "https://devgcx.sharepoint.com" // Update this link in Prod
  }
});

const LOG_SOURCE: string = 'ProBAccessApplicationCustomizer';

// Helper to normalize and compare names loosely
function normalizeName(name: string): string {
  return name.trim().toLowerCase();
}

function nameLooseMatch(userName: string, memberName: string): boolean {
  // Accepts: "adi makkar" ~ "Adi Makkar (psp)" etc
  const nUser = normalizeName(userName);
  const nMember = normalizeName(memberName);
  return nMember.includes(nUser) || nUser.includes(nMember);
}

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

      // get current user info (including display name)
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email ? currentUser.Email.toLowerCase() : '';
      const currentUserName = currentUser.Title ? currentUser.Title.trim() : '';
      if (!currentUserName) {
        throw new Error('Unable to determine user name');
      }
      console.log('Current user email:', currentUserEmail);
      console.log('Current user name:', currentUserName);

      // --- DOM Membership Check using MutationObserver ---
      const checkMembership = (): string[] => {
        const nodes = document.querySelectorAll('.ms-Persona-primaryText');
        const memberNames = Array.from(nodes).map(n => (n.textContent || '').trim()).filter(Boolean);
        console.log('Group member names:', memberNames);
        return memberNames;
      };

      const userInGroupPromise = new Promise<boolean>((resolve) => {
        const observer = new MutationObserver(() => {
          const memberNames = checkMembership();
          if (memberNames.length > 0) {
            const userInGroup = memberNames.some(member => nameLooseMatch(currentUserName, member));
            observer.disconnect();
            resolve(userInGroup);
          }
        });
        observer.observe(document.body, { childList: true, subtree: true });

        // Also run immediately in case the list is already present
        const memberNames = checkMembership();
        if (memberNames.length > 0) {
          observer.disconnect();
          resolve(memberNames.some(member => nameLooseMatch(currentUserName, member)));
        }
      });

      const userInGroup = await userInGroupPromise;
      if (!userInGroup) {
        // User is NOT a member/owner, redirect
        window.location.href = "https://devgcx.sharepoint.com";
        return Promise.resolve();
      }
      // User is authorized, continue
      console.log('User has the necessary access, no redirection needed.');
      return Promise.resolve();

    } catch (error) {
      // handle unexpected errors with redirection
      Log.error(LOG_SOURCE, error.message || error);
      console.error('Error:', error);

      // fallback redirection to the home page
      window.location.href = "https://devgcx.sharepoint.com";
      return Promise.resolve();
    }
  }
}