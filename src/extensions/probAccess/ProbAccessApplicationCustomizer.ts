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
    baseUrl: "https://devgcx.sharepoint.com" // Update in Prod
  }
});

const LOG_SOURCE: string = 'ProBAccessApplicationCustomizer';

function normalizeName(name: string): string {
  return name.trim().toLowerCase();
}

function nameLooseMatch(userName: string, memberName: string): boolean {
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

      // dheck if Protected B and not app catalog
      const isProtectedB = siteUrl.includes("/teams/b");
      if (!isProtectedB || siteUrl.includes('/sites/appcatalog/_layouts/15/tenantAppCatalog.aspx/manageApps')) {
        return Promise.resolve();
      }

      // check privacy
      const siteProperties = await sp.site.get();
      const isPublic = siteProperties.Privacy !== "Private";
      if (!isPublic) return Promise.resolve();

      // get current user name
      const currentUser = await sp.web.currentUser.get();
      const userName = currentUser.Title ? currentUser.Title.trim() : '';
      if (!userName) throw new Error('Unable to determine user name');
      console.log('Current user name:', userName);

      // 4. DOM loaded (in case membership list is rendered after page load, use MutationObserver or retry a few times)
      const getMemberNames = (): string[] => {
        // select all elements containing member names in the group membership list
        const nodes = document.querySelectorAll('.ms-Persona-primaryText');
        return Array.from(nodes).map(n => (n.textContent || '').trim());
      };

      // wait for the names to appear (retry for up to 3 seconds)
      const waitForMembers = (): Promise<string[]> =>
        new Promise((resolve) => {
          let tries = 0;
          const maxTries = 30;
          const interval = setInterval(() => {
            const names = getMemberNames();
            if (names.length > 0 || tries++ > maxTries) {
              clearInterval(interval);
              resolve(names);
            }
          }, 100);
        });

      const memberNames = await waitForMembers();
      console.log('Group member names:', memberNames);

      // match user name against member list
      const userInGroup = memberNames.some(member => nameLooseMatch(userName, member));
      if (!userInGroup) {
        // user is NOT a member/owner, redirect
        window.location.href = "https://devgcx.sharepoint.com";
        return Promise.resolve();
      }
      // user is authorized, continue
      console.log('User has the necessary access, no redirection needed.');
    } catch (error) {
      Log.error(LOG_SOURCE, error.message || error);
      window.location.href = "https://devgcx.sharepoint.com";
    }
    return Promise.resolve();
  }
}