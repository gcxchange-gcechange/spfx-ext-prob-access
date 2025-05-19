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

interface IGraphUser {
  mail?: string;
}

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

      const groupIdResponse = await fetch(`${sp.web.toUrl()}/_api/site?$select=GroupId`, {
        headers: { 'Accept': 'application/json;odata=verbose' }
      });
      const groupIdData = await groupIdResponse.json();
      const groupId = groupIdData.d.GroupId;
      if (!groupId || groupId === "00000000-0000-0000-0000-000000000000") {
        console.log('No Microsoft 365 Group backing this site. Skipping further checks.');
        return Promise.resolve();
      }
      console.log("GroupId for this community:", groupId);

      const graphClient = await this.context.msGraphClientFactory.getClient("3");

      // Get members
      const membersResponse = await graphClient
        .api(`/groups/${groupId}/members`)
        .version("v1.0")
        .get();

      const memberEmails: string[] = (membersResponse.value || [])
        .map((user: IGraphUser) => user.mail?.toLowerCase())
        .filter((mail: string | undefined): mail is string => !!mail);

      // Get owners
      const ownersResponse = await graphClient
        .api(`/groups/${groupId}/owners`)
        .version("v1.0")
        .get();

      const ownerEmails: string[] = (ownersResponse.value || [])
        .map((user: IGraphUser) => user.mail?.toLowerCase())
        .filter((mail: string | undefined): mail is string => !!mail);

      // Combine and deduplicate
      const allowedEmails = Array.from(new Set([...memberEmails, ...ownerEmails]));

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