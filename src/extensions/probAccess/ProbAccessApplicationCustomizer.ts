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

pnpSetup({
    sp: {
        baseUrl: "https://devgcx.sharepoint.com"
    }
});

const LOG_SOURCE: string = 'ProbAccessApplicationCustomizer';

export interface IProbAccessApplicationCustomizerProperties { }

export default class ProbAccessApplicationCustomizer extends BaseApplicationCustomizer<IProbAccessApplicationCustomizerProperties> {

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
                const groupInfoElement = document.querySelector('[data-automationid="SiteHeaderGroupType"]');

                if (groupInfoElement) {
                    const groupInfoText = groupInfoElement.textContent || '';
                    console.log('Group Info Text:', groupInfoText);
                    const isPublic = groupInfoText.includes('Public group');
                    console.log('Is Public:', isPublic);

                    if (isPublic) {
                        const currentWeb = sp.web;
                        const currentUser = await currentWeb.currentUser();
                        const userGroups = await currentWeb.siteGroups.get();

                        const isMemberOrOwner = await Promise.all(userGroups.map(async group => {
                            const groupUsers = await currentWeb.siteGroups.getById(group.Id).users.get();
                            return groupUsers.some(user => user.Id === currentUser.Id);
                        })).then(results => results.some(isMember => isMember));

                        console.log('Is Member or Owner (PnPjs):', isMemberOrOwner);

                        if (!isMemberOrOwner) {
                            console.log('User is not a member or owner, redirecting...');
                            window.location.href = "https://devgcx.sharepoint.com";
                            return Promise.resolve();
                        } else {
                            console.log('User is a member or owner, no redirection needed.');
                        }
                    } else {
                        console.log('Site is private, no redirection needed.');
                    }
                } else {
                    console.error('Group info element not found.');
                }
            }
        } catch (error) {
            Log.error(LOG_SOURCE, error);
            console.error('Error:', error);
            window.location.href = "https://devgcx.sharepoint.com";
            return Promise.resolve();
        }

        console.log('User has the necessary access, no redirection needed.');
        return Promise.resolve();
    }
}