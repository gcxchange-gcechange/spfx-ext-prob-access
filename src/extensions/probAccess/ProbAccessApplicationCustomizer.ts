import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/security";
import "@pnp/sp/sites";
import { IWebInfo } from '@pnp/sp/webs';
import "@pnp/sp/site-users/web";

// Initialize PnPjs
sp.setup({
  sp: {
    baseUrl: "https://devgcx.sharepoint.com"
  }
});

const LOG_SOURCE: string = 'ProbAccessApplicationCustomizer';

export interface IProbAccessApplicationCustomizerProperties {
  // Define properties if any
}

export default class ProbAccessApplicationCustomizer extends BaseApplicationCustomizer<IProbAccessApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ProbAccessApplicationCustomizer`);

    try {
      const currentWeb = await sp.web();
      const siteUrl = currentWeb.Url;
      const isProtectedB = siteUrl.includes("/teams/b");

      if (isProtectedB) {
        interface IWebInfoWithPrivacy extends IWebInfo {
          PrivacyComplianceLevel: string;
        }

        const privacySetting = await sp.web.select("Title", "PrivacyComplianceLevel").get() as IWebInfoWithPrivacy;
        const isPublic = privacySetting.PrivacyComplianceLevel === "Public";

        if (isPublic) {
          const userGroups = await sp.web.currentUser.groups.get();
          const isMemberOrOwner = userGroups.some((group: { Title: string; }) => group.Title.includes("Members") || group.Title.includes("Owners"));

          if (!isMemberOrOwner) {
            window.location.href = "https://devgcx.sharepoint.com";
          }
        } else {
          window.location.href = "https://devgcx.sharepoint.com";
        }
      } else {
        window.location.href = "https://devgcx.sharepoint.com";
      }
    } catch (error) {
      Log.error(LOG_SOURCE, error);
      window.location.href = "https://devgcx.sharepoint.com";
    }

    return Promise.resolve();
  }
}