import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

import * as strings from 'ProbAccessApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ProtectedBCommunityExtensionApplicationCustomizer';

export interface IProtectedBCommunityExtensionApplicationCustomizerProperties {
  testMessage: string;
}

export default class ProtectedBCommunityExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IProtectedBCommunityExtensionApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      try {
        await this.checkProtectedBCommunityAccess();
      } catch (error) {
        console.error('Error during Protected B community access check:', error);
      }
    } else {
      console.log("This extension only works in SharePoint environments.");
    }

    return Promise.resolve();
  }

  private async checkProtectedBCommunityAccess(): Promise<void> {
    try {
      const isProtectedB: boolean = await this.isSiteProtectedB();

      if (isProtectedB) {
        const isPublic: boolean = await this.isSitePublic();

        if (isPublic) {
          const isMemberOrOwner: boolean = await this.isUserMemberOrOwner();

          if (!isMemberOrOwner) {
            this.redirectUserToHomePage();
          }
        }
      }
    } catch (error) {
      console.error('Error during Protected B and privacy check:', error);
    }
  }

  private async isSiteProtectedB(): Promise<boolean> {
    const spOpts: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    };

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/AllProperties?$select=ProtectedB`,
      SPHttpClient.configurations.v1,
      spOpts
    );

    if (response.ok) {
      const properties = await response.json();
      return properties.ProtectedB === "true"; 
    } else {
      console.error("error getting site properties");
      return false;
    }
  }

  private async isSitePublic(): Promise<boolean> {
    const spOpts: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    };
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/site/Group`,
        SPHttpClient.configurations.v1,
        spOpts
    );

    if (response.ok) {
        const groupInfo = await response.json();
        return groupInfo.IsPublic;
    } else {
        console.error("error getting group information");
        return false;
    }
  }

  private async isUserMemberOrOwner(): Promise<boolean> {
    const spOpts: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    };
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/CurrentUser/Groups`,
      SPHttpClient.configurations.v1,
      spOpts
    );

    if (response.ok) {
      const groups = await response.json();
      for (const group of groups.value) {
        if (group.WebUrl === this.context.pageContext.web.absoluteUrl) {
          return true;
        }
      }
      return false;
    } else {
      console.error("error getting user groups");
      return false;
    }
  }

  private redirectUserToHomePage(): void {
    window.location.href = this.context.pageContext.web.absoluteUrl;
  }
}