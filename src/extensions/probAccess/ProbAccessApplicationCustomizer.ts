import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

// will initialize the PNP JS library with the required headers
sp.setup({
  sp: {
    headers: {
      Accept: "application/json;odata=verbose"
    }
  }
});

// function to check community access based on privacy settings and user requirements trying to access the community
async function checkCommunityAccess(): Promise<void> {
  try {
    const isProtectedB = await isCommunityProtectedB(); // will implement the function to check if the community is Protected B

    if (isProtectedB) {
      // will check if the privacy setting is set to public
      const sitePrivacySetting = await getSitePrivacySetting();
      if (sitePrivacySetting === "Public") {
        // will check if the user is a member or owner of the site
        const currentUser = await sp.web.currentUser();
        const isMemberOrOwner = await isUserMemberOrOwner(currentUser.Id);

        if (!isMemberOrOwner) {
          // will redirect the user to the home page if they are not a member or owner
          window.location.href = "https://devgcx.sharepoint.com/";
        }
      }
    }
  } catch (error) {
    console.error("Error checking community access:", error);
  }
}

// function to get the site privacy setting
async function getSitePrivacySetting(): Promise<string> {
  const siteProperties = await sp.web.allProperties.get(); // siteInfo is the place to look for siteProperties 
  return siteProperties.Privacy;
}

// function to check if a user is a member or owner of the site
async function isUserMemberOrOwner(userId: number): Promise<boolean> {
  try {
    const userGroups = await sp.web.currentUser.groups();
    const siteOwnersGroup = await getSiteOwnersGroup();
    const siteMembersGroup = await getSiteMembersGroup();

    // will check if the user is in the owners or members group
    return userGroups.some(group => group.Id === siteOwnersGroup.Id) || userGroups.some(group => group.Id === siteMembersGroup.Id);
  } catch (error) {
    console.error("Error checking user:", error);
    return false;
  }
}

// function to get site owners group
async function getSiteOwnersGroup(): Promise<any> {
  const ownersGroup = await sp.web.siteGroups.getByName("Site Owners").get();
  return ownersGroup;
}

// function to get site members group
async function getSiteMembersGroup(): Promise<any> {
  const membersGroup = await sp.web.siteGroups.getByName("Site Members").get();
  return membersGroup;
}

// function to check if the community is Protected B
async function isCommunityProtectedB(): Promise<boolean> {
  return true; 
}

// checks the community access
void checkCommunityAccess(); 
