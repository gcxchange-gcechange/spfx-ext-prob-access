import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

// will initialize the PNP JS library with required headers
sp.setup({
  sp: {
    headers: {
      Accept: "application/json;odata=verbose",
    },
  },
});

interface SiteGroup {
  Id: number;
}

// function to check community access based on privacy settings and user requirements 
async function checkCommunityAccess(): Promise<void> {
  try {
    console.log("Starting community access check.");

    const isProtectedB = await isCommunityProtectedB();
    console.log("isProtectedB:", isProtectedB);

    if (isProtectedB) {
      // is responsible for checking if the privacy setting is set to public
      const sitePrivacySetting = await getSitePrivacySetting();
      console.log("sitePrivacySetting:", sitePrivacySetting);

      if (sitePrivacySetting === "Public") {
        // is responsible for checking if the user is a member or owner of the site
        const currentUser = await sp.web.currentUser();
        console.log("currentUser:", currentUser);
        const isMemberOrOwner = await isUserMemberOrOwner(currentUser.Id);
        console.log("isMemberOrOwner:", isMemberOrOwner);

        if (!isMemberOrOwner) {
          // will redirect the user to the home page if they are not a member or owner
          console.log("Redirecting user.");
          window.location.href = sp.web.toUrl(); // redirect to the current site's homepage
        }
      }
    }
    console.log("Community access check completed.");
  } catch (error) {
    console.error("Error checking community access:", error);
  }
}

// function to get the site privacy setting
async function getSitePrivacySetting(): Promise<string> {
  try {
    console.log("Fetching site properties.");
    const siteProperties = await sp.web.allProperties.get();
    console.log("Site Properties:", siteProperties);
    if (siteProperties.Privacy) {
      return siteProperties.Privacy;
    } else {
      console.log("Privacy property not found");
      return "";
    }
  } catch (error) {
    console.error("Error getting Site Privacy Setting:", error);
    console.error("Full Error:", error);
    throw error;
  }
}

// function to check if a user is a member or owner of the site
async function isUserMemberOrOwner(userId: number): Promise<boolean> {
  try {
    console.log("Checking user membership.");
    const userGroups = await sp.web.currentUser.groups();
    console.log("User Groups:", userGroups);

    const siteOwnersGroup = await getSiteOwnersGroup();
    console.log("Site Owners Group:", siteOwnersGroup);

    const siteMembersGroup = await getSiteMembersGroup();
    console.log("Site Members Group:", siteMembersGroup);

    // it is responsible for checking if the user is in the owners or members group
    const isOwner = userGroups.some((group) => group.Id === siteOwnersGroup.Id);
    const isMember = userGroups.some((group) => group.Id === siteMembersGroup.Id);
    const result = isOwner || isMember;

    console.log(
      `User is owner: ${isOwner}, User is member: ${isMember}, User is member or owner: ${result}`
    );

    return result;
  } catch (error) {
    console.error("Error checking user:", error);
    return false;
  }
}

// function to get site owners group
async function getSiteOwnersGroup(): Promise<SiteGroup> {
  try {
    console.log("Getting Site Owners Group.");
    const ownersGroup = await sp.web.siteGroups.getByName("Site Owners").get();
    return ownersGroup;
  } catch (error) {
    console.error("Error getting Site Owners Group:", error);
    throw error; // re-throw the error if it comes up
  }
}

// function to get site members group
async function getSiteMembersGroup(): Promise<SiteGroup> {
  try {
    console.log("Getting Site Members Group.");
    const membersGroup = await sp.web.siteGroups.getByName("Site Members").get();
    return membersGroup;
  } catch (error) {
    console.error("Error getting Site Members Group:", error);
    throw error; // re-throw the error if it comes up
  }
}

// function to check if the community is Protected B
async function isCommunityProtectedB(): Promise<boolean> {
  try {
    console.log("Checking if community is Protected B.");
    const siteProperties = await sp.web.allProperties.get();
    console.log("Site Description:", siteProperties.Description);
    if (siteProperties.Description) {
      const isProtected = siteProperties.Description.includes(
        "PROTECTED B - PROTÉGÉ B"
      );
      console.log("Is Protected B:", isProtected);
      return isProtected;
    } else {
      console.log("Site description is empty");
      return false;
    }
  } catch (error) {
    console.error("Error checking if community is Protected B:", error);
    return false;
  }
}

// is responsible for calling the function to check community access
checkCommunityAccess()
  .then(() => console.log("Community access checked successfully"))
  .catch((error) => console.error("Error checking community access:", error));