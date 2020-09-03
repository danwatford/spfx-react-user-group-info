import { sp, PrincipalType, ISiteGroupInfo, spODataEntity } from "@pnp/sp/presets/all";

import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { IUser } from "@pnp/graph/users";

/**
 * Contains various identifiers of a SharePoint site user that can then be used to query SharePoint and Graph APIs.
 */
export interface ISiteUserIdentifiers {
  id: number;
  email: string;
}

export interface ISpGroupMembership {
  spGroup: string | undefined;
  spGroupId: number | undefined;
  membershipViaPrincipalName: string;
  membershipViaPrincipalType: PrincipalType;
  membershipViaPrincipalSpId: number;
}

class UserGroupLookup {
  private spUserAndMemberGroupsPromises: Map<number, Promise<ISiteUserInfo>> = new Map();
  private aadUserPromises: Map<number, Promise<any>> = new Map();
  private currentAadUserGroupIds: Promise<string[]>;
  private aadGroupSpMemberships: Promise<ISpGroupMembership[]>;

  /**
   * Returns the specified user's membership of SP site groups where:
   * - The user is a directly assigned member of the SP site group.
   * - The user is a member of an AAD group which is itself a member of the SP site group.
   *
   * Included in the results are cases where the user is a member of an AAD group which
   * is known to the SP site, and the AAD group is therefore represented as an SP site user,
   * but where the AAD group is not a member of any SP site group.
   *
   * @param siteUserIdentifier The SharePoint site user Id and LoginName of the user to retrieve information for.
   * If the id property is 0 then the current user's information is retrieved.
   */
  public async getUserMemberships(siteUserIdentifier: ISiteUserIdentifiers): Promise<ISpGroupMembership[]> {
    const userDirectMemberships = await this.getCurrentSpUserMemberships(siteUserIdentifier.id);
    const aadGroupMemberships = await this.getAadGroupSpMemberships();

    return [...userDirectMemberships, ...aadGroupMemberships];
  }

  public getSpUserAndMemberGroupsPromise(siteUserId: number) {
    if (this.spUserAndMemberGroupsPromises.has(siteUserId)) {
      return this.spUserAndMemberGroupsPromises.get(siteUserId);
    } else {
      const spUserAndMemberGroupsPromise = siteUserId
        ? sp.web.getUserById(siteUserId).expand("Groups").get()
        : sp.web.currentUser.expand("Groups").get();
      this.spUserAndMemberGroupsPromises.set(siteUserId, spUserAndMemberGroupsPromise);
      return spUserAndMemberGroupsPromise;
    }
  }

  private async getCurrentSpUserMemberships(siteUserId: number): Promise<ISpGroupMembership[]> {
    const currentSiteUserInfo = await this.getSpUserAndMemberGroupsPromise(siteUserId);
    // There MUST be a better way to do this rather than casting.
    // We know that the ISiteUserInfo was expanded to include Groups, but the ISiteUserInfo type
    // doesn't have the Groups property. There is probably something I should be doing with union
    // types here!
    return ((currentSiteUserInfo as any).Groups as ISiteGroupInfo[]).map((siteGroup) => {
      return {
        spGroup: siteGroup.Title,
        spGroupId: siteGroup.Id,
        membershipViaPrincipalName: currentSiteUserInfo.UserPrincipalName,
        membershipViaPrincipalType: currentSiteUserInfo.PrincipalType,
        membershipViaPrincipalSpId: currentSiteUserInfo.Id,
      };
    });
  }

  public async getAadUser(siteUserIdentifier: ISiteUserIdentifiers): Promise<IUser> {
    if (this.aadUserPromises.has(siteUserIdentifier.id)) {
      return this.aadUserPromises.get(siteUserIdentifier.id);
    } else {
      let aadUserPromise: Promise<IUser>;
      if (siteUserIdentifier.id) {
        console.debug("Getting AAD user by email:", siteUserIdentifier.email);
        aadUserPromise = graph.users.getById(siteUserIdentifier.email).get();
      } else {
        console.debug("Getting AAD user for current user");
        aadUserPromise = graph.me() as Promise<IUser>;
      }
      this.aadUserPromises.set(siteUserIdentifier.id, aadUserPromise);
      return aadUserPromise;
    }
  }

  private getCurrentAadUserGroupIds() {
    if (!this.currentAadUserGroupIds) {
      this.currentAadUserGroupIds = graph.me.getMemberGroups();
    }
    return this.currentAadUserGroupIds;
  }

  private getAadGroupSpMemberships(): Promise<ISpGroupMembership[]> {
    if (!this.aadGroupSpMemberships) {
      this.aadGroupSpMemberships = this.populateAAdGroupsAsSpUsers();
    }
    return this.aadGroupSpMemberships;
  }

  private async populateAAdGroupsAsSpUsers(): Promise<ISpGroupMembership[]> {
    const aadGroupIds = await this.getCurrentAadUserGroupIds();
    const filter = aadGroupIds.map((id) => `substringof('|${id}',LoginName)`).join(" or ");
    const groupSiteUserInfos = await sp.web.siteUsers.filter(filter).get();

    return groupSiteUserInfos.map((groupSiteUserInfo) => {
      return {
        spGroup: undefined,
        spGroupId: undefined,
        membershipViaPrincipalName: groupSiteUserInfo.Title,
        membershipViaPrincipalType: groupSiteUserInfo.PrincipalType,
        membershipViaPrincipalSpId: groupSiteUserInfo.Id,
      };
    });
  }
}

export default UserGroupLookup;
