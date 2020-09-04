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
  private aadUserGroupIdsPromises: Map<number, Promise<string[]>> = new Map();
  private aadGroupSpMembershipsPromises: Map<number, Promise<ISpGroupMembership[]>> = new Map();

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
    const userDirectMemberships = await this.getSpUserMemberships(siteUserIdentifier.id);
    const aadGroupMemberships = await this.getAadGroupSpMemberships(siteUserIdentifier);

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

  private async getSpUserMemberships(siteUserId: number): Promise<ISpGroupMembership[]> {
    const siteUserInfo = await this.getSpUserAndMemberGroupsPromise(siteUserId);
    // There MUST be a better way to do this rather than casting.
    // We know that the ISiteUserInfo was expanded to include Groups, but the ISiteUserInfo type
    // doesn't have the Groups property. There is probably something I should be doing with union
    // types here!
    const siteGroups = (siteUserInfo as any).Groups as ISiteGroupInfo[];
    if (siteGroups.length) {
      return siteGroups.map((siteGroup) => {
        return {
          spGroup: siteGroup.Title,
          spGroupId: siteGroup.Id,
          membershipViaPrincipalName: siteUserInfo.UserPrincipalName,
          membershipViaPrincipalType: siteUserInfo.PrincipalType,
          membershipViaPrincipalSpId: siteUserInfo.Id,
        };
      });
    } else {
      return [
        {
          spGroup: undefined,
          spGroupId: undefined,
          membershipViaPrincipalName: siteUserInfo.Title,
          membershipViaPrincipalType: siteUserInfo.PrincipalType,
          membershipViaPrincipalSpId: siteUserInfo.Id,
        },
      ];
    }
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

  private getAadUserGroupIds(siteUserIdentifier: ISiteUserIdentifiers): Promise<string[]> {
    if (this.aadUserGroupIdsPromises.has(siteUserIdentifier.id)) {
      return this.aadUserGroupIdsPromises.get(siteUserIdentifier.id);
    } else {
      let aadUserGroupIds: Promise<string[]>;
      if (siteUserIdentifier.id) {
        console.debug("Getting AAD group ids for user by email:", siteUserIdentifier.email);
        aadUserGroupIds = graph.users.getById(siteUserIdentifier.email).getMemberGroups();
      } else {
        console.debug("Getting AAD group ids for current user");
        aadUserGroupIds = graph.me.getMemberGroups();
      }
      this.aadUserGroupIdsPromises.set(siteUserIdentifier.id, aadUserGroupIds);
      return aadUserGroupIds;
    }
  }

  private getAadGroupSpMemberships(siteUserIdentifier: ISiteUserIdentifiers): Promise<ISpGroupMembership[]> {
    if (this.aadGroupSpMembershipsPromises.has(siteUserIdentifier.id)) {
      return this.aadGroupSpMembershipsPromises.get(siteUserIdentifier.id);
    } else {
      const aadGroupSpMemberships = this.populateAAdGroupsAsSpUsers(siteUserIdentifier);
      this.aadGroupSpMembershipsPromises.set(siteUserIdentifier.id, aadGroupSpMemberships);
      return aadGroupSpMemberships;
    }
  }

  private async populateAAdGroupsAsSpUsers(siteUserIdentifier: ISiteUserIdentifiers): Promise<ISpGroupMembership[]> {
    const aadGroupIds = await this.getAadUserGroupIds(siteUserIdentifier);
    console.debug("Retrieved AAD group ids for user", siteUserIdentifier, aadGroupIds);
    if (aadGroupIds.length === 0) {
      return Promise.resolve([]);
    }

    const filter = aadGroupIds.map((id) => `substringof('|${id}',LoginName)`).join(" or ");
    const groupSiteUserInfos = await sp.web.siteUsers.filter(filter).get();
    console.debug("Found SP site users corresponding to AAD groups", groupSiteUserInfos);

    const groupSiteUserMembershipsPromises = groupSiteUserInfos.map((groupSiteUserInfo) =>
      this.getSpUserMemberships(groupSiteUserInfo.Id)
    );

    const groupSiteUserMemberships = await Promise.all(groupSiteUserMembershipsPromises);

    return ([] as ISpGroupMembership[]).concat(...groupSiteUserMemberships);
  }
}

export default UserGroupLookup;
