import { sp, PrincipalType, ISiteGroupInfo } from "@pnp/sp/presets/all";

import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

export interface ISpGroupMembership {
  spGroup: string | undefined;
  spGroupId: number | undefined;
  membershipViaPrincipalName: string;
  membershipViaPrincipalType: PrincipalType;
  membershipViaPrincipalSpId: number;
}

class UserGroupLookup {
  private currentSpUserAndMemberGroupsPromise: Promise<ISiteUserInfo>;
  private currentAadUserPromise: Promise<any>;
  private currentAadUserGroupIds: Promise<string[]>;
  private aadGroupSpMemberships: Promise<ISpGroupMembership[]>;

  /**
   * Returns the current user's membership of SP site groups where:
   * - The user is a directly assigned member of the SP site group.
   * - The user is a member of an AAD group which is itself a member of the SP site group.
   *
   * Included in the results are cases where the user is a member of an AAD group which
   * is known to the SP site, and the AAD group is therefore represented as an SP site user,
   * but where the AAD group is not a member of any SP site group.
   */
  public async getCurrentUserMemberships(): Promise<ISpGroupMembership[]> {
    const userDirectMemberships = await this.getCurrentSpUserMemberships();
    const aadGroupMemberships = await this.getAadGroupSpMemberships();

    return [...userDirectMemberships, ...aadGroupMemberships];
  }

  public getCurrentSpUserAndMemberGroupsPromise() {
    if (!this.currentSpUserAndMemberGroupsPromise) {
      this.currentSpUserAndMemberGroupsPromise = sp.web.currentUser.expand("Groups").get();
    }
    return this.currentSpUserAndMemberGroupsPromise;
  }

  public async getCurrentSpUserMemberships(): Promise<ISpGroupMembership[]> {
    const currentSiteUserInfo = await this.getCurrentSpUserAndMemberGroupsPromise();
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

  public getCurrentAadUser() {
    if (!this.currentAadUserPromise) {
      this.currentAadUserPromise = graph.me();
    }
    return this.currentAadUserPromise;
  }

  public getCurrentAadUserGroupIds() {
    if (!this.currentAadUserGroupIds) {
      this.currentAadUserGroupIds = graph.me.getMemberGroups();
    }
    return this.currentAadUserGroupIds;
  }

  public getAadGroupSpMemberships(): Promise<ISpGroupMembership[]> {
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

export default new UserGroupLookup();
