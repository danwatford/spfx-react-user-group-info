import * as React from "react";
import { useState } from "react";
import { ISpGroupMembership } from "../../../services/SpUserGroupLookup";
import { DetailsList, IColumn } from "office-ui-fabric-react/lib/DetailsList";
import { PrincipalType } from "@pnp/sp";

export interface IUserGroupMembershipsProps {
  membershipsPromise: Promise<ISpGroupMembership[]>;
}

const columns: IColumn[] = [
  {
    key: "spGroup",
    name: "SP Group (SP Id)",
    minWidth: 200,
    onRender: (item) => {
      const membership = item as ISpGroupMembership;
      return membership.spGroupId ? (
        <span>
          {membership.spGroup} ({membership.spGroupId})
        </span>
      ) : (
        <span>none</span>
      );
    },
    isResizable: true,
  },
  {
    key: "membershipViaPrincipal",
    name: "Membership Via Principal",
    minWidth: 200,
    onRender: (item) => {
      const membership = item as ISpGroupMembership;
      return <span>{membership.membershipViaPrincipalName}</span>;
    },
    isResizable: true,
  },
  {
    key: "membershipViaPrincipalType",
    name: "Principal Type",
    minWidth: 100,
    onRender: (item) => {
      const membership = item as ISpGroupMembership;

      const principalTypeNames: string[] = [];
      const principalType = membership.membershipViaPrincipalType;
      if (principalType & PrincipalType.User) {
        principalTypeNames.push("User");
      }
      if (principalType & PrincipalType.DistributionList) {
        principalTypeNames.push("DistributionList");
      }
      if (principalType & PrincipalType.SecurityGroup) {
        principalTypeNames.push("SecurityGroup");
      }
      if (principalType & PrincipalType.SharePointGroup) {
        principalTypeNames.push("SharePointGroup");
      }

      return (
        <span>
          {principalTypeNames.join(" ")} ({principalType})
        </span>
      );
    },
    isResizable: true,
  },
  {
    key: "membershipViaPrincipalSpId",
    name: "Principal's SP Id",
    minWidth: 100,
    onRender: (item) => {
      const membership = item as ISpGroupMembership;
      return <span>{membership.membershipViaPrincipalSpId}</span>;
    },
    isResizable: true,
  },
];

const UserGroupMemberships: React.FunctionComponent<IUserGroupMembershipsProps> = (props) => {
  const [loading, setLoading] = useState(true);
  const [memberships, setMemberships] = useState(undefined as ISpGroupMembership[]);

  React.useEffect(() => {
    if (props.membershipsPromise) {
      props.membershipsPromise.then((spGroupMemberships) => {
        setMemberships(spGroupMemberships);
        setLoading(false);
      });
    }
  }, [props.membershipsPromise]);

  return loading ? (
    <h4>User's Group Information - Loading...</h4>
  ) : (
    <div>
      <h4>User's Group Information</h4>
      {memberships.length ? (
        <DetailsList items={memberships} columns={columns} />
      ) : (
        <p style={{ textAlign: "center" }}>User is not a member of any AAD groups known to this SharePoint site</p>
      )}
    </div>
  );
};

export default UserGroupMemberships;
