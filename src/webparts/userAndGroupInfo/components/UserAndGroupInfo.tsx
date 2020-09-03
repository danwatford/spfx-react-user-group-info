import * as React from "react";
import styles from "./UserAndGroupInfo.module.scss";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IUserAndGroupInfoProps } from "./IUserAndGroupInfoProps";
import UserGroupLookup, { ISiteUserIdentifiers } from "../../../services/UserGroupLookup";
import UserInfo from "./UserInfo";
import UserGroupMemberships from "./UserGroupMemberships";

const UserAndGroupInfo: React.FunctionComponent<IUserAndGroupInfoProps> = (props) => {
  const [userGroupLookup, setUserGroupLookup] = React.useState(new UserGroupLookup());

  // We use a SiteUserId of 0 to refer to the current user.
  const [currentSiteUserIdentifiers, setCurrentSiteUserId] = React.useState({
    id: 0,
    email: "",
  } as ISiteUserIdentifiers);

  // Since we are defaulting to the current user we can consider that a user is selected.
  const [userIsSelected, setUserIsSelected] = React.useState(true);

  const selectedUserChangedHandler = (items: any[]) => {
    if (items.length) {
      setUserIsSelected(true);
      const loginName: string = items[0].loginName;
      const email = loginName.substring(loginName.lastIndexOf("|") + 1);
      setCurrentSiteUserId({ id: items[0].id, email: email });
    } else {
      setUserIsSelected(false);
    }
  };

  return (
    <div className={styles.userAndGroupInfo}>
      <div className={styles.container}>
        <PeoplePicker
          context={props.context}
          titleText="User"
          placeholder="Enter user name"
          principalTypes={[PrincipalType.User]}
          selectedItems={selectedUserChangedHandler}
          webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
          ensureUser={true}
          defaultSelectedUsers={[props.context.pageContext.user.email]}
        />
        <UserInfo
          siteUserInfoPromise={userGroupLookup.getSpUserAndMemberGroupsPromise(currentSiteUserIdentifiers.id)}
          currentAadUserPromise={userGroupLookup.getAadUser(currentSiteUserIdentifiers)}
        />
        <UserGroupMemberships membershipsPromise={userGroupLookup.getUserMemberships(currentSiteUserIdentifiers)} />
      </div>
    </div>
  );
};

export default UserAndGroupInfo;
