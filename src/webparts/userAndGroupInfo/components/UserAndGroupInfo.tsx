import * as React from "react";
import styles from "./UserAndGroupInfo.module.scss";
import { IUserAndGroupInfoProps } from "./IUserAndGroupInfoProps";
import userGroupLookup from "../../../services/UserGroupLookup";
import UserInfo from "./UserInfo";
import UserGroupMemberships from "./UserGroupMemberships";

export default class UserAndGroupInfo extends React.Component<IUserAndGroupInfoProps, {}> {
  public render(): React.ReactElement<IUserAndGroupInfoProps> {
    return (
      <div className={styles.userAndGroupInfo}>
        <div className={styles.container}>
          <div className={styles.row}>
            <UserInfo
              siteUserInfoPromise={userGroupLookup.getCurrentSpUserAndMemberGroupsPromise()}
              currentAadUserPromise={userGroupLookup.getCurrentAadUser()}
            />
            <UserGroupMemberships membershipsPromise={userGroupLookup.getCurrentUserMemberships()} />
          </div>
        </div>
      </div>
    );
  }
}
