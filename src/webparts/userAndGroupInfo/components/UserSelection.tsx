import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import TenantUserPicker from "./TenantUserPicker";
import SpUserGroupLookup from "../../../services/SpUserGroupLookup";

export interface IUserSelectionProps {
  context: WebPartContext;
  userGroupLookup: SpUserGroupLookup;
  defaultUserSelectionMode?: UserSelectionMode | undefined;
  onSelectedUserChanged: (siteUserId: number, email: string) => void;
}

export enum UserSelectionMode {
  TenantSearch = 0,
  SiteFilter = 1,
}

const UserSelection: React.FunctionComponent<IUserSelectionProps> = (props) => {
  const defaultUserSelectionMode: UserSelectionMode = props.defaultUserSelectionMode
    ? props.defaultUserSelectionMode
    : UserSelectionMode.TenantSearch;

  const [userSelectionMode, setUserSelectionMode] = React.useState(defaultUserSelectionMode);

  const pickedTenantUserHandler = (loginName: string) => {
    if (loginName) {
      const email = loginName.substring(loginName.lastIndexOf("|") + 1);
      props.userGroupLookup.getSpSiteUserByLoginName(loginName).then((userInfo) => {
        if (userInfo) {
          props.onSelectedUserChanged(userInfo.Id, email);
        } else {
          props.onSelectedUserChanged(undefined, email);
        }
      });
    } else {
      props.onSelectedUserChanged(undefined, undefined);
    }
  };

  return (
    <>
      <TenantUserPicker context={props.context} onPickedTenantUserChanged={pickedTenantUserHandler} />
    </>
  );
};

export default UserSelection;
