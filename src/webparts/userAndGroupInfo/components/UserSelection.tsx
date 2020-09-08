import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { Toggle } from "office-ui-fabric-react/lib/Toggle";

import TenantUserPicker from "./TenantUserPicker";
import SpUserGroupLookup from "../../../services/SpUserGroupLookup";
import SiteUserPicker from "./SiteUserPicker";

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
  const [pickedUserEmail, setPickedUserEmail] = React.useState(props.context.pageContext.user.email);

  const selectionModeChangedHandler = (event: React.MouseEvent<HTMLElement>, checked: boolean) => {
    setUserSelectionMode(checked ? UserSelectionMode.TenantSearch : UserSelectionMode.SiteFilter);
  };

  const pickedTenantUserHandler = (loginName: string) => {
    if (loginName) {
      const email = loginName.substring(loginName.lastIndexOf("|") + 1);
      props.userGroupLookup.getSpSiteUserByLoginName(loginName).then((userInfo) => {
        setPickedUserEmail(email);
        if (userInfo) {
          props.onSelectedUserChanged(userInfo.Id, email);
        } else {
          props.onSelectedUserChanged(undefined, email);
        }
      });
    } else {
      setPickedUserEmail(undefined);
      props.onSelectedUserChanged(undefined, undefined);
    }
  };

  const pickedSiteUserHandler = (siteUserId: number, email: string) => {
    setPickedUserEmail(email);
    props.onSelectedUserChanged(siteUserId, email);

    // Jump back to Tenant search to reduce the vertical size of the web part.
    setUserSelectionMode(UserSelectionMode.TenantSearch);
  };

  return (
    <>
      <Toggle
        label="Choose user from"
        offText="Site users"
        onText="Tenant users"
        checked={userSelectionMode === UserSelectionMode.TenantSearch ? true : false}
        onChange={selectionModeChangedHandler}
      />
      {userSelectionMode === UserSelectionMode.TenantSearch && (
        <TenantUserPicker
          context={props.context}
          pickedUserEmail={pickedUserEmail}
          onPickedTenantUserChanged={pickedTenantUserHandler}
        />
      )}
      {userSelectionMode === UserSelectionMode.SiteFilter && (
        <SiteUserPicker
          context={props.context}
          spUserGroupLookup={props.userGroupLookup}
          onSelectedUserChanged={pickedSiteUserHandler}
        />
      )}
    </>
  );
};

export default UserSelection;
