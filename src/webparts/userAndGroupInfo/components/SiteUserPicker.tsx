import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { List } from "office-ui-fabric-react/lib/List";
import { Persona } from "office-ui-fabric-react/lib/Persona";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import SpUserGroupLookup from "../../../services/SpUserGroupLookup";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

import styles from "./UserAndGroupInfo.module.scss";

export interface ISiteUserPickerProps {
  context: WebPartContext;
  spUserGroupLookup: SpUserGroupLookup;
  onSelectedUserChanged: (siteUserId: number, email: string) => void;
}

const SiteUserPicker: React.FunctionComponent<ISiteUserPickerProps> = (props) => {
  const [filter, setFilter] = React.useState("");
  const [siteUserInfos, setSiteUserInfos] = React.useState([] as ISiteUserInfo[]);
  const [filteredSiteUserInfos, setFilteredSiteUserInfos] = React.useState([]);
  const [selectedUser, setSelectedUser] = React.useState(undefined);

  React.useEffect(() => {
    props.spUserGroupLookup.getSpSiteUsers().then((_) => setSiteUserInfos(_));
  }, [props.spUserGroupLookup]);

  React.useEffect(() => {
    if (filter) {
      setFilteredSiteUserInfos(siteUserInfos.filter((i) => i.Title.toLowerCase().includes(filter.toLowerCase())));
    } else {
      setFilteredSiteUserInfos(siteUserInfos);
    }
  }, [filter, siteUserInfos]);

  React.useEffect(() => {
    setFilteredSiteUserInfos([...filteredSiteUserInfos]);
  }, [selectedUser]);

  const onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    setFilter(text);
  };

  const personaClickedHandler = (siteUserInfo: ISiteUserInfo) => {
    console.log("Click for " + siteUserInfo.Id);
    setSelectedUser(siteUserInfo);
    props.onSelectedUserChanged(siteUserInfo.Id, siteUserInfo.Email);
  };

  const onRenderCell = (item: ISiteUserInfo, index: number | undefined) => {
    const classes = [styles.SiteUser];
    if (item === selectedUser) {
      classes.push(styles.Active);
    }

    console.log("Render item", item);
    if (selectedUser) {
      console.log("Selected item", selectedUser);
    }

    return (
      <div className={classes.join(" ")} onClick={() => personaClickedHandler(item)}>
        <Persona text={item.Title} />
      </div>
    );
  };

  return (
    <>
      <TextField label="Filter by Name:" onChange={onFilter} />
      <List className={styles.SiteUserList} items={filteredSiteUserInfos} onRenderCell={onRenderCell} />
    </>
  );
};

export default SiteUserPicker;
