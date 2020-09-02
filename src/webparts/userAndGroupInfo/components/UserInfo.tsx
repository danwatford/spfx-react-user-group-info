import * as React from "react";
import { useState } from "react";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { Overlay, Spinner, SpinnerSize } from "office-ui-fabric-react";

export interface IUserInfoProps {
  siteUserInfoPromise: Promise<ISiteUserInfo>;
  currentAadUserPromise: Promise<any>;
}

const UserInfo: React.FunctionComponent<IUserInfoProps> = (props) => {
  const [loading, setLoading] = useState(true);
  const [spUserInfo, setSpUserInfo] = useState(undefined as ISiteUserInfo);
  const [aadUserInfo, setAadUserInfo] = useState({ id: "" } as { id: string });

  React.useEffect(() => {
    if (props.siteUserInfoPromise) {
      props.siteUserInfoPromise.then((_) => {
        setSpUserInfo(_);
        setLoading(false);
      });
    }
  }, [props.siteUserInfoPromise]);

  React.useEffect(() => {
    if (props.currentAadUserPromise) {
      props.currentAadUserPromise.then((_) => {
        setAadUserInfo(_);
      });
    }
  }, [props.currentAadUserPromise]);

  return loading ? (
    <Overlay isDarkThemed>
      <Spinner size={SpinnerSize.large} />
    </Overlay>
  ) : (
    <div>
      <h1>{spUserInfo.Title}</h1>
      <dl>
        <dt>Email</dt>
        <dd>{spUserInfo.Email}</dd>
        <dt>Login</dt>
        <dd>{spUserInfo.LoginName}</dd>
        <dt>SP User Id</dt>
        <dd>{spUserInfo.Id}</dd>
        <dt>AAD UserId</dt>
        <dd>{aadUserInfo.id}</dd>
      </dl>
    </div>
  );
};

export default UserInfo;
