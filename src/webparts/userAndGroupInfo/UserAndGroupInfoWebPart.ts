import { override } from "@microsoft/decorators";
import { setup as pnpSetup } from "@pnp/common";
import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "UserAndGroupInfoWebPartStrings";
import UserAndGroupInfo from "./components/UserAndGroupInfo";
import { IUserAndGroupInfoProps } from "./components/IUserAndGroupInfoProps";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";

export interface IUserAndGroupInfoWebPartProps {
  description: string;
}

export default class UserAndGroupInfoWebPart extends BaseClientSideWebPart<IUserAndGroupInfoWebPartProps> {
  @override
  public async onInit(): Promise<void> {
    await super.onInit();
    console.log("Initialising SPFX context", this.context);
    pnpSetup({ spfxContext: this.context });

    const getCurrentUserPromise = sp.web.currentUser();
    const currentUser = await getCurrentUserPromise;

    console.log("Current user from onInit", currentUser);

    const aadCurrentUser = await graph.me();
    console.log("AAD user", aadCurrentUser);
  }

  public render(): void {
    const element: React.ReactElement<IUserAndGroupInfoProps> = React.createElement(UserAndGroupInfo, {});

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
