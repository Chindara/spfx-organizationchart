import * as React from "react";
import { IOrganizationChartProps } from "./IOrganizationChartProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IOrganizationChartState } from "./IOrganizationChartState";
import {
  IPersonaSharedProps,
  Persona,
  PersonaPresence,
  PersonaSize,
} from "office-ui-fabric-react";

export default class OrganizationChart extends React.Component<
  IOrganizationChartProps,
  IOrganizationChartState
> {
  constructor(props: IOrganizationChartProps) {
    super(props);
    this.state = {
      Me: null,
      Manager: null,
      Reports: null,
    };
  }

  public componentDidMount(): void {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me")
          //.select("userPrincipalName,displayName,jobTitle")
          .get((error, response: any, rawResponse?: any) => {
            console.log(response);
            this.setState({
              Me: {
                imageUrl: this.getProfilePhoto(response.userPrincipalName),
                text: response.displayName,
                secondaryText: response.jobTitle,
              },
            });
          });
      });

    console.log(this.state.Me);

    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me/manager")
          .select("displayName,jobTitle")
          .get((error, response: any, rawResponse?: any) => {
            this.setState({
              Manager: {
                text: response.displayName,
                secondaryText: response.jobTitle,
              },
            });
          });
      });

    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me/directReports")
          .select("displayName,jobTitle")
          .get((error, responses: any, rawResponse?: any) => {
            let reportsArr: IPersonaSharedProps[] = [];
            responses.value.forEach((item: any) => {
              let response: IPersonaSharedProps = {
                text: item.displayName,
                secondaryText: item.jobTitle,
              };
              reportsArr.push(response);
            });

            this.setState({ Reports: reportsArr });
          });
      });
  }

  private getProfilePhoto(userPrincipalName: string) {
    console.log(userPrincipalName);
    let blobUrl: string;
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/users/" + userPrincipalName + "/photo/$value")
          //.select("displayName,jobTitle")
          .get((error, response: any, rawResponse?: any) => {
            //console.log(response);

            const url = window.URL || window.webkitURL;
            blobUrl = url.createObjectURL(response.data);
            
          });
      });
      return blobUrl;
  }

  public render(): React.ReactElement<IOrganizationChartProps> {
    const users = this.state.Reports;

    return (
      <>
        <p>Manager</p>
        <Persona
          {...this.state.Manager}
          size={PersonaSize.size48}
          presence={PersonaPresence.none}
        />
        <p>You</p>
        <Persona
          {...this.state.Me}
          size={PersonaSize.size48}
          presence={PersonaPresence.none}
        />
        {users !== null ? (
          <div>
            {users.length > 0 ? (
              <div>
                <p>Reports</p>
                {users.map((user, index) => (
                  <div key={index}>
                    <Persona
                      {...user}
                      size={PersonaSize.size48}
                      presence={PersonaPresence.none}
                    />
                    <br />
                  </div>
                ))}
              </div>
            ) : null}
          </div>
        ) : null}
      </>
    );
  }
}
