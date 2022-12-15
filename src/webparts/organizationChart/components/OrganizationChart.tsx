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

    //this.getProfilePhoto = this.getProfilePhoto.bind(this);
    //this.getImageUrl = this.getImageUrl.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.getData();
  }

  public getData = async (): Promise<void> => {
    const client: MSGraphClientV3 =
      await this.props.context.msGraphClientFactory.getClient("3");
    const meResponse = await client
      .api("/me")
      .select("userPrincipalName,displayName,jobTitle")
      .get();
    //const meImage = await this.getProfilePhoto(meResponse.userPrincipalName);

    this.setState({
      Me: {
        text: meResponse.displayName,
        secondaryText: meResponse.jobTitle,
      },
    });

    const managerResponse = await client
      .api("/me/manager")
      .select("userPrincipalName,displayName,jobTitle")
      .get();
    //const managerImage = await this.getProfilePhoto(managerResponse.userPrincipalName);

    this.setState({
      Manager: {
        text: managerResponse.displayName,
        secondaryText: managerResponse.jobTitle,
      },
    });

    const reportsResponse = await client
      .api("/me/directReports")
      .select("userPrincipalName,displayName,jobTitle")
      .get();

    let reportsArr: IPersonaSharedProps[] = [];
    reportsResponse.value.forEach((item: any) => {
      let response: IPersonaSharedProps = {
        text: item.displayName,
        secondaryText: item.jobTitle,
      };

      reportsArr.push(response);
    });

    this.setState({ Reports: reportsArr });
  };

  //   private getProfilePhoto = async(userPrincipalName: string): Promise<string> => {
  //     const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");
  //     const response = await client.api("/users/" + userPrincipalName + "/photo/$value").get();
  //     return URL.createObjectURL(response);
  //   }

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
