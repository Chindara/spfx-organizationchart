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
          .select("displayName,jobTitle")
          .get((error, response: any, rawResponse?: any) => {
            // console.log(response);
            this.setState({
              Me: {
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
          .api("/me/manager")
          .select("displayName,jobTitle")
          .get((error, response: any, rawResponse?: any) => {
            //console.log(response);
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
            //console.log(responses.value);

            let reportsArr: IPersonaSharedProps[] = [];
            responses.value.forEach((item: any) => {
              let response: IPersonaSharedProps = {
                text: item.displayName,
                secondaryText: item.jobTitle,
              };
              //console.log(response);
              reportsArr.push(response);
            });

            //console.log(reportsArr);
            this.setState({ Reports: reportsArr });
          });
      });

    // this.state.Reports.map((user) => console.log(user));
  }

  public render(): React.ReactElement<IOrganizationChartProps> {
    //console.log(this.state.Manager);
    //console.log(this.state.Me);
    //console.log(this.state.Reports);
    const users = this.state.Reports;
    console.log(users);

    return (
      <>
        <h4>Manager2</h4>
        <Persona
          {...this.state.Manager}
          size={PersonaSize.size48}
          presence={PersonaPresence.none}
        />
        <h4>You</h4>
        <Persona
          {...this.state.Me}
          size={PersonaSize.size48}
          presence={PersonaPresence.none}
        />
        <h4>Reports</h4>
        {/* {users[0].text} */}

        {/* {users.length > 0 ? (
          <div>
            {users.map((user, index) => (
              <div key={index}>{user.text}</div>
            ))}
          </div>
        ) : null} */}
      </>
    );
  }
}
