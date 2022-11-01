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

export default class OrganizationChart extends React.Component<IOrganizationChartProps,IOrganizationChartState> {
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

  public render(): React.ReactElement<IOrganizationChartProps> {
    const users = this.state.Reports;

    return (
      <>
        <h4>Manager</h4>
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
        {users != null ? (
          <div>
            {users.length > 0 ? (
              <div>
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
