import * as React from "react";
import styles from "./OrganizationChart.module.scss";
import { IOrganizationChartProps } from "./IOrganizationChartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
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
            console.log(response);
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
            console.log(response);
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
          .get((error, response: any, rawResponse?: any) => {
            console.log(response);
            // this.setState({
            //   Manager: {
            //     text: response.displayName,
            //     secondaryText: response.jobTitle,
            //   },
            // });
          });
      });
  }

  public render(): React.ReactElement<IOrganizationChartProps> {
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
      </>
    );
  }
}
