import * as React from "react";
import { IOrganizationChartProps } from "./IOrganizationChartProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IOrganizationChartState } from "./IOrganizationChartState";
import {
  FontSizes,
  Icon,
  IIconStyles,
  IPersonaProps,
  IPersonaSharedProps,
  IPersonaStyles,
  mergeStyleSets,
  Persona,
  PersonaPresence,
  PersonaSize,
} from "office-ui-fabric-react";
import UserService from "../../../services/UserService";

const presenceStatus: any[] = [];
presenceStatus["Available"] = PersonaPresence.online;
presenceStatus["AvailableIdle"] = PersonaPresence.online;
presenceStatus["Away"] = PersonaPresence.away;
presenceStatus["BeRightBack"] = PersonaPresence.away;
presenceStatus["Busy"] = PersonaPresence.busy;
presenceStatus["BusyIdle"] = PersonaPresence.busy;
presenceStatus["DoNotDisturb"] = PersonaPresence.dnd;
presenceStatus["Offline"] = PersonaPresence.offline;
presenceStatus["PresenceUnknown"] = PersonaPresence.none;

const personaStyles: Partial<IPersonaStyles> = {
  root: { margin: "0 0 10px 0" },
};
const iconStyles: Partial<IIconStyles> = { root: { marginRight: 5 } };

const classNames = mergeStyleSets({
  wrapper: {
    fontSize: '11px',
  }});

export default class OrganizationChart extends React.Component<IOrganizationChartProps, IOrganizationChartState> {
  private userService: UserService;
  constructor(props: IOrganizationChartProps) {
    super(props);
    this.state = {
      Me: null,
      Manager: null,
      Reports: null,
    };

    this.userService = new UserService();

    //this.getProfilePhoto = this.getProfilePhoto.bind(this);
    //this.getImageUrl = this.getImageUrl.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.getData();
  }

  public getData = async (): Promise<void> => {
    const meResponse: IPersonaSharedProps = await this.userService.getMe(this.props.context);
    if (meResponse) {
      this.setState({
        Me: meResponse,
      });
    }

    const managerResponse: IPersonaSharedProps = await this.userService.getManager(this.props.context);
    if (managerResponse) {
      this.setState({
        Manager: managerResponse,
      });
    }

    const reportsResponse: IPersonaSharedProps[] = await this.userService.getDirectReports(this.props.context);
    if (reportsResponse) {
      this.setState({
        Reports: reportsResponse,
      });
    }
  };

  public _onRenderSecondaryText(props: IPersonaProps): JSX.Element {
    return (
      <div>
        <Icon
          iconName="Suitcase"
          styles={iconStyles}
        />
        <i className={classNames.wrapper}>{props.secondaryText}</i>
      </div>
    );
  }

  public render(): React.ReactElement<IOrganizationChartProps> {
    const users = this.state.Reports;
    console.log(users);

    return (
      <>
        <p>Manager</p>
        <Persona
          {...this.state.Manager}
          size={PersonaSize.size48}
          onRenderSecondaryText={this._onRenderSecondaryText}
          styles={personaStyles}
        />
        <p>You</p>
        <Persona
          {...this.state.Me}
          size={PersonaSize.size48}
          onRenderSecondaryText={this._onRenderSecondaryText}
          styles={personaStyles}
        />
        <p>Reports</p>
        {users !== null ? (
          <div>{users.length}</div>
        ): null}



        {/* {users !== null ? (
          <div>
            {users.length > 0 ? (
              <div> */}
                {/* <p>Reports</p>
                {users.map((user, index) => (
                  <div key={index}>
                    <Persona
                      {...user}
                      size={PersonaSize.size48}
                      onRenderSecondaryText={this._onRenderSecondaryText}
                      styles={personaStyles}
                    />
                    <br />
                  </div>
                ))} */}
              {/* </div>
            ) : null}
          </div>
        ) : null} */}
      </>
    );
  }
}
