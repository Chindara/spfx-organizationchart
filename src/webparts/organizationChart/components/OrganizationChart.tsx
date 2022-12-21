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

export default class OrganizationChart extends React.Component<
  IOrganizationChartProps,
  IOrganizationChartState
> {
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
    console.log(reportsResponse);

    if (reportsResponse) {
      this.setState({
        Reports: reportsResponse,
      });
    }


    //const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");
    // const meResponse = await client
    //   .api("/me")
    //   .select("id,userPrincipalName,displayName,jobTitle")
    //   .get();
    // //const meImage = await this.getProfilePhoto(meResponse.userPrincipalName);

    // 

    // this.setState({
    //   Me: {
    //     Id: meResponse.id,
    //     DisplayName: meResponse.displayName,
    //     JobTitle: meResponse.jobTitle,
    //     Email: meResponse.userPrincipalName,
    //     Presence: PersonaPresence.none,
    //   },
    // });

    // this.getPresence(meResponse.id);

    // const managerResponse = await client
    //   .api("/me/manager")
    //   .select("userPrincipalName,displayName,jobTitle")
    //   .get();
    // //const managerImage = await this.getProfilePhoto(managerResponse.userPrincipalName);

    // this.setState({
    //   Manager: {
    //     text: managerResponse.displayName,
    //     secondaryText: managerResponse.jobTitle,
    //   },
    // });

    // const reportsResponse = await client
    //   .api("/me/directReports")
    //   .select("userPrincipalName,displayName,jobTitle")
    //   .get();

    // let reportsArr: IPersonaSharedProps[] = [];
    // reportsResponse.value.forEach((item: any) => {
    //   let response: IPersonaSharedProps = {
    //     text: item.displayName,
    //     secondaryText: item.jobTitle,
    //   };

    //   reportsArr.push(response);
    // });

    // this.setState({ Reports: reportsArr });
  };

  //   private getProfilePhoto = async(userPrincipalName: string): Promise<string> => {
  //     const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");
  //     const response = await client.api("/users/" + userPrincipalName + "/photo/$value").get();
  //     return URL.createObjectURL(response);
  //   }

  // private getPresence = async (userId: string): Promise<void> => {
  //   console.log("getPresence");

  //   const client: MSGraphClientV3 =
  //     await this.props.context.msGraphClientFactory.getClient("3");

  //   const reportsResponse = await client
  //     .api("/users/" + { userId } + "/presence")
  //     .get();

  //   console.log(reportsResponse);
  // };

  public render(): React.ReactElement<IOrganizationChartProps> {
    const users = this.state.Reports;
    console.log(users);


    return (
      <>
        <p>Manager</p>
        <Persona
          {...this.state.Manager}
          size={PersonaSize.size48}     
        />
        <p>You</p>
        <Persona
          {...this.state.Me}
          size={PersonaSize.size48}
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
