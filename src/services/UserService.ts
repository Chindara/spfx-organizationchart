import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IPerson } from "../models/IPersonaListProps";
import { IPersonaSharedProps, PersonaPresence } from "office-ui-fabric-react";

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

export default class UserService {
  constructor() {}

  public async getMe(context: WebPartContext): Promise<IPersonaSharedProps> {
    const client: MSGraphClientV3 =
      await context.msGraphClientFactory.getClient("3");

    const apiResponse = await client
      .api("/me")
      .select("id,userPrincipalName,displayName,jobTitle")
      .get();

      console.log(apiResponse);
    // let person: IPerson = {
    //   Id: meResponse.id,
    //   DisplayName: meResponse.displayName,
    //   JobTitle: meResponse.jobTitle,
    // };

    const { availability, activity } = await this.getPresence(context,String(apiResponse.id));

    let response: IPersonaSharedProps = {
      text: apiResponse.displayName,
      secondaryText: apiResponse.jobTitle,
      presence: presenceStatus[availability],
      presenceTitle: activity,
    };

    console.log(response);
    return response;
  }

  public async getManager(
    context: WebPartContext
  ): Promise<IPersonaSharedProps> {
    const client: MSGraphClientV3 =
      await context.msGraphClientFactory.getClient("3");

    const apiResponse = await client
      .api("/me/manager")
      .select("id,userPrincipalName,displayName,jobTitle")
      .get();

    // let person: IPerson = {
    //   Id: meResponse.id,
    //   DisplayName: meResponse.displayName,
    //   JobTitle: meResponse.jobTitle,
    // };

    const { availability, activity } = await this.getPresence(context,String(apiResponse.id));

    let response: IPersonaSharedProps = {
      text: apiResponse.displayName,
      secondaryText: apiResponse.jobTitle,
      presence: presenceStatus[availability],
      presenceTitle: activity,
    };

    return response;
  }

  public async getDirectReports(
    context: WebPartContext
  ): Promise<IPersonaSharedProps[]> {
    const client: MSGraphClientV3 =
      await context.msGraphClientFactory.getClient("3");

    const apiResponse = await client
      .api("/me/directReports")
      .select("id,userPrincipalName,displayName,jobTitle")
      .get();

    //console.log(meResponse);

    // let personArr: IPerson[] = [];
    // meResponse.value.forEach((item: any) => {

    //   let person: IPerson = {
    //     Id: item.id,
    //     DisplayName: item.displayName,
    //     JobTitle: item.jobTitle,
    //   };

    //   personArr.push(person);
    // });

    //console.log(personArr);

    let reportsArr: IPersonaSharedProps[] = [];
    apiResponse.value.forEach(async (item: any) => {
      const { availability, activity } = await this.getPresence(context,String(item.id));

      let response: IPersonaSharedProps = {
        text: item.displayName,
        secondaryText: item.jobTitle,
        presence: presenceStatus[availability],
        presenceTitle: activity,
      };

      //console.log(response);

      reportsArr.push(response);
    });

    //console.log(reportsArr);

    return reportsArr;
  }

  private getPresence = async (context: WebPartContext,userId: string): Promise<any> => {
    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");

    return await client.api(`users/${userId}/presence`).version("beta").get();
  };
}
