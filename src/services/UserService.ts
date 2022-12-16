import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IPerson } from "../models/IPersonaListProps";
import { IPersonaSharedProps, PersonaPresence } from "office-ui-fabric-react";

export default class UserService {
  constructor() {}

  public async getCurrentUser(context: WebPartContext): Promise<IPersonaSharedProps> {
    const client: MSGraphClientV3 =
      await context.msGraphClientFactory.getClient("3");

    const meResponse = await client
      .api("/me")
      .select("id,userPrincipalName,displayName,jobTitle")
      .get();

    let person: IPerson = {
      Id: meResponse.id,
      DisplayName: meResponse.displayName,
      Email: meResponse.userPrincipalName,
      JobTitle: meResponse.jobTitle,
      Presence: PersonaPresence.none,
    };

    console.log(person);

    let response: IPersonaSharedProps = {
      text: person.DisplayName,
      secondaryText: person.JobTitle,
    };

    this.getPresence(context, String(person.Id));

    return response;
  }

  private getPresence = async (context: WebPartContext,userId: string): Promise<void> => {
    console.log("getPresence");

    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");

    const reportsResponse = await client.api(`/users/${userId}/presence`).version("beta").get();

    console.log(reportsResponse);
  };
}
