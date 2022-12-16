import { IPersonaSharedProps, PersonaPresence } from "office-ui-fabric-react";

export interface IPersonaListProps {
  Personas: IPersonaSharedProps[];
}

export interface IPerson {
  Email: string;
  DisplayName: string;
  JobTitle: string;
  Id: string;
  Presence: PersonaPresence;
}

export interface IPersonList {
  People: IPerson[];
}
