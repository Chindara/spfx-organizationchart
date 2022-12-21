import { IPersonaSharedProps, PersonaPresence } from "office-ui-fabric-react";

export interface IPersonaListProps {
  Personas: IPersonaSharedProps[];
}

export interface IPerson {
  DisplayName: string;
  JobTitle: string;
  Id: string;
}

export interface IPersonList {
  People: IPerson[];
}
