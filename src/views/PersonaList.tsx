import * as React from "react";
import { IPersonaListProps } from "../models/IPersonaListProps";
import { Persona, PersonaPresence, PersonaSize } from "office-ui-fabric-react";

export default class PersonaList extends React.Component<IPersonaListProps,{}> {

  constructor(props: IPersonaListProps) {
    super(props);
  }

  public render(): React.ReactElement<IPersonaListProps> {
    return (
      <>
        {this.props.Personas.map((user) => {
          <Persona
            {...user}
            size={PersonaSize.size48}
            presence={PersonaPresence.none}
          />;
        })}
      </>
    );
  }
}
