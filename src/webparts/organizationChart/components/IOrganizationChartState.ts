import { IPersonaSharedProps } from "office-ui-fabric-react";
import { IPerson } from "../../../models/IPersonaListProps";

export interface IOrganizationChartState {
    Me: IPersonaSharedProps;
    //Me: IPerson;
    Manager: IPersonaSharedProps;
    Reports: IPersonaSharedProps[];
}