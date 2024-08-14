import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICustomDropdownOption } from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";

export interface ISonyEdibProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  listId:any;
  update: string;
  libraryName: string;
  // multiColumn:string[];
  collectionData:ICollection[];
  customOptions:ICustomDropdownOption[];
  siteUrl: string;
  FormType:string;
  listName:string;
  isBoardApprovalsRequired:boolean;
  DibissuersGroup:string;
  DibGnAdminGroup:string;
  IsAttachmentsRequired:boolean;
  ApproversGroupName:string;

}
export interface ICollection{
  id: string;
  options: ICustomDropdownOption[];
  uniqueId ?:string,
  FieldName: string, 
  Tab: string,
  Required:boolean, 
  sortIdx?:number
}
