import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneTextField,
  PropertyPaneButton
} from '@microsoft/sp-property-pane';
import { PropertyFieldCollectionData, CustomCollectionFieldType,
  //  ICustomDropdownOption 
  } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'SonyEdibWebPartStrings';
// import SonyEdib from './components/SonyEdib';
import { ICollection, ISonyEdibProps } from './components/ISonyEdibProps';
import CreateForm from './components/Createform';
import spService from './components/Serivce/spService';
import EditForm from './components/EditForm';
import ViewForm from './components/ViewForm';
// import { IDropdownOption } from 'office-ui-fabric-react';
// import * as ReactDOM from 'react-dom';

export interface ISonyEdibWebPartProps {
  FormType: string;
  description: string,
  listId: any,
  collectionData: ICollection[],
  customOptions: any[],
  siteUrl: string;
  update: string;
  listName: string;
  isBoardApprovalsRequired: boolean;
  libraryName: string;
  DibissuersGroup: string;
  DibGnAdminGroup: string;
  DibApprGroup:string;
  IsAttachmentsRequired:boolean;
  ApproversGroupName:string;
  Refresh:string;
}

export interface listDetails {
  id?: string,
  title?: string,
  webUrl?: string
}
export default class SonyEdibWebPart extends BaseClientSideWebPart<ISonyEdibWebPartProps> {
  private GetFormType = (): string => {
    const params = new URLSearchParams(window.location.search);
    const formType = params.get('FormType');
    // console.log(formType);
    return formType;
  };
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listId: listDetails;
  // private _renderedElement: React.ReactElement<ISonyEdibProps, string | React.JSXElementConstructor<any>>
  private _spService: spService = null;
  public render(): void {
    let element: React.ReactElement<ISonyEdibProps>;
    if (this.properties.FormType === "New") {
      element = React.createElement(CreateForm,
        {
          description: this.properties.description,
          update: "New",
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          listId: this.properties.listId,
          collectionData: this.properties.collectionData,
          customOptions: this.properties.customOptions,
          siteUrl: this.context.pageContext.web.absoluteUrl,
          FormType: this.properties.FormType,
          listName: this.properties.listName,
          isBoardApprovalsRequired: this.properties.isBoardApprovalsRequired,
          libraryName: this.properties.libraryName,
          DibissuersGroup: this.properties.DibissuersGroup,
          DibGnAdminGroup: this.properties.DibGnAdminGroup,
          IsAttachmentsRequired:this.properties.IsAttachmentsRequired,
          ApproversGroupName:this.properties.ApproversGroupName

        }
      );
    }
    else if (this.properties.FormType === "Edit") {
      element = React.createElement(
        EditForm,
        {
          description: this.properties.description,
          update: "Edit",
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          listId: this.properties.listId,
          // multiColumn:string[];
          collectionData: this.properties.collectionData,
          customOptions: this.properties.customOptions,
          siteUrl: this.context.pageContext.web.absoluteUrl,
          FormType: this.properties.FormType,
          listName: this.properties.listName,
          isBoardApprovalsRequired: this.properties.isBoardApprovalsRequired,
          libraryName: this.properties.libraryName,
          DibissuersGroup: this.properties.DibissuersGroup,
          DibGnAdminGroup: this.properties.DibGnAdminGroup,
          IsAttachmentsRequired:this.properties.IsAttachmentsRequired,
          ApproversGroupName:this.properties.ApproversGroupName
        }
      );
    }
    else if (this.properties.FormType === "View") {
      element = React.createElement(
        ViewForm,
        {
          description: this.properties.description,
          update: "View",
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          listId: this.properties.listId,
          // multiColumn:string[];
          collectionData: this.properties.collectionData,
          customOptions: this.properties.customOptions,
          siteUrl: this.context.pageContext.web.absoluteUrl,
          FormType: this.properties.FormType,
          listName: this.properties.listName,
          isBoardApprovalsRequired: this.properties.isBoardApprovalsRequired,
          libraryName: this.properties.libraryName,
          DibissuersGroup: this.properties.DibissuersGroup,
          DibGnAdminGroup: this.properties.DibGnAdminGroup,
          IsAttachmentsRequired:this.properties.IsAttachmentsRequired,
          ApproversGroupName:this.properties.ApproversGroupName
        }
      );

    }
    else {
      return;
    }
    ReactDom.render(element, this.domElement);
    // save the element so we can unmount it later
    // this._renderedElement = element;
  }
  protected async onInit(): Promise<void> {
    await this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });

    this._spService = new spService(this.context);
    // return Promise.resolve();
    // this.properties.customOptions =  this._spService.getcolumnInfo(this._listId.title);
    // await this._spService.getcolumnInfo(this._listId.title);
   
   
    return new Promise<void>((resolve, reject) => {
      if ((this.GetFormType() === "View") || (this.GetFormType() === "Edit")) {
        this.properties.FormType = this.GetFormType();
      }
    
      resolve(undefined);
    });
  }
  protected onPropertyPaneFieldChanged = async (propertyPath: string, oldValue: listDetails, newValue: listDetails): Promise<void> => {
    if (propertyPath === "listId" && newValue) {
      this._listId = newValue;
      this.properties.listName = this._listId.title
      this.properties.customOptions = await this._spService.getcolumnInfo(this._listId.title);
      
      this.render();
      this.context.propertyPane.refresh();
    } 
    else{
      this.properties.customOptions = await this._spService.getcolumnInfo(this.properties.listName);
      this.render();
    }
    this.context.propertyPane.refresh();
  }

  // public getFiledOptions=async ()=>{
  //   console.log(this.properties.collectionData);
  // }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  // protected onDispose(): void {
  //   ReactDom.unmountComponentAtNode(this.domElement);
  // }
  protected onDispose(): void {
    // if (this._renderedElement) {
    ReactDom.unmountComponentAtNode(this.domElement);
    // this._renderedElement = undefined;
    // }
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.render();
   
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              // groupName: strings.BasicGroupName,
              groupName: "",
              groupFields: [
                // PropertyPaneTextField('description', {
                //   label: strings.DescriptionFieldLabel
                // }),
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: true,
                  includeListTitleAndUrl: true,
                  orderBy: PropertyFieldListPickerOrderBy.Id,
                  disabled: false,
                  baseTemplate: 100,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect: false,
                }),
                PropertyPaneTextField('libraryName', {
                  label: "Enter Library Name"
                }), PropertyPaneTextField('DibissuersGroup', {
                  label: "DIB Issuers Group Name"
                }), PropertyPaneTextField('DibGnAdminGroup', {
                  label: "DIB Gen_Admin Group Name"
                }),
                PropertyPaneTextField('ApproversGroupName', {
                  label: "DIB Approvers Group Name"
                }),
                PropertyPaneDropdown('FormType', {
                  label: "FormType",
                  selectedKey: 'New',
                  options: [
                    { key: 'New', text: 'New' },
                    { key: 'View', text: 'View' },
                    { key: 'Edit', text: 'Edit' }

                  ]
                }),

                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  

                  fields: [
                    {
                      id: "FieldName",
                      title: "Field Name",
                      
                      type: CustomCollectionFieldType.dropdown,
                      options: this.properties.customOptions,
                      
                      required: true,
                 

                    },

                    {
                      id: "Tab",
                      title: "Tab",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "DIB Content",
                          text: "DIB Content"
                        },
                        {
                          // key: "Related check items",
                          key: "Related check items",
                          text: "Related check items"
                        },

                      ],
                      required: true
                    },
                    {
                      id: "Required",
                      title: "Required",
                      type: CustomCollectionFieldType.boolean
                    }
                  ],
                  disabled: false
                }),
                PropertyPaneToggle("IsAttachmentsRequired", {
                  label: 'Attachments required'
                }),
                PropertyPaneButton("refresh",{
                  text:"Refresh",
                  buttonType:1,
                  onClick: (value: any) =>{return value}
                })
                // PropertyPaneToggle("isBoardApprovalsRequired", {
                //   label: 'Board approval required'
                // }),
               
              ]
            }
          ]
        }
      ]
    };
  }
  // protected get disableReactivePropertyChanges(): boolean {
  //   return true;
  // }
}
