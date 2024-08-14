import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/items/get-all";
import { IItemAddResult } from "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IApproverField } from "../Createform";
import { IFileDetails } from "../Createform";
export interface IApproCheck{
    status:string;
    ActionTaken:string
}

export default class spService {
    private _sp;
    constructor(private context: WebPartContext) {
        this._sp = spfi().using(SPFx(this.context))
    }

    public getcolumnInfo = async (listName: string): Promise<{key:string,text:string}[]> => {
        const temp: { key: string; text: string }[] = []
        await this._sp.web.lists.getByTitle(listName).fields.filter(" Hidden eq false and ReadOnlyField eq false")().then(field =>{
            field.filter((value: {
                InternalName?: string; Title: string; TypeDisplayName?: string; Choices?: string[], TypeAsString?: string, SchemaXml?: string
            }) => {
                if (!(value.InternalName.match(/^ApprFld/g) || 
                
                value.InternalName === "Attachments" || 
                value.InternalName === "ContentType" ||
                value.InternalName === "Status" ||
                value.InternalName === "StatusNo" || 
                value.InternalName === "AuditTrail" || 
                value.InternalName ==="BoardApprovers" || 
                value.InternalName ==="PreviousApprovars" || 
                value.InternalName ==="Comments" || 
                value.InternalName ==="PreviousStatus" ||
                value.InternalName ==="StartProcessing"
                )) {
                    // console.log("first");
                    temp.push({
                                key: value.InternalName,
                                text: value.Title})
    
                }
    
            })
        });
        return temp
    }

    public getfieldDetails = async (listName: string): Promise<{ key?: string; text?: string; dataType: string, option?: string[]; internalName: string }[]> => {
        const temp: { key: string; text: string; dataType: string, option?: any[]; internalName: string;DefaultValue?:any;FillInChoice?:boolean }[] = []
        await this._sp.web.lists.getByTitle(listName).fields.filter(" Hidden eq false and ReadOnlyField eq false")().then(field =>{
            console.log(field,"fileds")
            field.filter((value: {
                DefaultValue?: any;
                InternalName: string; Title: string; TypeDisplayName: string; Choices?: string[], TypeAsString: string, SchemaXml?: string;FillInChoice?:boolean
            }) => {
                if (!(value.InternalName.match(/^ApprFld/g) || 
                value.InternalName === "Attachments" || 
                value.InternalName === "ContentType"||
                value.InternalName === "Status" ||
                value.InternalName === "StatusNo" || 
                value.InternalName === "AuditTrail" || 
                value.InternalName ==="BoardApprovers" || 
                value.InternalName ==="PreviousApprovars" || 
                value.InternalName ==="Comments" || 
                value.InternalName ==="PreviousStatus" ||
                value.InternalName ==="StartProcessing"

                )) {
                    // console.log("first"); StartProcessing
                    if (value.TypeAsString === "Choice" && value.SchemaXml.match(/Dropdown/)) {
                        temp.push({
                            key: value.Title,
                            text: value.Title,
                            dataType: "Dropdown",
                            option: value.Choices,
                            internalName: value.InternalName,
                            FillInChoice:value.FillInChoice
                        })
                    }
                    else if (value.TypeAsString === "Choice" || value.TypeAsString === "MultiChoice") {
                        temp.push({
                            key: value.Title,
                            text: value.Title,
                            dataType: value.TypeAsString,
                            option: value.Choices,
                            internalName: value.InternalName,
                            DefaultValue:value.DefaultValue,
                             FillInChoice:value.FillInChoice
                        })
                    } 
                    else if(value.TypeAsString === "Boolean" ){
                        temp.push({
                            key: value.Title,
                            text: value.Title,
                            dataType: value.TypeAsString,
                            option: ["true","false"],
                            internalName: value.InternalName
                        })

                    }
                    else {
                        temp.push({
                            key: value.Title,
                            text: value.Title,
                            dataType: value.TypeAsString,
                            internalName: value.InternalName
                        })
    
                    }
    
                }
    
            })
        });
        // console.log(temp)
        return temp
    }

    public approvarFields = async (listName: string): Promise<{ internalName: string; Title: string }[]> => {
        const apprFields: { internalName: string; Title: string }[] = []
        const field = await this._sp.web.lists.getByTitle(listName).fields.filter(" Hidden eq false and ReadOnlyField eq false")();
        // const apprvConfig = await this._sp.web.lists.getByTitle("SonyDibApprovers").items.getAll();//SonyDibApprovers
        // console.log("field", apprvConfig);
        field.map((ele: {
            Title: string; InternalName: string;
        }) => {
            if (ele.InternalName.match(/^ApprFld/g)) {
                apprFields.push({
                    internalName: ele.InternalName,
                    Title: ele.Title
                });
            }
        });
        return apprFields;
    }

    public submitDataSP = async (listName: string, listObj: Record<string, string|number|string[]|number[]|boolean>): Promise<IItemAddResult> => {
        const iar: IItemAddResult = await this._sp.web.lists.getByTitle(listName).items.add(
            listObj
        );
        return iar;
    }

    // update list Item
    public updateListItem = async (listName: string, updateObj: Record<string, string|number|string[]|number[]|boolean>, Id: number) :Promise<IItemAddResult>=> {
        const iar: IItemAddResult = await this._sp.web.lists.getByTitle(listName).items.getById(Id).update(
            updateObj
        );
        return iar;
    }

    // get items by id
    public getListItemsById = async (listName: string, Id: number):Promise<any> => {
        const iar = await this._sp.web.lists.getByTitle(listName).items.getById(Id).select("*","Author/EMail").expand("Author")();
        return iar;
    }
    public getListItemsByIdBoardApprJson = async (listName: string, Id: number):Promise<{ [x: string]: string  }> => {
        const iar: { [x: string]: string  } = await this._sp.web.lists.getByTitle(listName).items.getById(Id)();
        return iar;
    }

    public uploadAttachemnt = async (file: IFileDetails[], ServerRelativeUrl: string, siteUrl?: string):Promise<void> => {
        //    Get the file from File DOM
        const files = file;
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            //Upload a file to the SharePoint Library
            await this._sp.web.getFolderByServerRelativePath(ServerRelativeUrl).files.addUsingPath(file.name, file.content, { Overwrite: true }).then(res => res);
        }
    }

    public getApproversEmail = async (listTitle: string, apprColumn: string, Id: number):Promise<string[]> => {
        const apprTemp: string[] = []
        apprTemp.push(apprColumn);
        await this._sp.web.lists.getByTitle(listTitle).items.getById(Id).select(`${apprColumn}/EMail`).expand(apprColumn)().then(result => {
            const apprEmail = result.apprColumn;
            for (let i = 0; i < apprEmail.length; i++) {
                apprTemp.push(apprEmail[i].EMail);
            }
        });
        return apprTemp;
    }

    public GetFolder = async (siteUrl: string):Promise<void> => {
        await this._sp.web.getFolderByServerRelativePath(siteUrl)
            .files().then(res => res);
    }

    public apprFieldCheck = (apprinfo: { [x: string]: string | number[]; }, apprFiledInfo: IApproverField[]):IApproCheck[]=> {
        // debugger;
        const apprDetails:IApproCheck[]=[]
        // const apprField = apprFiledInfo.map(x => x.internalName);
        const apprField = apprFiledInfo
        for (let i = 0; i < apprField.length; i++) {
            if (i === 0 && typeof apprinfo[apprField[i].internalName] !== "undefined" && ( apprinfo[apprField[i].internalName] !==null && apprinfo[apprField[i].internalName].length > 0)&& !apprField[i].Disable) {
                apprDetails.push({
                    "status": "1000",
                    // "ActionTaken": " Pending for Related Person 1"
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 1 && typeof apprinfo[apprField[i].internalName] !== "undefined" &&apprinfo[apprField[i].internalName] !==null && apprinfo[apprField[i].internalName].length > 0 && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "2000",
                    // "ActionTaken": "Pending for Related person 2 "
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 2 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0 &&!apprField[i].Disable) {
                apprDetails.push({
                    "status": "3000",
                    // "ActionTaken": "Pending for Related Person 3 "
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 3 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0 && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "4000",
                    // "ActionTaken": "Pending for Related Person 4"
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 4 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0  && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "5000",
                    // "ActionTaken": "Pending for Related person 5 "
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 5 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0 && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "6000",
                    // "ActionTaken": "Pending for  group leader "
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 6 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0 && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "7000",
                    // "ActionTaken": "Pending for  Borad Leader"
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 7 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0 && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "8000",
                    // "ActionTaken": "Pending for ASE "
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
            else if (i === 8 && typeof apprinfo[apprField[i].internalName] !== "undefined" && apprinfo[apprField[i].internalName] !==null &&apprinfo[apprField[i].internalName].length > 0 && !apprField[i].Disable) {
                apprDetails.push({
                    "status": "9000",
                    // "ActionTaken": "Pending for  Chassis Leader"
                    "ActionTaken": `Pending for ${apprField[i].Title}`
                })
                break;
            }
        }
        return apprDetails;
    }
    // get approver configuration
    public apprConfigu=async (listName:string):Promise<IApproverField[]>=>{
        const filterApprfield: { internalName: string; Title: string; COnfigTtle?: string }[] = [];
        const apprFields: {
            internalName: string; Title: string; ConfTitle: string;
        }[] = [];
        const ConfigAppr: IApproverField[] = [];
        const field = await this._sp.web.lists.getByTitle(listName).fields.filter(" Hidden eq false and ReadOnlyField eq false")();
        const apprvConfig = await this._sp.web.lists.getByTitle("Dib_ApprovalConfiguration").items.select("*", "Title,Required,Disable,Approvers/Title").expand("Approvers").getAll();
        // console.log("field", apprvConfig);
        field.map((ele: { InternalName: string; Title: string }, index) => {
            if (ele.InternalName.match(/^ApprFld/g)) {
                filterApprfield.push({
                    internalName: ele.InternalName + "Id",
                    Title: ele.Title,

                    // COnfigTtle:"level"+index+1

                });
            }
        });
        filterApprfield.map((ele, index: number) => {
            apprFields.push({
                internalName: ele.internalName,
                Title: ele.Title,
                ConfTitle: "level" + (index + 1)
            })
        });
        apprvConfig.filter(ele => {
            return apprFields.filter(val => {
                if (ele.Title === val.ConfTitle) {
                    ConfigAppr.push({
                        Title: val.Title,
                        internalName: val.internalName,
                        Required: ele.Required,
                        Email: ele.Approvers,
                        levelType: val.ConfTitle,
                        Disable: ele.Disable
                    })
                }
            })
        });
        return ConfigAppr;
    }
}
