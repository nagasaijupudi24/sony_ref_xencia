import * as React from 'react';
import styles from './SonyEdib.module.scss';
import { ISonyEdibProps } from './ISonyEdibProps';
import { Dialog, DialogFooter, DialogType, IPersonaProps, IconButton, Persona, PersonaSize, Pivot, PivotItem, PrimaryButton } from '@fluentui/react';
import spService from './Serivce/spService';
import './custom.css'
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/files";
import "@pnp/sp/profiles";
import "@pnp/sp/site-groups";
import { UserProfileProperties } from './ViewForm';
// import { escape } from '@microsoft/sp-lodash-subset';

export interface IFieldCollection {
    Title: string,
    DataType: string,
    Required: boolean,
    Tab: string,
    internalName: string,
    Option?: string[],
    DefaultValue?: any;
    FillInChoice?: boolean
}

export interface IApproverField {
    Title: string,
    internalName: string,
    Required: boolean,
    Email?: string[],
    levelType?: string,
    Disable?: boolean;
}
export interface IPeoplePickerItems {
    id: number;
    loginName?: string;
    secondaryText?: string;

}
export interface IAppREmails {
    [x: string]: string[];
}
export type Items = string;
export interface IItemsObject {
    [x: string]: string | number | boolean | Items[]
}
export interface IAlertMsg {
    fieldName?: string,
    errorMsg: string
}
export interface IFileDetails {
    name?: string,
    content?: File,
    index?: number,
    fileUrl?: string,
    ServerRelativeUrl?: string,
    isExists?: boolean,
    Modified?: string,
    isSelected?: boolean
}
export interface ISonyEdibState {
    fieldCollection: IFieldCollection[],
    Data: any,
    YesOrNo: any,
    approverFields: IApproverField[];
    approversEmail: IAppREmails;
    configApprover: string[];
    checkboxItems: string[]
    attachfiles: IFileDetails[];
    Approver1: string[];
    Approver2: string[];
    Approver3: string[];
    Approver4: string[];
    Approver5: string[];
    Approver6: string[];
    Approver7: string[];
    Approver8: string[];
    Approver9: string[];
    FullName: string,
    userDepartment: string,
    pictureUrl: string,
    firstName: string,
    lastName: string,
    reqRole: string,
    imageinitials: string;
    hideDialog: boolean;
    AlertMsg: IAlertMsg[];
    isDIBIssuer: boolean;
    ApprGrpUsr: string[];
    isSpecifySelected?: any;
    SpecifyOwnvalues?: any;
    // BoardAppr: any[]
}
export interface IBoardAppr {
    Board: string;
    isUpdated: boolean;
    isCompleted: boolean;
    CCT: { Status: string; isChanged: boolean, value?: number,Comments?:string };
    PWB: { Status: string; isChanged: boolean, value?: number,Comments?:string }

}

export default class CreateForm extends React.Component<ISonyEdibProps, ISonyEdibState, {}> {
    private _spService: spService = null;
    private _checkBoxItems: { [x: string]: string; }[] = [];
    private _sp;
    // private fileInfos: any[];
    private boardApprovers: IBoardAppr[] = [];
    private _userName: string;
    private _UserRole: string;
    private _userpictureUrl: string;
    private _userfirstName: string;
    private _userlastName: string;

    constructor(props: ISonyEdibProps) {
        super(props);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any


        this.state = {
            fieldCollection: [],
            Data: {},
            YesOrNo: {},
            approverFields: [],
            approversEmail: {},
            configApprover: [],
            attachfiles: [],
            checkboxItems: [],
            Approver1: [],
            Approver2: [],
            Approver3: [],
            Approver4: [],
            Approver5: [],
            Approver6: [],
            Approver7: [],
            Approver8: [],
            Approver9: [],
            FullName: "",
            userDepartment: "",
            pictureUrl: "",
            firstName: "",
            lastName: "",
            reqRole: "",
            imageinitials: "",
            hideDialog: true,
            AlertMsg: [],
            isDIBIssuer: false,
            ApprGrpUsr: [],
            isSpecifySelected: {},
            SpecifyOwnvalues: {}
            // BoardAppr: []
        }

        this._sp = spfi().using(SPFx(this.props.context))
        this._spService = new spService(this.props.context);
        this.GetUserProperties().then(res => res).catch(err => err);
        this.getSiteGroups().then(res => res).catch(err => err);
        // this.GetUserProperties().then(res => res).catch(err => err);
        this.formInput().then(res => res).catch(err => err);
        // console.log(this.props.collectionData)
    }

    // // get user details
    private GetUserProperties = async (): Promise<void> => {
        await this._sp.profiles.myProperties().then((result: { UserProfileProperties: UserProfileProperties; DisplayName: string; }) => {
            const props = result.UserProfileProperties;
            this._userName = result.DisplayName;
            for (let i = 0; i < props.length; i++) {
                const allProperties = props[i];
                if (allProperties.Key === "PictureURL") {
                    this._userpictureUrl = allProperties.Value;
                    // console.log(pics);
                }
                else if (allProperties.Key === "Title") {
                    this._UserRole = allProperties.Value;
                    // console.log(reqRole);
                }
                else if (allProperties.Key === "FirstName") {
                    const frstname = allProperties.Value;
                    this._userfirstName = frstname.substring(0, 1);
                    //console.log(firstconstterFN);
                }
                else if (allProperties.Key === "LastName") {
                    const lastname = allProperties.Value;
                    this._userlastName = lastname.substring(0, 1);
                    //console.log(firstletterLN);
                }
            }

            this.setState({
                FullName: this._userName,
                userDepartment: "",
                pictureUrl: this._userpictureUrl,
                firstName: this._userfirstName,
                lastName: this._userlastName,
                reqRole: this._UserRole,
                imageinitials: this._userfirstName + this._userlastName
            });
        }).catch((error: string) => console.log("Error: " + error));
    }

    // Requeter primarytext
    private _onRenderPrimaryText = (props: IPersonaProps): JSX.Element => {
        return (
            <div>
                <span className="userName">{this._userName}</span>
            </div>
        );
    }

    //Requeter secondarytext
    private _onRenderSecondaryText = (props: IPersonaProps): JSX.Element => {
        return (
            <div>
                <div>Role: {this.state.reqRole}</div>
                <span className="" >{this.state.userDepartment}</span>
            </div>
        );
    }

    // get inputs 
    // public formInput = async (): Promise<void> => {
    //     const collectionfileds = this.props.collectionData;
    //     const fieldIfo: { Title: string; DataType: string; Required: boolean; Tab: string; Option: string[]; internalName: string; DefaultValue?: any }[] = [];
    //     if (collectionfileds !== null) {
    //         await this._spService.getfieldDetails(this.props.listName).then((res) => {
    //             // console.log(res);
    //             res.filter(function (newData: { text: string; Title: string; dataType: string; option: string[]; internalName: string; DefaultValue?: any }) {
    //                 return collectionfileds.filter(function (oldData: { FieldName: string; Required: boolean; Tab: string; }) {
    //                     if (newData.internalName === oldData.FieldName) {
    //                         if (newData.dataType === "UserMulti" || newData.dataType === "User") {
    //                             fieldIfo.push({
    //                                 'Title': newData.text,
    //                                 'DataType': newData.dataType,
    //                                 'Required': oldData.Required,
    //                                 'Tab': oldData.Tab,
    //                                 "Option": newData.option,
    //                                 "internalName": newData.internalName + "Id"
    //                             })
    //                         }
    //                         else {
    //                             fieldIfo.push({
    //                                 'Title': newData.text,
    //                                 'DataType': newData.dataType,
    //                                 'Required': oldData.Required,
    //                                 'Tab': oldData.Tab,
    //                                 "Option": newData.option,
    //                                 "internalName": newData.internalName,
    //                                 "DefaultValue": newData.DefaultValue
    //                             })

    //                         }
    //                     }
    //                 })
    //             });
    //             this.setState({ fieldCollection: fieldIfo });
    //         }).catch(error => {
    //             console.log("Something went wrong! please contact admin for more information.", error);
    //         })
    //     }
    // }
    public formInput = async (): Promise<void> => {
        const collectionfileds = this.props.collectionData;
        const fieldIfo: { Title: string; DataType: string; Required: boolean; Tab: string; Option: string[]; internalName: string; DefaultValue?: any; FillInChoice?: boolean }[] = [];
        if (collectionfileds !== null) {
            await this._spService.getfieldDetails(this.props.listName).then((res) => {
                // console.log(res);
                res.filter(function (newData: { text: string; Title: string; dataType: string; option: string[]; internalName: string; DefaultValue?: any; FillInChoice?: boolean }) {
                    return collectionfileds.filter(function (oldData: { FieldName: string; Required: boolean; Tab: string; }) {
                        if (newData.internalName === oldData.FieldName) {
                            if (newData.dataType === "UserMulti" || newData.dataType === "User") {
                                fieldIfo.push({
                                    'Title': newData.text,
                                    'DataType': newData.dataType,
                                    'Required': oldData.Required,
                                    'Tab': oldData.Tab,
                                    "Option": newData.option,
                                    "internalName": newData.internalName + "Id"
                                })
                            }
                            else if (newData.dataType === "Choice" || newData.dataType === "Dropdown") {
                                fieldIfo.push({
                                    'Title': newData.text,
                                    'DataType': newData.dataType,
                                    'Required': oldData.Required,
                                    'Tab': oldData.Tab,
                                    "Option": newData.FillInChoice ? [...newData.option, "Specify your own value :"] : newData.option,
                                    "internalName": newData.internalName,
                                    "DefaultValue": newData.DefaultValue,
                                    "FillInChoice": newData.FillInChoice
                                })

                            }
                            else {
                                fieldIfo.push({
                                    'Title': newData.text,
                                    'DataType': newData.dataType,
                                    'Required': oldData.Required,
                                    'Tab': oldData.Tab,
                                    "Option": newData.option,
                                    "internalName": newData.internalName,
                                    "DefaultValue": newData.DefaultValue,
                                    "FillInChoice": newData.FillInChoice
                                })

                            }
                        }
                    })
                });

            fieldIfo.filter(x=>{
                if(x.DataType==="Choice"&& x.DefaultValue !==null && x.DefaultValue !==undefined){
                    this.setState(prev=>({Data:{...prev.Data,[x.internalName]:x.DefaultValue}}))
                }
            })

                this.setState({ fieldCollection: fieldIfo });
            }).catch(error => {
                console.log("Something went wrong! please contact admin for more information.", error);
            })
        }
    }

    // 
    public resetTheRadioBox = () :void=> {
        this.state.fieldCollection.map(x => {
            if (x.FillInChoice && typeof this.state.Data[x.internalName] !== "undefined" && this.state.Data[x.internalName] !== null) {
                if (x.DataType === "Choice" ||x.DataType === "Dropdown") {
                    if (!x.Option.includes(this.state.Data[x.internalName])) {
                        this.setState(prevState => ({
                            SpecifyOwnvalues: { ...prevState.SpecifyOwnvalues, [x.internalName]: this.state.Data[x.internalName] },
                            Data: { ...prevState.Data, [x.internalName]: "Specify your own value :" }

                        }));
                    }

                }
            }
        });

    }
    // get approvers inputs
    public approvarFields = async (): Promise<void> => {
        await this._spService.apprConfigu(this.props.listName).then(res => this.setState({ approverFields: res })).catch(err => err);
        this.getApprovarInfos().then(res => res).catch(err => err);
    }

    // get Approver in from configuration list
    public getApprovarInfos = async (): Promise<void> => {
        this.state.approverFields.map(async (x, index) => {
            await this._sp.web.lists.getByTitle("Dib_ApprovalConfiguration").items.getById(index + 1).select("*", "Approvers/Id,Approvers/Title").expand("Approvers")().then(res => {
                if (typeof res.Approvers !== "undefined" && !res.Disable) {//LoginName
                    const apprEmails: string[] = []
                    const apprId: number[] = []
                    for (let i = 0; i < res.Approvers.length; i++) {
                        if (this.state.ApprGrpUsr.includes(res.Approvers[i].Id) && index !== 9) {
                            apprEmails.push(res.Approvers[i].Title);
                            apprId.push(res.Approvers[i].Id);
                        }
                        else if (index === 9) {
                            apprEmails.push(res.Approvers[i].Title);
                            apprId.push(res.Approvers[i].Id);

                        }
                    }
                    this.setState(prevState => ({ approversEmail: { ...prevState.approversEmail, [x.internalName]: apprEmails } }));
                    this.setState(prevState => ({ Data: { ...prevState.Data, [x.internalName]: apprId } }));
                }
            });
        });
    }

    // get test ,radio, dropdown values eventHandlerBoolean
    public eventHandler = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
    }
    public SpecifyhandleChxBoxChange = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;

        if (typeof this.state.Data[name] === "undefined") {
            this.setState(prve => ({ Data: { ...prve.Data, [name]: [] } }));
        }
        this.setState((prevState) => {
            const updatedIsSpecifySelected = { ...prevState.isSpecifySelected };
            if (!Object.keys(updatedIsSpecifySelected).includes(name)) {
                updatedIsSpecifySelected[name] = Boolean(value);
            } else {
                updatedIsSpecifySelected[name] = !updatedIsSpecifySelected[name];
            }
            return { isSpecifySelected: updatedIsSpecifySelected };
        });
        // console.log(this.state.Data[name]);
    }

    public SpecifyeventHandler = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        this.setState(prevState => ({ SpecifyOwnvalues: { ...prevState.SpecifyOwnvalues, [name]: value } }));
        // console.log(this.state.SpecifyOwnvalues[name]);
    }
    // get checkbox valuesstartsWith
    public handleChxBoxChange = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        if (name === "Board") {
            // if (name.match(/Board/g)) {
            if (!(this.boardApprovers.some(object => object.Board === value))) {
                this.boardApprovers.push({
                    Board: value,
                    isUpdated: false,
                    isCompleted: false,
                    CCT: { Status: "", isChanged: false,Comments:"" },
                    PWB: { Status: "", isChanged: false ,Comments:""}
                })
            }
            else {
                const filteredArray = this.boardApprovers.filter(obj => obj.Board === value);
                const index = this.boardApprovers.indexOf(filteredArray[0]);
                this.boardApprovers.splice(index, 1);
            }
        }

        const fltrCh = this._checkBoxItems.map(x => x[name])
        if (!fltrCh.includes(value)) {
            this._checkBoxItems.push({
                [name]: value
            });


        } else {

            const itemToRemoveIndex = this._checkBoxItems.findIndex(function (item) {
                return item[name] === value;
            });

            // proceed to remove an item only if it exists.
            if (itemToRemoveIndex !== -1) {
                this._checkBoxItems.splice(itemToRemoveIndex, 1);
            }
        }
        const fltrCheckItems = this._checkBoxItems.map(x => {
            if (typeof x[name] !== "undefined") {
                return x[name]
            }
        })
        const CheckTemp: string[] = []
        fltrCheckItems.filter(ele => {
            if (typeof ele !== "undefined") {
                CheckTemp.push(ele);
            }
        });
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: CheckTemp.sort(), "BoardApprovers": JSON.stringify(this.boardApprovers.sort((a, b) => a.Board.localeCompare(b.Board))) } }));
        // console.log(JSON.stringify(this.boardApprovers.sort((a, b) => a.Board.localeCompare(b.Board))), CheckTemp.sort());
        // this.setState(prevState => ({ Data: { ...prevState.Data, "BoardApprovers": JSON.stringify(this.boardApprovers.sort()) } }));

    }

    // get textarea values
    public handleTextareaChange = (event: React.ChangeEvent<HTMLTextAreaElement>, ind?: number): void => {
        const { name, value } = event.target;
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
    }

    private getSiteGroups = async (): Promise<void> => {
        const tempUsers: any[] = []
        const groups = await this._sp.web.currentUser.groups();
        const gropsTitle = groups.map(x => x.Title);
        if (gropsTitle.includes(this.props.DibissuersGroup)) {
            this.setState({
                isDIBIssuer: true
            });
        }
        // get all users of group
        const users = await this._sp.web.siteGroups.getByName(this.props.ApproversGroupName).users();
        users.map(x => {
            tempUsers.push(x.Id);

        })

        this.setState({ ApprGrpUsr: tempUsers });
        // console.log(tempUsers);
    }

    // get people picker values
    private _getPeoplePickerItems(nm: string, items: IPeoplePickerItems[]): void {
        const apprIds: number[] = [];
        const apprEmails: string[] = []
        const item = items;
        for (let i = 0; i < item.length; i++) {
            //    id ..........
            if (!apprIds.includes(item[i].id)) {
                apprIds.push(item[i].id);
            }
            else {
                const index = apprIds.indexOf(item[i].id);
                if (index > -1) {
                    apprIds.splice(index, 1);
                }
            }

            // emails..................................
            if (!apprEmails.includes(item[i].secondaryText)) {
                apprEmails.push(item[i].secondaryText);
            }
            else {
                const index = apprEmails.indexOf(item[i].secondaryText);
                if (index > -1) {
                    apprEmails.splice(index, 1);
                }
            }
        }
        this.setState(prevState => ({ Data: { ...prevState.Data, [nm]: apprIds } }));
        this.setState(prevState => ({ approversEmail: { ...prevState.approversEmail, [nm]: apprEmails } }));
    }

  
    public setSpecifyValueforCheckBox = (): void => {

        const { fieldCollection, isSpecifySelected, SpecifyOwnvalues, Data } = this.state;

        fieldCollection.forEach(_x => {
            if (_x.FillInChoice && Data[_x.internalName] !== undefined) {
                if (Data[_x.internalName] === "Specify your own value :") {
                    this.setState(prevState => ({
                        Data: {
                            ...prevState.Data,
                            [_x.internalName]: SpecifyOwnvalues[_x.internalName] || null
                        }
                    }));
                }


            }
            if (_x.FillInChoice && Data[_x.internalName] !== undefined && _x.DataType === "MultiChoice") {
                if (!Data[_x.internalName].includes(SpecifyOwnvalues[_x.internalName])) {
                    this.setState(prevState => ({
                        Data: {
                            ...prevState.Data,
                            [_x.internalName]: isSpecifySelected[_x.internalName] ? [...Data[_x.internalName], SpecifyOwnvalues[_x.internalName] || []] : Data[_x.internalName]
                        }
                    }));
                    // console.log(this.state.Data[_x.internalName], this.state.Data[_x.internalName]);
                }
            }
        });
    };

    // Submit data to sharepoint list
    public onSubmitData = async (e: { preventDefault: () => void; }): Promise<void> => {
        await this.setSpecifyValueforCheckBox();
        // console.log(this.state.Data)
        e.preventDefault();

        const emptyFileds: { fieldName?: string; errorMsg: string; }[] = []
        const fltrArry = this.state.fieldCollection.filter(ele => ele.Required);
        const fltrArryApprover = this.state.approverFields.filter(ele => ele.Required && !ele.Disable);
        const isValidFields = fltrArry.map((x) => {
            let isValid = false;
            let isEmpty = false;

            if (Object.keys(this.state.Data).includes(x.internalName)) {
                if (
                    typeof this.state.Data[x.internalName] !== "undefined" &&
                    this.state.Data[x.internalName] !== null &&
                    this.state.Data[x.internalName].length !== 0
                ) {
                    isValid = true;
                }
                // else if(x.FillInChoice){
                //     isValid = true;

                // }
                else {
                    isEmpty = true;
                }
            } else {
                isEmpty = true;
            }

            if (isEmpty) {
                emptyFileds.push({
                    fieldName: x.Title,
                    errorMsg: "Please fill in the below field"
                });
            }
            return isValid;
        });

        const isValidApprFields = fltrArryApprover.map((x) => {
            let isValid = false;
            let isEmpty = false;

            if (Object.keys(this.state.Data).includes(x.internalName)) {
                if (
                    typeof this.state.Data[x.internalName] !== "undefined" &&
                    this.state.Data[x.internalName] !== null &&
                    this.state.Data[x.internalName].length !== 0
                    // this.state.Data[x.internalName].length >= 2
                ) {
                    isValid = true;
                } else {
                    isEmpty = true;
                }
            } else {
                isEmpty = true;
            }

            if (isEmpty) {
                emptyFileds.push({
                    fieldName: x.Title,
                    errorMsg: "Please fill in the below field"
                });
            }

            return isValid;
        });
        let AttachemntsValiadtion: boolean = false;
        if (this.props.IsAttachmentsRequired === true) {
            if (this.state.attachfiles.length > 0) {
                AttachemntsValiadtion = true
            }
            else {
                emptyFileds.push({
                    fieldName: "Upload Attachments",
                    errorMsg: "Please fill in the below field"
                });


            }
        }
        else {
            AttachemntsValiadtion = true
        }


        const isValidation = isValidFields.every((isValid) => isValid);
        const isApprFieldValidation = isValidApprFields.every((isValid) => isValid);
        if (isValidation && isApprFieldValidation && AttachemntsValiadtion) {
            const auditlog: { Actioner: string; ActionTaken: string; Role: string; ActionTakenOn: string; Comments: string; }[] = [];
            const actioner = this._userName;
            const role = this._UserRole;
            const comments = "No Comments";
            const obj = {
                "Actioner": actioner,
                "ActionTaken": "Submitted",
                "Role": role,
                "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                "Comments": comments
            };
            auditlog.push(obj);
            const updateObj = {
                AuditTrail: JSON.stringify(auditlog),
                Status: "submitted",
                StatusNo: "20",
                StartProcessing: true
            }
            // console.log(this.state.Data);
            let isUpdate: boolean = false;
            await this._spService.submitDataSP(this.props.listName, this.state.Data).then(async res => {
                await this.generatereqno(res.data.Id).then(res => res).catch(err => err);
                await this._spService.updateListItem(this.props.listName, updateObj, res.data.Id).then(res => {
                    // console.log(res);
                    isUpdate = true;

                }).catch(err => err);
                if (isUpdate === true) {
                    this.apprCheckingUpdate(res.data.Id).then(res => res).catch(err => err);
                }
                emptyFileds.push({
                    errorMsg: "Request has been submitted successfully",
                    fieldName: ""
                });
                this._checkBoxItems.length = 0;
               this.boardApprovers.length=0;
                this.setState({ attachfiles: [], approversEmail: {}, checkboxItems: [], Data: {} });
                // debugger;
                this.getApprovarInfos().then(res => res).catch(err => err);
                // this.approvarFields().then(res => res).catch(err => err);
                // this.formInput().then(res => res).catch(err => err);


            }).catch(err => err);

        }
        this.setState({ AlertMsg: emptyFileds, hideDialog: false });


    }

    // Draft
    public onDraftData = async (e: { preventDefault: () => void; }): Promise<void> => {
        await this.setSpecifyValueforCheckBox();
        e.preventDefault();

        const emptyFileds: { fieldName?: string; errorMsg: string; }[] = []

        // const fltrArry = this.state.fieldCollection.filter(ele => ele.Required);
        // const fltrArryApprover = this.state.approverFields.filter(ele => ele.Required && !ele.Disable);
        // // const emptyFileds: { fieldName?: string; errorMsg: string; }[] = []
        // const isValidFields = fltrArry.map((x) => {
        //     let isValid = false;
        //     let isEmpty = false;

        //     if (Object.keys(this.state.Data).includes(x.internalName)) {
        //         if (
        //             typeof this.state.Data[x.internalName] !== "undefined" &&
        //             this.state.Data[x.internalName] !== null &&
        //             this.state.Data[x.internalName].length !== 0
        //         ) {
        //             isValid = true;
        //         } else {
        //             isEmpty = true;
        //         }
        //     } else {
        //         isEmpty = true;
        //     }

        //     if (isEmpty) {
        //         emptyFileds.push({
        //             fieldName: x.Title,
        //             errorMsg: "Please fill in the below field"
        //         });
        //     }

        //     return isValid;
        // });
        // const isValidApprFields = fltrArryApprover.map((x) => {
        //     let isValid = false;
        //     let isEmpty = false;

        //     if (Object.keys(this.state.Data).includes(x.internalName)) {
        //         if (
        //             typeof this.state.Data[x.internalName] !== "undefined" &&
        //             this.state.Data[x.internalName] !== null &&
        //             this.state.Data[x.internalName].length !== 0
        //         ) {
        //             isValid = true;
        //         } else {
        //             isEmpty = true;
        //         }
        //     } else {
        //         isEmpty = true;
        //     }

        //     if (isEmpty) {
        //         emptyFileds.push({
        //             fieldName: x.Title,
        //             errorMsg: "Please fill in the below field"
        //         });
        //     }

        //     return isValid;
        // });
        // // attachemnts valiadtion;
        // let AttachemntsValiadtion:boolean=false;
        // if(this.props.IsAttachmentsRequired === true){
        //     if(this.state.attachfiles.length>0){
        //         AttachemntsValiadtion=true
        //     }
        //     else{
        //         emptyFileds.push({
        //             fieldName: "Upload Attachments",
        //             errorMsg: "Please fill in the below field"
        //         });


        //     }
        // }
        // else{
        //     AttachemntsValiadtion=true
        // }

        // const isValidation = isValidFields.every((isValid) => isValid);
        // const isApprFieldValidation = isValidApprFields.every((isValid) => isValid);
        // if (isValidation && isApprFieldValidation && AttachemntsValiadtion) {
        const auditlog: { Actioner: string; ActionTaken: string; Role: string; ActionTakenOn: string; Comments: string; }[] = [];
        const actioner = this._userName;
        const role = this._UserRole;
        const comments = "No Comments";
        const obj = {
            "Actioner": actioner,
            "ActionTaken": "Drafted",
            "Role": role,
            "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
            "Comments": comments
        };
        auditlog.push(obj);
        const updateObj = {
            AuditTrail: JSON.stringify(auditlog),
            Status: "Drafted",
            StatusNo: "5",
            StartProcessing: true
        }
        // console.log(this.state.Data);
        // let isUpdate: boolean = false;
        await this._spService.submitDataSP(this.props.listName, this.state.Data).then(async res => {
            await this.generatereqno(res.data.Id).then(res => res).catch(err => err);
            await this._spService.updateListItem(this.props.listName, updateObj, res.data.Id).then(res => {
                // console.log(res);
                // isUpdate = true;
                return res

            }).catch(err => err);
            emptyFileds.push({
                errorMsg: "Request has been drafted successfully",
                fieldName: ""
            });
            this._checkBoxItems.length = 0;
            this.boardApprovers.length=0;
            this.setState({ attachfiles: [], approversEmail: {}, checkboxItems: [], Data: {},
            
         });
            this.getApprovarInfos().then(res => res).catch(err => err);
            // this.approvarFields().then(res => res).catch(err => err);
            // this.formInput().then(res => res).catch(err => err);

        }).catch(err => {
            console.log("err", err);

        });
        this.setState({ AlertMsg: emptyFileds, hideDialog: false });
        // }

    }

    public apprCheckingUpdate = async (Id: number): Promise<void> => {
        const apprCheckingIsEmtyOrNot = this._spService.apprFieldCheck(this.state.Data, this.state.approverFields);
        if (apprCheckingIsEmtyOrNot.length > 0) {
            const varStatus = apprCheckingIsEmtyOrNot[0].status;
            const varActionTaken = apprCheckingIsEmtyOrNot[0].ActionTaken;
            const updateObj = {
                Status: varActionTaken,
                StatusNo: varStatus,
                StartProcessing: true
            }
            await this._spService.updateListItem(this.props.listName, updateObj, Id).then(res => res).catch(err => err);
        }
    }

    // Generate request number and update
    private generatereqno = async (Id: number): Promise<void> => {
        const reqformat = "eDIB";
        const reqid = 100000 + Id
        const res = reqid.toString().substring(1, 6);
        const reqno = reqformat + res;
        // console.log("reqno", reqno);
        // const listUpdtObj = { Title: reqno, "DIB_ID": Id };
        const listUpdtObj = { Title: reqno };
        await this._spService.updateListItem(this.props.listName, listUpdtObj, Id).then(res => {
            this.createFolder(reqno).then(res => res).catch(err => err);
            return res
        }).catch(err => err);
    }

    // // create folder under document library
    public createFolder = async (itemId: string): Promise<void> => {
        const tenantUrl = window.location.protocol + "//" + window.location.host;
        const siteUrl = this.props.siteUrl.replace(tenantUrl, "");
        await this._sp.web.rootFolder.folders.addUsingPath(siteUrl + `/${this.props.libraryName}/` + itemId).then(async res => {
            await this._spService.uploadAttachemnt(this.state.attachfiles, res.data.ServerRelativeUrl, siteUrl).then(res => res).catch(err => err);
        })
    }
    // add Attachments
    private addAttacment = async (): Promise<void> => {
        const fileInfo: { name: string; content: File; index: number; fileUrl: string; ServerRelativeUrl: string; isExists: boolean; Modified: string; isSelected: boolean; }[] = [];
        const fileInput = document.getElementById('Docfiles') as HTMLInputElement;
        const fileCount = fileInput.files.length;
        for (let i = 0; i < fileCount; i++) {
            // const file = fileInput["files"][i];
            const file = fileInput.files[i];
            const filesId = Math.floor((Math.random() * 1000000000) + 1);
            const reader = new FileReader();
            reader.onload = ((file) => {
                return (e) => {
                    //Push the converted file into array
                    e.preventDefault();
                    const isObjectExists = this.state.attachfiles.map((obJ: { name: string; }) => obJ.name);
                    if (!isObjectExists.includes(file.name)) {
                        fileInfo.push({
                            "name": file.name,
                            "content": file,
                            "index": filesId,
                            "fileUrl": "",
                            "ServerRelativeUrl": "",
                            "isExists": false,
                            "Modified": new Date().toISOString(),
                            "isSelected": false
                        });
                    }
                    this.setState({ attachfiles: [...this.state.attachfiles, ...fileInfo] });
                    // this.fileInfos.push(fileInfo);
                };
            })(file);
            reader.readAsArrayBuffer(file);
        }
    }
    // Remove Attachemnts
    public onRemoveAttachments = (file: IFileDetails): void => {
        // debugger;
        // console.log(file)
        const { attachfiles } = this.state;
        const fltrArry = attachfiles.filter(obj => obj.index !== file.index);
        // const index = attachfiles.indexOf(fltrArry[0]);
        // if (index > -1) {
        //     attachfiles.splice(index, 1);
        // }
        this.setState({ attachfiles: fltrArry });
        // alert(index);
    }
    public toggleHideDialog = (): void => {
        this.setState({ hideDialog: false });
    }
    // homapage /SitePages/DibNew.aspx MyDIBRequest
    public homePage = (): void => {
        const pageURL: string = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/mydibrequest.aspx`;
        window.location.href = pageURL;
        // window.location.reload();
        this.setState({ hideDialog: true });
    }

    public onCancel = (): void => {
        this.state.AlertMsg.map(x => {
            if (x.fieldName === "") {
                this.homePage();
            }
            else {
                this.setState({ hideDialog: true });
                this.resetTheRadioBox();
            }
        })
     
        // this.setState({ hideDialog: true });
        //         this.resetTheRadioBox();
    }

    public componentDidMount = (): void => {
        this.approvarFields().then(res => res).catch(err => err);
        // this.formInput().then(res => res).catch(err => err);
    }

    public render(): React.ReactElement<ISonyEdibProps> {
        const {
            hasTeamsContext,
        } = this.props;

        const dialogContentProps = {
            type: DialogType.normal,
            title: 'Information!',
        }
        if (this.state.isDIBIssuer === true) {
            return (
                <section className={`${styles.sonyEdib} ${hasTeamsContext ? styles.teams : ''}`} >
                    <div className={styles.sonydibcontainer}>
                        <div className={styles.frmtitle}>New Request Form</div>
                        <Pivot aria-label="Large Link Size Pivot Example" >
                            <PivotItem linkText="Requestor Details">
                                <fieldset style={{ height: 60, "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className="ms-Grid-row personarow" style={{ paddingTop: 5 }}>
                                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6 personaCol1">
                                            <div className="ms-Grid-col ms-sm6 ms-lg6">
                                                <Persona
                                                    //{...personaWithInitials}                        
                                                    imageUrl={this.state.pictureUrl}
                                                    size={PersonaSize.size48}
                                                    imageInitials={this.state.imageinitials}
                                                    onRenderPrimaryText={this._onRenderPrimaryText}
                                                    onRenderSecondaryText={this._onRenderSecondaryText}
                                                />
                                            </div>
                                        </div>
                                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6 personaCol2">
                                            <div className="ms-Grid-row presonaCol2Row" >
                                                <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg2 presonaCol2col1">
                                                    <span className="hdrTtle">Status: </span>
                                                </div>
                                                <div className="ms-Grid-col ms-sm8 ms-md9 ms-lg10 presonaCol2col2">
                                                    <span>New</span>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </fieldset>
                            </PivotItem>
                        </Pivot>
                        < form >
                            <Pivot aria-label="Large Link Size Pivot Example" >
                                <PivotItem headerText="DIB content">
                                    <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                        <div className='conationer'>
                                            <div className='ms-Grid'>
                                                <div className='ms-Grid-row'>
                                                    <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                            {/* <form action=""> */}
                                                            {this.state.fieldCollection.map((value, index) => {
                                                                if (value.Tab === "DIB Content") {
                                                                    switch (value.DataType) {
                                                                        case "Text":
                                                                            return (
                                                                                <>
                                                                                    {(value.internalName === "Title" || value.internalName === "DIB_ID") ? null :
                                                                                        <div className='fieldEditor' key={index}>
                                                                                            <div>
                                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                            </div>
                                                                                            <div className='feildDisplay'>
                                                                                                <div className='fieldGroup'>
                                                                                                    <input
                                                                                                        type="text"
                                                                                                        name={value.internalName}
                                                                                                        id={value.internalName}
                                                                                                        value={Object.keys(this.state.Data).includes(value.internalName) ? this.state.Data[value.internalName] : ""}
                                                                                                        required={value.Required}
                                                                                                        // style={{ border: "none", outline: "none" }}
                                                                                                        className='inputText'
                                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                                    />
                                                                                                </div>
                                                                                            </div>
                                                                                        </div>
                                                                                    }
                                                                                </>


                                                                            )
                                                                        // break;
                                                                        case "Note":
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>
                                                                                        {/* <div className='fieldGroup'> */}
                                                                                        <textarea
                                                                                            name={value.internalName}
                                                                                            id={value.internalName}
                                                                                            value={this.state.Data[value.internalName] || ""}
                                                                                            style={{ border: "1px solid black", outline: "none" }}
                                                                                            required={value.Required}
                                                                                            rows={3}
                                                                                            className="textarea"
                                                                                            onChange={(event) => this.handleTextareaChange(event, index)}
                                                                                        />
                                                                                    </div>
                                                                                    {/* </div> */}
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "MultiChoice":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        {/* <label className="label" htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='radioCheckbox'>
                                                                                        {value.Option.map((ele: string) => {
                                                                                            return (
                                                                                                // <label htmlFor={value.internalName} key={index}>{ele}
                                                                                                <div key={index} className={styles.container}>

                                                                                                    <input type="checkbox" name={value.internalName} value={ele}
                                                                                                        onChange={(event) => this.handleChxBoxChange(event, index)}
                                                                                                        checked={this._checkBoxItems.some(obj => obj[value.internalName] === ele)}
                                                                                                        required={value.Required}
                                                                                                    />
                                                                                                    <span className={styles.checkmark} />
                                                                                                    {ele}

                                                                                                </div>
                                                                                                // </label>
                                                                                            )
                                                                                        })}
                                                                                        {/* specify own values---------- */}
                                                                                        {value.FillInChoice ? <>
                                                                                            <div className='radioCheckbox'>
                                                                                                <div key={index} className={styles.container}>
                                                                                                    <input type="checkbox" name={value.internalName} value={"true"}
                                                                                                        onChange={(event) => this.SpecifyhandleChxBoxChange(event, index)}
                                                                                                        checked={this.state.isSpecifySelected[value.internalName]}
                                                                                                    // required={value.Required}

                                                                                                    />
                                                                                                    <span className={styles.checkmark} />{"Specify your own value :"}


                                                                                                </div>
                                                                                            </div>
                                                                                            <input
                                                                                                type="text"
                                                                                                name={value.internalName}
                                                                                                // id={value.internalName}
                                                                                                disabled={!value.FillInChoice}
                                                                                                value={this.state.isSpecifySelected[value.internalName]?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                                style={{ outline: "none"}}
                                                                                                className={styles.eDIBSpecifyOwnvalueInput}
                                                                                                // onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                                onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                            />
                                                                                        </>
                                                                                            : null}
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "Choice":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='radioCheckbox'>
                                                                                        {value.Option.map((ele: string | number | readonly string[]) => {
                                                                                            return (
                                                                                                <div key={index} className={styles.radiocontainer}>
                                                                                                    {/* <input type="radio" name={value.internalName} value={ele} onChange={(event) => this.eventHandler(event, index)} required={value.Required} />{ele} */}
                                                                                                    <input type="radio" name={value.internalName} value={ele}
                                                                                                        onChange={(event) => this.eventHandler(event, index)}

                                                                                                        checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                        required={value.Required} />
                                                                                                    <div className={styles.radiocheckmark}>
                                                                                                        <span className={styles.radioinsidecircle} />
                                                                                                    </div>
                                                                                                    {ele}
                                                                                                </div>
                                                                                            )
                                                                                        })}
                                                                                        {/* specify own values---------- */}
                                                                                        {value.FillInChoice ?
                                                                                            <input
                                                                                                type="text"
                                                                                                name={value.internalName}
                                                                                                // id={value.internalName}
                                                                                                disabled={!value.FillInChoice}
                                                                                                // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                                value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                                // value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":this.setState(prev=>({SpecifyOwnvalues:{...prev.SpecifyOwnvalues,[value.internalName]:""}}))}
                                                                                                // value={this.state.isSpecifySelected[value.internalName]?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                                style={{ outline: "none"}}
                                                                                                className={styles.eDIBSpecifyOwnvalueInput}
                                                                                                // onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                                onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                            />

                                                                                            : null}
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                        case "Boolean":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='radioCheckbox'>
                                                                                        {value.Option.map((ele: any) => {
                                                                                            return (
                                                                                                <div key={index} className={styles.radiocontainer}>
                                                                                                    {/* <input type="radio" name={value.internalName} value={ele} onChange={(event) => this.eventHandler(event, index)} required={value.Required} />{ele} */}
                                                                                                    <input type="radio" name={value.internalName} value={ele}
                                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                                        checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                        required={value.Required} />
                                                                                                    <div className={styles.radiocheckmark}>
                                                                                                        <span className={styles.radioinsidecircle} />
                                                                                                    </div>
                                                                                                    {ele === "true" ? "Yes" : "No"}
                                                                                                </div>
                                                                                            )
                                                                                        })}
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        case "Dropdown":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        {/* <label className="label" htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div style={{ width: "50%" }}>
                                                                                        <select className='dropdwon'
                                                                                            id={value.internalName}
                                                                                            name={value.internalName}
                                                                                            value={this.state.Data[value.internalName] || ""}
                                                                                            onChange={(event) => this.eventHandler(event, index)}
                                                                                            placeholder='Select option'
                                                                                        >
                                                                                            <option>{""}</option>
                                                                                            {value.Option.map((ele: string) => {
                                                                                                return (<option key={ele}>{ele}</option>)
                                                                                            })}

                                                                                        </select>
                                                                                        {value.FillInChoice ? <input
                                                                                            type="text"
                                                                                            name={value.internalName}
                                                                                            // id={value.internalName}
                                                                                            disabled={!value.FillInChoice}
                                                                                            // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                            // value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":this.setState(prev=>({SpecifyOwnvalues:{...prev.SpecifyOwnvalues,[value.internalName]:""}}))}
                                                                                            value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":""}

                                                                                            // value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":this.setState(prev=>({SpecifyOwnvalues:{...prev.SpecifyOwnvalues,[value.internalName]:""}}))}
                                                                                            style={{ outline: "none"}}
                                                                                            className={styles.eDIBSpecifyOwnvalueInputdrp}
                                                                                            // onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                            onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                        /> : null}
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "DateTime":
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="date"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                required={value.Required}
                                                                                                value={this.state.Data[value.internalName] || ""}
                                                                                                // style={{ border: "none", outline: "none", width: "100%" }}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "Number":
                                                                            return (
                                                                                <div className='fieldEditor' key={index} hidden={false}>
                                                                                    <div>
                                                                                        {/* <label className='label' htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="number"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                required={value.Required}
                                                                                                value={this.state.Data[value.internalName] || ""}
                                                                                                // style={{ border: "none", outline: "none", width: "100%" }}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "UserMulti"
                                                                            :
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        {/* <label className='label' htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='peoplepicker'>
                                                                                            <PeoplePicker
                                                                                                context={this.props.context}
                                                                                                // titleText={value.Title}
                                                                                                personSelectionLimit={3}
                                                                                                groupName={""} // Leave this blank in case you want to filter from all users
                                                                                                showtooltip={true}
                                                                                                required={value.Required}
                                                                                                disabled={false}
                                                                                                defaultSelectedUsers={Object.keys(this.state.approversEmail).includes(value.internalName) ? this.state.approversEmail[value.internalName] : null}
                                                                                                onChange={this._getPeoplePickerItems.bind(this, value.internalName)}
                                                                                                showHiddenInUI={false}
                                                                                                ensureUser={true}
                                                                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                                                                resolveDelay={1000} />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>

                                                                            )
                                                                        // break;
                                                                        case "User":
                                                                            return (
                                                                                <div>
                                                                                    <PeoplePicker
                                                                                        context={this.props.context}
                                                                                        titleText={value.Title}
                                                                                        personSelectionLimit={3}
                                                                                        groupName={""} // Leave this blank in case you want to filter from all users
                                                                                        showtooltip={true}
                                                                                        required={value.Required}
                                                                                        defaultSelectedUsers={this.state.approversEmail[value.internalName]}
                                                                                        // disabled={true}
                                                                                        onChange={this._getPeoplePickerItems.bind(this, value.internalName)}
                                                                                        showHiddenInUI={false}
                                                                                        principalTypes={[PrincipalType.User]}
                                                                                        resolveDelay={1000} />
                                                                                </div>
                                                                            )
                                                                        default:
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        {/* <label className='label' htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="test"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                value={this.state.Data[value.internalName] || ""}
                                                                                                required={value.Required}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                    }
                                                                }
                                                            })}
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </fieldset>
                                </PivotItem>
                                <PivotItem headerText="Related Check Items">
                                    <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                        <div className='conationer'>
                                            <div className='ms-Grid'>
                                                <div className='ms-Grid-row'>
                                                    <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                            {/* <form action=""> */}
                                                            {this.state.fieldCollection.map((value, index) => {
                                                                if (value.Tab === "Related check items") {
                                                                    switch (value.DataType) {
                                                                        case "Text":
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>
                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="text"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                value={Object.keys(this.state.Data).includes(value.internalName) ? this.state.Data[value.internalName] : ""}
                                                                                                required={value.Required}
                                                                                                // style={{ border: "none", outline: "none" }}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "Note":
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>
                                                                                        <textarea
                                                                                            name={value.internalName}
                                                                                            id={value.internalName}
                                                                                            value={this.state.Data[value.internalName] || ""}
                                                                                            style={{ border: "1px solid black", outline: "none" }}
                                                                                            required={value.Required}
                                                                                            rows={3}
                                                                                            className="textarea"
                                                                                            onChange={(event) => this.handleTextareaChange(event, index)}
                                                                                        />
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        case "MultiChoice":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='radioCheckbox'>
                                                                                        {value.Option.map((ele: string) => {
                                                                                            return (
                                                                                                <div key={index} className={styles.container}>
                                                                                                    <input type="checkbox" name={value.internalName} value={ele}
                                                                                                        onChange={(event) => this.handleChxBoxChange(event, index)}
                                                                                                        checked={this._checkBoxItems.some(obj => obj[value.internalName] === ele)}
                                                                                                        required={value.Required}
                                                                                                    />
                                                                                                    <span className={styles.checkmark} />
                                                                                                    {ele}
                                                                                                </div>
                                                                                            )
                                                                                        })}
                                                                                        {/* specify own values---------- */}
                                                                                        {value.FillInChoice ? <>
                                                                                            <div className='radioCheckbox'>
                                                                                                <div key={index} className={styles.container}>
                                                                                                    <input type="checkbox" name={value.internalName} value={"true"}
                                                                                                        onChange={(event) => this.SpecifyhandleChxBoxChange(event, index)}
                                                                                                        checked={this.state.isSpecifySelected[value.internalName]}
                                                                                                    // required={value.Required}

                                                                                                    />
                                                                                                    <span className={styles.checkmark} />{"Specify your own value :"}


                                                                                                </div>
                                                                                            </div>
                                                                                            <input
                                                                                                type="text"
                                                                                                name={value.internalName}
                                                                                                // id={value.internalName}
                                                                                                disabled={!value.FillInChoice}
                                                                                                // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                                value={this.state.isSpecifySelected[value.internalName]?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                                style={{ outline: "none" }}
                                                                                                className={styles.eDIBSpecifyOwnvalueInput}
                                                                                                // onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                                onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                            />
                                                                                        </>
                                                                                            : null}
                                                                                    </div>
                                                                                </div>
                                                                            )

                                                                        // break;
                                                                        case "Choice":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='radioCheckbox'>
                                                                                        {value.Option.map((ele: string | number | readonly string[]) => {
                                                                                            return (
                                                                                                <div key={index} className={styles.radiocontainer}>
                                                                                                    <input type="radio" name={value.internalName} value={ele}
                                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                                        // defaultChecked={(value.DefaultValue !== null && value.DefaultValue === ele) ? true : false}
                                                                                                        // checked={(this.state.Data[value.internalName] === ele ||value.DefaultValue ===ele &&value.DefaultValue !==null) ? true : false}
                                                                                                        checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                        required={value.Required} />
                                                                                                    <div className={styles.radiocheckmark}>
                                                                                                        <span className={styles.radioinsidecircle} />
                                                                                                    </div>
                                                                                                    {ele}
                                                                                                </div>
                                                                                            )
                                                                                        })}
                                                                                        {/* specify own values---------- */}
                                                                                        {value.FillInChoice ?
                                                                                            <input
                                                                                                type="text"
                                                                                                name={value.internalName}
                                                                                                // id={value.internalName}
                                                                                                disabled={!value.FillInChoice}
                                                                                                // value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":this.setState(prev=>({SpecifyOwnvalues:{...prev.SpecifyOwnvalues,[value.internalName]:""}}))}
 
                                                                                                value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                                // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                                // value={this.state.isSpecifySelected[value.internalName]?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                                style={{ outline: "none"}}
                                                                                                className={styles.eDIBSpecifyOwnvalueInput}
                                                                                                // className='inputText'
                                                                                                // onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                                onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                            />

                                                                                            : null}
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        case "Boolean":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='radioCheckbox'>
                                                                                        {value.Option.map((ele: any) => {
                                                                                            return (
                                                                                                <div key={index} className={styles.radiocontainer}>
                                                                                                    {/* <input type="radio" name={value.internalName} value={ele} onChange={(event) => this.eventHandler(event, index)} required={value.Required} />{ele} */}
                                                                                                    <input type="radio" name={value.internalName} value={ele}
                                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                                        checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                        required={value.Required} />
                                                                                                    <div className={styles.radiocheckmark}>
                                                                                                        <span className={styles.radioinsidecircle} />
                                                                                                    </div>
                                                                                                    {ele === "true" ? "Yes" : "No"}
                                                                                                </div>
                                                                                            )
                                                                                        })}
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // case "Boolean":
                                                                        //     return (
                                                                        //         <div className='fieldEditor'>
                                                                        //             <div>
                                                                        //                 <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                        //             </div>
                                                                        //             <div className='radioCheckbox'>
                                                                        //                 {value.Option.map((ele: string | number | readonly string[]) => {
                                                                        //                     return (
                                                                        //                         <div key={index} className={styles.radiocontainer}>
                                                                        //                             <input type="radio" name={value.internalName} value={ele}
                                                                        //                                 onChange={(event) => this.eventHandlerBoolean(event, index)}
                                                                        //                                 checked={this.state.YesOrNo[value.internalName] === ele ? true : false}
                                                                        //                                 required={value.Required} />
                                                                        //                             <div className={styles.radiocheckmark}>
                                                                        //                                 <span className={styles.radioinsidecircle} />
                                                                        //                             </div>
                                                                        //                             {ele}
                                                                        //                         </div>
                                                                        //                     )
                                                                        //                 })}
                                                                        //             </div>
                                                                        //         </div>
                                                                        //     )

                                                                        case "Dropdown":
                                                                            return (
                                                                                <div className='fieldEditor'>
                                                                                    <div>
                                                                                        {/* <label className="label" htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div style={{ width: "50%" }}>
                                                                                        <select className='dropdwon'
                                                                                            id={value.internalName}
                                                                                            name={value.internalName}
                                                                                            value={this.state.Data[value.internalName] || ""}
                                                                                            onChange={(event) => this.eventHandler(event, index)}
                                                                                        >
                                                                                            {value.Option.map((ele: string) => {
                                                                                                return (<option key={ele}>{ele}</option>)
                                                                                            })}

                                                                                        </select>
                                                                                        {value.FillInChoice ? <input
                                                                                            type="text"
                                                                                            name={value.internalName}
                                                                                            // id={value.internalName}
                                                                                            disabled={!value.FillInChoice}
                                                                                            // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                            value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                            style={{ outline: "none"}}
                                                                                            className={styles.eDIBSpecifyOwnvalueInputdrp}
                                                                                            // className='inputText'
                                                                                            // onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                            onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                        /> : null}

                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "DateTime":
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="date"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                required={value.Required}
                                                                                                value={this.state.Data[value.internalName] || ""}
                                                                                                // style={{ border: "none", outline: "none", width: "100%" }}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "Number":
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        {/* <label className='label' htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>
                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="number"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                required={value.Required}
                                                                                                value={this.state.Data[value.internalName] || ""}
                                                                                                // style={{ border: "none", outline: "none", width: "100%" }}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                            )
                                                                        // break;
                                                                        case "UserMulti"
                                                                            :
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        {/* <label className='label' htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='peoplepicker'>
                                                                                            <PeoplePicker
                                                                                                context={this.props.context}
                                                                                                // titleText={value.Title}
                                                                                                personSelectionLimit={3}
                                                                                                groupName={""} // Leave this blank in case you want to filter from all users
                                                                                                showtooltip={true}
                                                                                                required={value.Required}
                                                                                                disabled={false}
                                                                                                defaultSelectedUsers={Object.keys(this.state.approversEmail).includes(value.internalName) ? this.state.approversEmail[value.internalName] : null}
                                                                                                onChange={this._getPeoplePickerItems.bind(this, value.internalName)}
                                                                                                showHiddenInUI={false}
                                                                                                ensureUser={true}
                                                                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                                                                resolveDelay={1000} />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>

                                                                            )
                                                                        // break;
                                                                        case "User":
                                                                            return (
                                                                                <div>
                                                                                    <PeoplePicker
                                                                                        context={this.props.context}
                                                                                        titleText={value.Title}
                                                                                        personSelectionLimit={3}
                                                                                        groupName={""} // Leave this blank in case you want to filter from all users
                                                                                        showtooltip={true}
                                                                                        required={value.Required}
                                                                                        defaultSelectedUsers={Object.keys(this.state.approversEmail).includes(value.internalName) ? this.state.approversEmail[value.internalName] : null}
                                                                                        // disabled={true}
                                                                                        onChange={this._getPeoplePickerItems.bind(this, value.internalName)}
                                                                                        showHiddenInUI={false}
                                                                                        principalTypes={[PrincipalType.User]}
                                                                                        resolveDelay={1000} />
                                                                                </div>
                                                                            )
                                                                        default:
                                                                            return (
                                                                                <div className='fieldEditor' key={index}>
                                                                                    <div>
                                                                                        {/* <label className='label' htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                        <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                    </div>
                                                                                    <div className='feildDisplay'>

                                                                                        <div className='fieldGroup'>
                                                                                            <input
                                                                                                type="text"
                                                                                                name={value.internalName}
                                                                                                id={value.internalName}
                                                                                                value={this.state.Data[value.internalName] || ""}
                                                                                                required={value.Required}
                                                                                                className='inputText'
                                                                                                onChange={(event) => this.eventHandler(event, index)}
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </div>

                                                                            )
                                                                    }
                                                                }
                                                            })}
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </fieldset>
                                </PivotItem>
                                <PivotItem headerText="Attachments">
                                    <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                        <div className='fieldEditor'>
                                            <div style={{ width: "50%" }}>
                                                <label className='label' htmlFor='Attachments'>Upload Attachments {this.props.IsAttachmentsRequired === true ? <span className='labelIcon'>*</span> : null} </label>
                                            </div>
                                            {/* <div className='feildDisplay'> */}
                                            <div className='fieldGroupAttch' style={{ width: "50%" }}>
                                                <input
                                                    type="file"
                                                    name='files'
                                                    id="Docfiles"
                                                    onChange={this.addAttacment}
                                                    // required
                                                    style={{ border: "none", outline: "none" }}
                                                />
                                            </div>
                                            {/* </div> */}
                                        </div>
                                        <div className={styles.viewfeildFilesDisplay}>
                                            <span>
                                                <ul style={{ margin: "unset", padding: "unset" }}>
                                                    {this.state.attachfiles.map((file, ind) =>
                                                        <div className={styles.viewfeildFilesDisplay} key={ind}>
                                                            <span className={styles.attachmentSpan}>
                                                                <IconButton
                                                                    iconProps={{ iconName: 'Delete' }}
                                                                    onClick={() => this.onRemoveAttachments(file)}
                                                                    styles={{
                                                                        icon: { fontSize: 18 },
                                                                        root: {
                                                                            // width: 100,
                                                                            // height: 100,
                                                                            // backgroundColor: 'black',
                                                                            // selectors: {
                                                                            //   ':hover .ms-Button-icon': {
                                                                            //     color: 'red'
                                                                            //   },
                                                                            //   ':active .ms-Button-icon': {
                                                                            //     color: 'yellow'
                                                                            //   }
                                                                            // }
                                                                        },
                                                                        rootHovered: { backgroundColor: "white" },
                                                                        rootPressed: { backgroundColor: 'white' }
                                                                    }}
                                                                />

                                                            </span>
                                                            <span>
                                                                <li key={ind}>{file.name}</li>
                                                            </span>

                                                        </div>)}
                                                </ul>
                                            </span>
                                        </div>
                                    </fieldset>
                                </PivotItem>
                                <PivotItem headerText="Approval">
                                    <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                        <div className='conationer'>
                                            {this.state.approverFields.map((ele, ind) => {
                                                // if (ele.Title === "Distribution list") {
                                                if (ele.internalName === "ApprFldReletedPerson10Id") {
                                                    return (<>
                                                        {ele.Disable === true ? null :
                                                            <div className='fieldEditor' key={ind}>
                                                                <div>
                                                                    <label className='ApprovarLabel' htmlFor={ele.internalName}>{ele.Title}{ele.Required ? <span className='labelIcon'>*</span> : null} </label>
                                                                </div>
                                                                <div className='feildDisplay'>
                                                                    <div className='peoplepicker' >
                                                                        <PeoplePicker
                                                                            context={this.props.context}
                                                                            // titleText={value.Title}
                                                                            personSelectionLimit={50}
                                                                            groupName={""} // Leave this blank in case you want to filter from all users
                                                                            showtooltip={true}
                                                                            required={ele.Required}
                                                                            disabled={false}
                                                                            defaultSelectedUsers={this.state.approversEmail[ele.internalName]}
                                                                            onChange={this._getPeoplePickerItems.bind(this, ele.internalName)}
                                                                            showHiddenInUI={false}
                                                                            ensureUser={true}
                                                                            // allowUnvalidated={true}
                                                                            principalTypes={[PrincipalType.SharePointGroup, PrincipalType.User, PrincipalType.SecurityGroup]}
                                                                            // principalTypes={[PrincipalType.DistributionList]}
                                                                            resolveDelay={1000} />
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        }
                                                    </>

                                                    )

                                                } else {
                                                    return (<>
                                                        {ele.Disable === true ? null :
                                                            <div className='fieldEditor' key={ind}>
                                                                <div>
                                                                    <label className='ApprovarLabel' htmlFor={ele.internalName}>{ele.Title}{ele.Required ? <span className='labelIcon'>*</span> : null} </label>
                                                                </div>
                                                                <div className='feildDisplay'>
                                                                    <div className='peoplepicker' >
                                                                        <PeoplePicker
                                                                            context={this.props.context}
                                                                            // titleText={value.Title}
                                                                            personSelectionLimit={50}
                                                                            groupName={this.props.ApproversGroupName} // Leave this blank in case you want to filter from all users
                                                                            showtooltip={true}
                                                                            required={ele.Required}
                                                                            disabled={false}
                                                                            defaultSelectedUsers={this.state.approversEmail[ele.internalName]}
                                                                            // defaultSelectedUsers={["Gen_Admin"]}
                                                                            onChange={this._getPeoplePickerItems.bind(this, ele.internalName)}
                                                                            showHiddenInUI={false}
                                                                            ensureUser={true}
                                                                            // allowUnvalidated={true}
                                                                            // principalTypes={[PrincipalType.SharePointGroup]}
                                                                            principalTypes={[PrincipalType.User]}
                                                                            resolveDelay={1000} />
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        }
                                                    </>

                                                    )
                                                }


                                            })}
                                        </div>
                                    </fieldset>
                                </PivotItem>
                            </Pivot>

                            <div className='Spanbutton'>
                                <span style={{ marginRight: "1%" }}><PrimaryButton type='reset' text='Submit' onClick={this.onSubmitData} /></span>
                                <span style={{ marginRight: "1%" }}><PrimaryButton type='reset' text='Draft' onClick={this.onDraftData} /></span>

                                {/* <span><PrimaryButton type='reset' text='Exit' onClick={this.RedirecthomePage} /></span> */}
                                <span><PrimaryButton type='reset' text='Exit' onClick={this.homePage} /></span>
                            </div>
                        </form >
                    </div>
                    <Dialog
                        hidden={this.state.hideDialog}
                        onDismiss={this.toggleHideDialog}
                        minWidth={300}
                        dialogContentProps={dialogContentProps}
                    >
                        <div className={styles.dailogContent}>
                            {this.state.AlertMsg.every(ele => ele.fieldName === "") ? (<>{this.state.AlertMsg.map((ele, ind) => <p className={styles.Successmsg} key={ind}>{ele.errorMsg}</p>)}</>) :
                                <>
                                    <p >Please fill below details.</p>
                                    <ul style={{ listStyleType: "none" }}>
                                        {this.state.AlertMsg.map((ele, index) => <li key={index}>-{ele.fieldName}</li>)}
                                    </ul>
                                </>}
                        </div>
                        <DialogFooter>
                            <PrimaryButton text="Ok" onClick={this.onCancel} />
                        </DialogFooter>
                    </Dialog>
                </section >
            );
        }
        else {
            return (
                <section className={`${styles.sonyEdib} ${hasTeamsContext ? styles.teams : ''}`} >
                    <div>
                        <h1 style={{ textAlign: "center" }}>Sorry! you are not authorized.</h1>
                    </div>
                    <div className='Spanbutton'>
                        <span><PrimaryButton type='reset' text='Exit' onClick={this.homePage} /></span>
                        {/* <span><PrimaryButton type='reset' text='Exit' onClick={this.RedirecthomePage} /></span> */}
                    </div>
                </section>
            )
        }
    }
}
