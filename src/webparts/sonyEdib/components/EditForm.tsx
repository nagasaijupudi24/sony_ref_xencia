import * as React from 'react';
import styles from './SonyEdib.module.scss';
import { ISonyEdibProps } from './ISonyEdibProps';
import { ActivityItem, Dialog, DialogFooter, DialogType, Icon, IconButton, IPersonaProps, Link, mergeStyleSets, Persona, PersonaSize, Pivot, PivotItem, PrimaryButton } from '@fluentui/react';
import spService from './Serivce/spService';
import './custom.css'
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
// import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/files";
import "@pnp/sp/profiles";
import "@pnp/sp/batching";
import { IAlertMsg, IAppREmails, IBoardAppr, IFileDetails } from './Createform';
import { UserProfileProperties } from './ViewForm';
export interface IFieldCollection {
    Title: string,
    DataType: string,
    Required: boolean,
    Tab: string,
    internalName: string,
    Option?: string[];
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
export interface IAuditTrail {

    Actioner: string,
    ActionTaken: string,
    Role: string,
    ActionTakenOn: string,
    Comments?: string
}
// export interface IFileInfo {
//     name?: string,
//     content?: string,
//     index?: string | number,
//     fileUrl?: string,
//     ServerRelativeUrl?: string,
//     isExists?: boolean,
//     Modified?: string,
//     isSelected?: boolean
// }
export interface ISonyEdibState {
    fieldCollection: IFieldCollection[],
    Data: any,
    getItems: any,
    auditLog: IAuditTrail[]
    approverFields: IApproverField[];
    approversEmail: IAppREmails;
    configApprover: string[];
    checkboxItems: string[]
    attachfiles: IFileDetails[];
    getattchements: IFileDetails[];
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
    imageinitials: string,
    RequsterEmail: string;
    hideDialog: boolean;
    AlertMsg: IAlertMsg[];
    fileUrl: string[],
    isDIBIssuer: boolean,
    isSpecifySelected?: any;
    SpecifyOwnvalues?: any;


}
const getIdFromUrl = (): string => {
    const params = new URLSearchParams(window.location.search);
    const Id = params.get('itemId');
    // console.log(Id);
    return Id;
};
const activityItemExamples: { key: string | number; activityDescription?: JSX.Element[]; activityIcon?: JSX.Element; comments?: JSX.Element[]; timeStamp?: string; }[] = []
const classNames = mergeStyleSets({
    exampleRoot: {
        marginTop: '20px',
    },
    nameText: {
        fontWeight: 'bold',
    },
    space: {
        padding: "0 5px"
    }
});
export default class EditForm extends React.Component<ISonyEdibProps, ISonyEdibState, {}> {
    private _spService: spService = null;
    // private _checkBoxItems: any[] = [];
    private _sp;
    // private fileInfos: any = [];
    private _userName: string;
    private _UserRole: string;
    // private _userName: string;
    private _UserEmail: string;
    private _userpictureUrl: string;
    private _userfirstName: string;
    private _userlastName: string;
    private _itemId: number = Number(getIdFromUrl());
    private _fldrName: string = "";
    private boardApprovers: IBoardAppr[] = [];
    constructor(props: ISonyEdibProps) {
        super(props);
        this.state = {
            fieldCollection: [],
            Data: {},
            getItems: {},
            auditLog: [],
            approverFields: [],
            approversEmail: {},
            configApprover: [],
            attachfiles: [],
            getattchements: [],
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
            RequsterEmail: "",
            hideDialog: true,
            AlertMsg: [],
            fileUrl: [],
            isDIBIssuer: false,
            isSpecifySelected: {},
            SpecifyOwnvalues: {}
        }
        this._sp = spfi().using(SPFx(this.props.context))
        this._spService = new spService(this.props.context);

        Promise.all([
            this.getSiteGroups(),
            this.GetUserProperties(),
            this.getApproverEmail(),
            this.getItems(),
            this.approvarFields()
        ]).catch(err => console.error(err));
        // this.getSiteGroups().then(res => res).catch(err => err);
        // this.GetUserProperties().then(res => res).catch(err => err);
        // this.getApproverEmail().then(res => res).catch(err => err);
        // // this.approvarFields().then(res => res).catch(err => err);
        // // this.formInput();
        // this.getItems().then(res => res).catch(err => err);

    }

    // get user details
    private GetUserProperties = async (): Promise<void> => {
        await this._sp.profiles.myProperties().then((result: { UserProfileProperties: UserProfileProperties; DisplayName: string; }) => {
            const props = result.UserProfileProperties;
            this._userName = result.DisplayName;
            // var properties = props.UserProfileProperties;
            for (let i = 0; i < props.length; i++) {
                const allProperties = props[i];
                if (allProperties.Key === "PictureURL") {
                    this._userpictureUrl = allProperties.Value;
                }
                else if (allProperties.Key === "Title") {
                    this._UserRole = allProperties.Value;
                }
                else if (allProperties.Key === "FirstName") {
                    const frstname = allProperties.Value;
                    this._userfirstName = frstname.substring(0, 1);
                }
                else if (allProperties.Key === "LastName") {
                    const lastname = allProperties.Value;
                    this._userlastName = lastname.substring(0, 1);
                }
                else if (allProperties.Key === "WorkEmail") {
                    this._UserEmail = allProperties.Value;
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
                <div className='userName'>Role: {this.state.reqRole}</div>
                <span className="" >{this.state.userDepartment}</span>
            </div>
        );
    }
    public getItems = async (): Promise<void> => {
        if (this._itemId !== null) {


            // await this._sp.web.lists.getByTitle(this.props.listName).items.getById(this._itemId).select("Author/EMail").expand("Author")().then(res => {
            //     // console.log("requster", res.Author.EMail)
            //     this.setState({ RequsterEmail: res.Author.EMail })
            // })
            await this._spService.getListItemsById(this.props.listName, this._itemId).then(res => {
                // console.log(res)
                // console.log(res.Author.EMail)
                this._fldrName = res.Title
                this.setState({
                    auditLog: JSON.parse(res.AuditTrail),
                    Data: res,
                    RequsterEmail: res.Author.EMail

                });

                // console.log(res)
                // this.setState({ Data: res, auditLog: JSON.parse(res.AuditTrail) });
                // console.log("res", this._fldrName);
                const audit: IAuditTrail[] = JSON.parse(res.AuditTrail);
                this.boardApprovers = JSON.parse(res.BoardApprovers);
                // audit.push( JSON.parse(res.AuditTrail))
                if (audit !== null) {
                    activityItemExamples.length = 0
                    audit.forEach((ele) => {
                        if (ele.ActionTaken === "Submitted") {
                            activityItemExamples.unshift({
                                key: Math.floor(Math.random() * 10000),
                                activityDescription: [
                                    <span className={classNames.nameText} key={1}>{ele.Actioner}</span>,
                                    <span className={classNames.space} key={2}>{ele.ActionTaken}</span>,
                                    <div key={ele.Role}><span key={3}>Role: {ele.Role}</span></div>,

                                ],
                                activityIcon: <Icon iconName={'TextDocumentShared'} />,
                                comments: [
                                    <span key={1}> {ele.Comments}</span>
                                ],
                                timeStamp: ele.ActionTakenOn
                            });
                        }
                        else if (ele.ActionTaken === "Resubmitted") {
                            activityItemExamples.unshift({
                                key: Math.floor(Math.random() * 10000),
                                activityDescription: [
                                    <span className={classNames.nameText} key={1}>{ele.Actioner}</span>,
                                    <span className={classNames.space} key={2}>{ele.ActionTaken}</span>,
                                    <div key={ele.Role}><span key={3}>Role: {ele.Role}</span></div>,

                                ],
                                activityIcon: <Icon iconName={'Message'} />,
                                comments: [
                                    <span key={1}> {ele.Comments}</span>
                                ],
                                timeStamp: ele.ActionTakenOn
                            });
                        } else if (ele.ActionTaken.match(/Rejected/g)) {
                            activityItemExamples.unshift({
                                key: Math.floor(Math.random() * 10000),
                                activityDescription: [
                                    <span className={classNames.nameText} key={1}>{ele.Actioner}</span>,
                                    <span className={classNames.space} key={2}>{ele.ActionTaken}</span>,
                                    <div key={3}><span key={3}>Role: {ele.Role}</span></div>,

                                ],
                                activityIcon: <Icon iconName={'ErrorBadge'} />,
                                comments: [
                                    <span key={1}> {ele.Comments}</span>
                                ],
                                timeStamp: ele.ActionTakenOn
                            });
                        }
                        else if (ele.ActionTaken.match(/Approved/g)) {
                            activityItemExamples.unshift({
                                key: Math.floor(Math.random() * 10000),
                                activityDescription: [
                                    <span className={classNames.nameText} key={1}>{ele.Actioner}</span>,
                                    <span className={classNames.space} key={2}>{ele.ActionTaken}</span>,
                                    <div key={ele.Role}><span key={3}>Role: {ele.Role}</span></div>,

                                ],
                                activityIcon: <Icon iconName={'Completed'} />,
                                comments: [
                                    <span key={1}> {ele.Comments}</span>
                                ],
                                timeStamp: ele.ActionTakenOn
                            });
                        }
                        else {
                            activityItemExamples.unshift({
                                key: Math.floor(Math.random() * 10000),
                                activityDescription: [
                                    <span className={classNames.nameText} key={1}>{ele.Actioner}</span>,
                                    <span className={classNames.space} key={2}>{ele.ActionTaken}</span>,
                                    <div key={ele.Role}><span key={3}>Role: {ele.Role}</span></div>,

                                ],
                                activityIcon: <Icon iconName={'TextDocumentShared'} />,
                                comments: [
                                    <span key={1}> {ele.Comments}</span>
                                ],
                                timeStamp: ele.ActionTakenOn
                            });
                        }
                    })

                }
                this.getFilesInsideFolder(this._fldrName).then(res => res).catch(err => err);

            });
            this.formInput(this.state.Data).then(res => res).catch(err => err);
        }
    }

    public async formInput(result: any): Promise<void> {
        const collectionfileds = this.props.collectionData;
        const fieldIfo: { Title: string; DataType: string; Required: boolean; Tab: string; Option: string[]; internalName: string; FillInChoice?: boolean }[] = []
        this._spService.getfieldDetails(this.props.listName).then((res) => {
            res.filter(function (newData: { text: string; Title: string; dataType: string; option: string[]; internalName: string; FillInChoice?: boolean }) {
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
                                // "DefaultValue": newData.DefaultValue,
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
                                FillInChoice: newData.FillInChoice
                            })

                        }
                    }
                });
            });

            // console.log("fieldIfo", fieldIfo);
            fieldIfo.filter(x => {

                if (typeof this.state.Data[x.internalName] !== "undefined" && this.state.Data[x.internalName] !== null) {
                    if (x.FillInChoice) {


                        if (x.DataType === "MultiChoice") {
                            const temp = [...new Set([...x.Option, ...this.state.Data[x.internalName]])];
                            temp.forEach(newObj => {
                                if (!x.Option.includes(newObj)) {
                                    const slectedOption: any[] = this.state.Data[x.internalName];
                                    const index = slectedOption.indexOf(newObj)
                                    // console.log(slectedOption, index)
                                    if (index > -1) {
                                        slectedOption.splice(index, 1);
                                    }
                                    // console.log(slectedOption, index)
                                    this.setState(prevState => ({
                                        SpecifyOwnvalues: { ...prevState.SpecifyOwnvalues, [x.internalName]: newObj },
                                        isSpecifySelected: { ...prevState.isSpecifySelected, [x.internalName]: true },
                                        Data: { ...prevState.Data, [x.internalName]: slectedOption },
                                        getItems: { ...prevState.getItems, [x.internalName]: slectedOption.length > 0 ? slectedOption : [] }
                                    }));
                                }
                            });
                        } else {
                            if (this.state.Data[x.internalName] !== undefined && this.state.Data[x.internalName] !== null && this.state.Data[x.internalName] !== "") {
                                if (!x.Option.includes(this.state.Data[x.internalName])) {
                                    this.setState(prevState => ({
                                        SpecifyOwnvalues: { ...prevState.SpecifyOwnvalues, [x.internalName]: this.state.Data[x.internalName] },
                                        Data: { ...prevState.Data, [x.internalName]: "Specify your own value :" }
                                    }));

                                }

                            }
                        }
                    }
                }
                else {
                    if (x.FillInChoice && x.DataType === "MultiChoice") {
                        this.setState(prevState => ({
                            Data: { ...prevState.Data, [x.internalName]: [] },
                            getItems: { ...prevState.getItems, [x.internalName]: [] }
                        }));

                    }
                }
            })
            this.setState({ fieldCollection: fieldIfo });
        }).catch(error => {
            console.log("Something went wrong! please contact admin for more information.", error);
        });

    }
    // 

    public approvarFields = async (): Promise<void> => {

        await this._spService.apprConfigu(this.props.listName).then(res => this.setState({ approverFields: res })).catch(err => err);
    }

    // get Approvers
    public getApproverEmail = async (): Promise<void> => {
        const listTitle: string = this.props.listName;
        if (this._itemId !== null) {
            const result = await this._sp.web.lists.getByTitle(listTitle).items.getById(this._itemId)
                .select(`ApprFldReletedPerson1/EMail,ApprFldReletedPerson2/EMail,ApprFldReletedPerson3/EMail,ApprFldReletedPerson4/EMail,ApprFldReletedPerson5/EMail,ApprFldReletedPerson6/EMail,ApprFldReletedPerson7/EMail,ApprFldReletedPerson8/EMail,ApprFldReletedPerson9/EMail,ApprFldReletedPerson10/Title`)
                .expand(`ApprFldReletedPerson1,ApprFldReletedPerson2,ApprFldReletedPerson3,ApprFldReletedPerson4,ApprFldReletedPerson5,ApprFldReletedPerson6,ApprFldReletedPerson7,ApprFldReletedPerson8,ApprFldReletedPerson9,ApprFldReletedPerson10`)();
            for (let i = 1; i <= 10; i++) {
                const fieldTitle = `ApprFldReletedPerson${i}`;
                const apprTemp: string[] = [];
                const apprEmail = result[fieldTitle];
                if (i === 10) {
                    if (typeof apprEmail !== "undefined") {
                        for (let j = 0; j < apprEmail.length; j++) {
                            apprTemp.push(apprEmail[j].Title);

                        }
                    }
                }
                else {
                    if (typeof apprEmail !== "undefined") {
                        for (let j = 0; j < apprEmail.length; j++) {
                            apprTemp.push(apprEmail[j].EMail);

                        }
                    }

                }
                //   });
                this.setState(prevState => ({
                    approversEmail: {
                        ...prevState.approversEmail,
                        [`${fieldTitle}Id`]: apprTemp
                    }
                }));
            }
        }
        // console.log(this.state.approversEmail)
    }

    public eventHandler = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
        this.setState(prevState => ({ getItems: { ...prevState.getItems, [name]: value } }));

        // console.log(this.state.getItems);
    }
    // yes or no columns
    public eventYesOrNoHandler = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        // console.log(value)
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
        this.setState(prevState => ({ getItems: { ...prevState.getItems, [name]: value } }));
        // console.log(this.state.Data, this.state.getItems)

        // console.log(this.state.getItems);
    }

    public handleChxBoxChange = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        let dataArray: string[] = this.state.Data[name]
        let isExit = false;
        if (dataArray === null) {
            dataArray = []
            dataArray.push(value);
            this.setState(prevState => ({ Data: { ...prevState.Data, [name]: dataArray } }));
            isExit = true;

        }
        else if (!dataArray.includes(value) && isExit === false) {
            dataArray.push(value);
            this.setState(prevState => ({ Data: { ...prevState.Data, [name]: dataArray } }));
        }

        else {

            const index = dataArray.indexOf(value);
            if (index > -1) {
                dataArray.splice(index, 1);
            }
        }
        // BoardAppr Json
        let isobjExit = false;
        if (name === "Board") {
            // if (name.match(/Board/g)) {

            if (this.boardApprovers === null) {
                this.boardApprovers = []
                this.boardApprovers.push({
                    Board: value,
                    isUpdated: false,
                    isCompleted: false,
                    CCT: { Status: "", isChanged: false, Comments: "" },
                    PWB: { Status: "", isChanged: false,Comments: ""  }
                })
                isobjExit = true

            }

            else if (!(this.boardApprovers.some(object => object.Board === value)) && isobjExit === false) {
                // debugger;
                this.boardApprovers.push({
                    Board: value,
                    isUpdated: false,
                    isCompleted: false,
                    CCT: { Status: "", isChanged: false, Comments: "" },
                    PWB: { Status: "", isChanged: false }
                })
            }
            else {
                const filteredArray = this.boardApprovers.filter(obj => obj.Board === value);
                const index = this.boardApprovers.indexOf(filteredArray[0]);
                this.boardApprovers.splice(index, 1);
            }
        }

        this.setState(prevState => ({ getItems: { ...prevState.getItems, [name]: dataArray, BoardApprovers: JSON.stringify(this.boardApprovers) } }));
    }

    public handleTextareaChange = (event: React.ChangeEvent<HTMLTextAreaElement>, ind?: number): void => {
        const { name, value } = event.target;
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
        this.setState(prevState => ({ getItems: { ...prevState.getItems, [name]: value } }));
    }

    private getSiteGroups = async (): Promise<void> => {

        const groups = await this._sp.web.currentUser.groups();
        const gropsTitle = groups.map(x => x.Title);
        if (gropsTitle.includes(this.props.DibissuersGroup)) {
            this.setState({
                isDIBIssuer: true
            })
        }
        // console.log(groups);
    }

    private _getPeoplePickerItems(nm: string, items: IPeoplePickerItems[]): void {
        const apprIds: number[] = []
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
        this.setState(prevState => ({ getItems: { ...prevState.getItems, [nm]: apprIds } }));
        this.setState(prevState => ({ Data: { ...prevState.Data, [nm]: apprIds } }));
        this.setState(prevState => ({ approversEmail: { ...prevState.approversEmail, [nm]: apprEmails } }));
        // console.log(this.state.getItems);
    }
    public SpecifyhandleChxBoxChange = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        this.setState((prevState) => {
            const updatedIsSpecifySelected = { ...prevState.isSpecifySelected };
            if (!Object.keys(updatedIsSpecifySelected).includes(name)) {
                updatedIsSpecifySelected[name] = Boolean(value);
            } else {
                updatedIsSpecifySelected[name] = !updatedIsSpecifySelected[name];
            }
            return { isSpecifySelected: updatedIsSpecifySelected };
        });

        // console.log(this.state.isSpecifySelected);

    }
    public SpecifyeventHandler = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        this.setState(prevState => ({ SpecifyOwnvalues: { ...prevState.SpecifyOwnvalues, [name]: value } }));
        // console.log(this.state.SpecifyOwnvalues[name]);
    }

    public setSpecifyValueforCheckBox = (): void => {

        const { fieldCollection, isSpecifySelected, getItems, SpecifyOwnvalues, Data } = this.state;

        fieldCollection.forEach(_x => {
            // console.log(this.state.isSpecifySelected[_x.internalName], _x.internalName)
            if (_x.FillInChoice && getItems[_x.internalName] !== undefined) {
                if (getItems[_x.internalName] === "Specify your own value :") {
                    this.setState(prevState => ({
                        getItems: {
                            ...prevState.getItems,
                            [_x.internalName]: SpecifyOwnvalues[_x.internalName] || null
                        }
                    }));
                }
            }
            if (_x.FillInChoice && Data[_x.internalName] !== undefined && _x.DataType === "MultiChoice") {
                if (!Data[_x.internalName].includes(SpecifyOwnvalues[_x.internalName])) {
                    this.setState(prevState => ({
                        getItems: {
                            ...prevState.getItems,
                            [_x.internalName]: isSpecifySelected[_x.internalName] ? [...Data[_x.internalName], SpecifyOwnvalues[_x.internalName] || []] : Data[_x.internalName]
                        },
                        Data: {
                            ...prevState.Data,
                            [_x.internalName]: isSpecifySelected[_x.internalName] ? [...Data[_x.internalName], SpecifyOwnvalues[_x.internalName] || []] : Data[_x.internalName]
                        },

                    }));
                    // console.log(this.state.Data[_x.internalName], this.state.Data[_x.internalName]);
                }
            }
        });
    };
    //onSubmitData onUpadate
    public onSubmitData = async (e: { preventDefault: () => void; }): Promise<void> => {
        await this.setSpecifyValueforCheckBox();
        e.preventDefault();
        const tenantUrl = window.location.protocol + "//" + window.location.host;
        const siteUrl = this.props.siteUrl.replace(tenantUrl, "") + `/${this.props.libraryName}/` + this._fldrName;
        const emptyFileds: { fieldName?: string; errorMsg: string; }[] = []
        // vaalidations
        const fltrArry = this.state.fieldCollection.filter(ele => ele.Required);
        const fltrArryApprover = this.state.approverFields.filter(ele => ele.Required && !ele.Disable);
        // const emptyFileds: { fieldName?: string; errorMsg: string; }[] = []
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
        const isValidApprFields = fltrArryApprover.map((x) => {
            let isValid = false;
            let isEmpty = false;

            if (Object.keys(this.state.Data).includes(x.internalName)) {
                if (
                    typeof this.state.Data[x.internalName] !== "undefined" &&
                    this.state.Data[x.internalName] !== null &&
                    this.state.Data[x.internalName].length !== 0
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
            let isUpdate: boolean = false;
            const auditlog = this.state.auditLog;
            const actioner = this._userName;
            const role = this._UserRole;
            const comments = "No Comments";
            let status: string = "";
            let statusNo: string = "";
            if (this.state.Data.StatusNo === "5") {
                status = "Submitted";
                statusNo = "20"
            }
            else if (this.state.Data.StatusNo === "15" || this.state.Data.StatusNo === "10000") {
                status = "Resubmitted";
                statusNo = "10"
            }
            const obj = {
                "Actioner": actioner,
                "ActionTaken": status,
                "Role": role,
                "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                "Comments": comments
            };
            auditlog.push(obj);
            const updateObj = {
                StatusNo: statusNo,
                Status: status,
                StartProcessing: true,
                AuditTrail: JSON.stringify(auditlog)
            }
            // this.setState(prevState => ({ Data: { ...prevState.Data,AuditTrail : JSON.stringify(auditlog) } }));
            await this._spService.updateListItem(this.props.listName, this.state.getItems, this._itemId).then(async res => {
                await this.deleteFilesFormSPLbry().then(res => res).catch(err => err);
                await this._spService.updateListItem(this.props.listName, updateObj, this._itemId).then(res => {
                    emptyFileds.push({
                        fieldName: "",
                        errorMsg: "Request has been submitted successfully"
                    });
                    isUpdate = true;
                    return res;
                }).catch(err => err);
                // for approve checking after update the audittrail
                if (isUpdate === true) {
                    await this.apprCheckingUpdate(this._itemId).then(res => res).catch(err => err);
                }
                if (this.state.attachfiles.length > 0) {
                    await this._spService.uploadAttachemnt(this.state.attachfiles, siteUrl).then(res => res).catch(err => err);
                }
            }).catch(err => err);
        }
        this.setState({ AlertMsg: emptyFileds, hideDialog: false });
    }
    // Approver fields checking------
    public apprCheckingUpdate = async (Id: number): Promise<void> => {
        const apprCheckingIsEmtyOrNot = await this._spService.apprFieldCheck(this.state.Data, this.state.approverFields);
        if (apprCheckingIsEmtyOrNot.length > 0) {
            const varStatus = apprCheckingIsEmtyOrNot[0].status;
            const varActionTaken = apprCheckingIsEmtyOrNot[0].ActionTaken;
            const updateObj = {
                Status: varActionTaken,
                StatusNo: varStatus
            }
            await this._spService.updateListItem(this.props.listName, updateObj, Id).then(res => res).catch(err => (err));
        }
    }

    public onDraft = async (e: { preventDefault: () => void; }): Promise<void> => {
        await this.setSpecifyValueforCheckBox();
        e.preventDefault();
        const tenantUrl = window.location.protocol + "//" + window.location.host;
        const siteUrl = this.props.siteUrl.replace(tenantUrl, "") + `/${this.props.libraryName}/` + this._fldrName;
        const emptyFileds: { fieldName?: string; errorMsg: string; }[] = []
        // vaalidations
        // const fltrArry = this.state.fieldCollection.filter(ele => ele.Required);
        // const fltrArryApprover = this.state.approverFields.filter(ele => ele.Required && !ele.Disable);
        // console.log(fltrArryApprover);
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
        // let AttachemntsValiadtion: boolean = false;
        // if (this.props.IsAttachmentsRequired === true) {
        //     if (this.state.attachfiles.length > 0) {
        //         AttachemntsValiadtion = true
        //     }
        //     else {
        //         emptyFileds.push({
        //             fieldName: "Upload Attachments",
        //             errorMsg: "Please fill in the below field"
        //         });


        //     }
        // }
        // else {
        //     AttachemntsValiadtion = true
        // }

        // const isValidation = isValidFields.every((isValid) => isValid);
        // const isApprFieldValidation = isValidApprFields.every((isValid) => isValid);
        // if (isValidation && isApprFieldValidation && AttachemntsValiadtion) {
        // let isUpdate: boolean = false;
        const auditlog = this.state.auditLog;
        const actioner = this._userName;
        const role = this._UserRole;
        const comments = "No Comments";
        let statusNo: string = "";
        if (this.state.Data.StatusNo === "10000" || this.state.Data.StatusNo === "15") {
            statusNo = "15"
        }
        else {
            statusNo = "5"
        }
        const obj = {
            "Actioner": actioner,
            "ActionTaken": "Drafted",
            "Role": role,
            "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
            "Comments": comments
        };
        auditlog.push(obj);
        const updateObj = {
            StatusNo: statusNo,
            Status: "Drafted",
            StartProcessing: true,
            AuditTrail: JSON.stringify(auditlog)
        }
        await this._spService.updateListItem(this.props.listName, Object.assign(this.state.getItems, updateObj), this._itemId).then(async res => {
            emptyFileds.push({
                fieldName: "",
                errorMsg: "Request has been drafted successfully"
            })
            await this.deleteFilesFormSPLbry().then(res => res).catch(err => err);
            if (this.state.attachfiles.length > 0) {
                await this._spService.uploadAttachemnt(this.state.attachfiles, siteUrl).then(res => res).catch(err => err);
            }
        }).catch(err => err);
        this.setState({ AlertMsg: emptyFileds, hideDialog: false });
    }
    // add Attachments
    private addAttacment = async (): Promise<void> => {
        const fileInfo: { name: string; content: File; index: number; fileUrl: string; ServerRelativeUrl: string; isExists: boolean; Modified: string; isSelected: boolean; }[] = [];
        const fileInput = document.getElementById('Docfiles') as HTMLInputElement;
        const fileCount = fileInput.files.length;
        for (let i = 0; i < fileCount; i++) {
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
                            //   fileid: this.state.fileId
                        });
                    }
                    this.setState({ attachfiles: [...this.state.attachfiles, ...fileInfo] });
                    this.setState({ getattchements: [...this.state.getattchements, ...fileInfo] });
                    // this.fileInfos = this.state.attachfiles;getattchements
                };
            })(file);
            reader.readAsArrayBuffer(file);
        }
    }
    // get files inside folder
    public getFilesInsideFolder = async (folderName: string): Promise<void> => {
        const files = await this._sp.web.getFolderByServerRelativePath(this.props.context.pageContext.web.serverRelativeUrl + `/${this.props.libraryName}/` + folderName).files();
        // console.log("files", files)
        const tenantUrl = window.location.protocol + "//" + window.location.host;
        const tempFiles: IFileDetails[] = []
        files.forEach(values => {
            const filesObj = {
                "name": values.Name,
                "content": null as File,
                "index": 0,
                "fileUrl": tenantUrl + values.ServerRelativeUrl,
                "ServerRelativeUrl": "",
                "isExists": true,
                "Modified": "",
                "isSelected": false
            }
            tempFiles.push(filesObj);
        });
        this.setState({ getattchements: tempFiles });
    }
    // Remove Attachemnts
    public onRemoveAttachments = (file: IFileDetails): void => {
        const { attachfiles, fileUrl, getattchements } = this.state;
        const temp: string[] = fileUrl;
        if (file.isExists) {
            temp.push(file.name);
        }
        this.setState({ fileUrl: temp });
        const fltrArry = attachfiles.filter(obj => obj.name === file.name);
        const index = attachfiles.indexOf(fltrArry[0]);
        if (index > -1) {
            attachfiles.splice(index, 1);
        }
        //////
        const fltrArry1 = getattchements.filter(obj => obj.name === file.name);
        const index1 = getattchements.indexOf(fltrArry1[0]);
        if (index1 > -1) {
            getattchements.splice(index1, 1);
        }
    }
    // files delete form library
    public deleteFilesFormSPLbry = async (): Promise<void> => {
        // debugger;
        const { fileUrl } = this.state;
        for (let i = 0; i < fileUrl.length; i++) {
            await this._sp.web.getFolderByServerRelativePath(this.props.context.pageContext.web.serverRelativeUrl + `/${this.props.libraryName}/` + this._fldrName).files.getByUrl(fileUrl[i]).delete();
        }
    }
    public toggleHideDialog = (): void => {
        this.setState({ hideDialog: false })
    }

    public onCancel = (): void => {
        this.state.AlertMsg.map(x => {
            if (x.fieldName === "") {
                this.homePage();
            }
            else {
                this.setState({ hideDialog: true });
            }
        })
    }
    public homePage = (): void => {
        const pageURL: string = this.props.context.pageContext.web.absoluteUrl;
        // console.log(pageURL);
        if(window.history.length>1){
            window.history.go(-1)
        }
        else{
            window.location.href = pageURL;
        }
        this.setState({ hideDialog: true });
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
                    {/* <form onSubmit={this.onSubmit}> */}
                    <div className={styles.sonydibcontainer}>
                        <div className={styles.frmtitle}>Edit Form</div>
                        <Pivot aria-label="Large Link Size Pivot Example" >
                            <PivotItem linkText="Requestor Details">
                                <fieldset style={{ height: 60, "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className="ms-Grid-row personarow" style={{ paddingTop: 5 }}>
                                        <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 personaCol1">
                                            <div className="ms-Grid-col ms-sm12 ms-lg12">
                                                <Persona
                                                    imageUrl={this.state.pictureUrl}
                                                    size={PersonaSize.size48}
                                                    imageInitials={this.state.imageinitials}
                                                    onRenderPrimaryText={this._onRenderPrimaryText}
                                                    onRenderSecondaryText={this._onRenderSecondaryText}
                                                />
                                            </div>
                                        </div>
                                        <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 personaCol2">
                                            <div className="ms-Grid-row presonaCol2Row" >
                                                <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg2 presonaCol2col1">
                                                    <span className="hdrTtle">Status: </span>
                                                </div>
                                                <div className="ms-Grid-col ms-sm8 ms-md9 ms-lg10 presonaCol2col2">
                                                    <span>{this.state.Data.Status}</span>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </fieldset>
                            </PivotItem>
                        </Pivot>
                        {/* < form > */}
                        <Pivot aria-label="Large Link Size Pivot Example" >
                            <PivotItem headerText="DIB content">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
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
                                                                        <div className='fieldEditor' key={index} >
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
                                                                                        disabled={(value.internalName === "DIB_ID" || value.internalName === "Title")}
                                                                                        readOnly={(value.internalName === "DIB_ID" || value.internalName === "Title")}
                                                                                        className='inputText'
                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                    />
                                                                                </div>
                                                                            </div>
                                                                        </div>
                                                                    )
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
                                                                                        <div key={index} className={styles.container}>
                                                                                            {/* <input type="radio" name={value.internalName} value={ele} onChange={(event) => this.eventHandler(event, index)} required={value.Required} />{ele} */}
                                                                                            <input type="checkbox" className="checkbox" name={value.internalName} value={ele}
                                                                                                onChange={(event) => this.handleChxBoxChange(event, index)}
                                                                                                checked={this.state.Data[value.internalName] !== null ? Object.values(this.state.Data[value.internalName]).includes(ele) : false}
                                                                                                // checked={this.state.Data[value.internalName] !== null || typeof this.state.Data[value.internalName] !== "undefined" ? Object.values(this.state.Data[value.internalName]).includes(ele) : false}
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
                                                                                        disabled={!value.FillInChoice}
                                                                                        // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                        value={this.state.isSpecifySelected[value.internalName] ? this.state.SpecifyOwnvalues[value.internalName] || "" : ""}
                                                                                        style={{ outline: "none" }}
                                                                                        className={styles.eDIBSpecifyOwnvalueInput}
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
                                                                                {/* <label className="label" htmlFor={value.internalName}>{value.Title}</label> */}
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className='radioCheckbox'>
                                                                                {value.Option.map((ele: string) => {
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
                                                                                {value.FillInChoice ?
                                                                                    <input
                                                                                        type="text"
                                                                                        name={value.internalName}
                                                                                        // id={value.internalName}
                                                                                        disabled={!value.FillInChoice}
                                                                                        // value={this.state.Data[value.internalName]==="Specify your own value :"?this.state.SpecifyOwnvalues[value.internalName] || "":""}
                                                                                        value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                        style={{ outline: "none" }}
                                                                                        className={styles.eDIBSpecifyOwnvalueInput}
                                                                                        onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                    />

                                                                                    : null}
                                                                            </div>
                                                                        </div>
                                                                    )
                                                                // ________________________________________________________________________________________________________________________________________________________________________________________________________________________________
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
                                                                                                onChange={(event) => this.eventYesOrNoHandler(event, index)}
                                                                                                // checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                checked={(this.state.Data[value.internalName] !== null && this.state.Data[value.internalName].toString() === ele) ? true : false}
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
                                                                                > {value.Option.includes(this.state.Data[value.internalName]) ?
                                                                                    <>
                                                                                        {value.Option.map((ele: string) => {
                                                                                            return (<option key={ele}>{ele}</option>)
                                                                                        })}
                                                                                    </>
                                                                                    : <>
                                                                                        <option />
                                                                                        {value.Option.map((ele: string) => {
                                                                                            return (<option key={ele}>{ele}</option>)
                                                                                        })}
                                                                                    </>}

                                                                                </select>
                                                                                {value.FillInChoice ? <input
                                                                                    type="text"
                                                                                    name={value.internalName}
                                                                                    // id={value.internalName}
                                                                                    disabled={!value.FillInChoice}
                                                                                    // value={this.state.SpecifyOwnvalues[value.internalName] || ""}
                                                                                    value={this.state.Data[value.internalName] === "Specify your own value :" ? this.state.SpecifyOwnvalues[value.internalName] || "" : ""}
                                                                                    style={{ outline: "none" }}
                                                                                    className={styles.eDIBSpecifyOwnvalueInputdrp}
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
                                                                                        value={this.state.Data[value.internalName || ""]}
                                                                                        className='inputText'
                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                    />
                                                                                </div>
                                                                            </div>
                                                                        </div>
                                                                    )
                                                                case "Number":
                                                                    return (
                                                                        <div className='fieldEditor' key={index} >
                                                                            <div>
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
                                                                                        className='inputText'
                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                    />
                                                                                </div>
                                                                            </div>
                                                                        </div>

                                                                    )
                                                                case "UserMulti"
                                                                    :
                                                                    return (
                                                                        <div className='fieldEditor' key={index}>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className='feildDisplay'>

                                                                                <div className='peoplepicker'>
                                                                                    <PeoplePicker
                                                                                        context={this.props.context}
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
                                </fieldset>
                            </PivotItem>
                            <PivotItem headerText="Related Check Items">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className='ms-Grid'>
                                        <div className='ms-Grid-row'>
                                            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
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
                                                                                        className='inputText'
                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                    />
                                                                                </div>
                                                                            </div>
                                                                        </div>
                                                                    )
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
                                                                                                // checked={Object.values(this.state.Data[value.internalName]).includes(ele)}
                                                                                                checked={this.state.Data[value.internalName] !== null || typeof this.state.Data[value.internalName] !== "undefined" ? Object.values(this.state.Data[value.internalName]).includes(ele) : false}
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
                                                                                        value={this.state.isSpecifySelected[value.internalName] ? this.state.SpecifyOwnvalues[value.internalName] || "" : ""}
                                                                                        style={{ outline: "none" }}
                                                                                        className={styles.eDIBSpecifyOwnvalueInput}
                                                                                        onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                    />
                                                                                </>
                                                                                    : null}
                                                                            </div>
                                                                        </div>
                                                                    )
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
                                                                                                checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                required={value.Required} />{ele}
                                                                                            <div className={styles.radiocheckmark}>
                                                                                                <span className={styles.radioinsidecircle} />
                                                                                            </div>
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
                                                                                        value={this.state.Data[value.internalName] === "Specify your own value :" ? this.state.SpecifyOwnvalues[value.internalName] || "" : ""}
                                                                                        style={{ outline: "none" }}
                                                                                        className={styles.eDIBSpecifyOwnvalueInput}
                                                                                        onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                    />

                                                                                    : null}
                                                                            </div>
                                                                        </div>
                                                                    )
                                                                // ___________________________________________________________________________________________________________________________________________________________________________________________________
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
                                                                                                onChange={(event) => this.eventYesOrNoHandler(event, index)}
                                                                                                checked={(this.state.Data[value.internalName] !== null && this.state.Data[value.internalName].toString() === ele) ? true : false}
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
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div style={{ width: "50%" }}>
                                                                                <select className='dropdwon'
                                                                                    id={value.internalName}
                                                                                    name={value.internalName}
                                                                                    value={this.state.Data[value.internalName] || ""}
                                                                                    onChange={(event) => this.eventHandler(event, index)}
                                                                                >
                                                                                    {value.Option.includes(this.state.Data[value.internalName]) ?
                                                                                        <>
                                                                                            {value.Option.map((ele: string) => {
                                                                                                return (<option key={ele}>{ele}</option>)
                                                                                            })}
                                                                                        </>
                                                                                        : <>
                                                                                            <option />
                                                                                            {value.Option.map((ele: string) => {
                                                                                                return (<option key={ele}>{ele}</option>)
                                                                                            })}
                                                                                        </>}
                                                                                </select>
                                                                                {value.FillInChoice ? <input
                                                                                    type="text"
                                                                                    name={value.internalName}
                                                                                    // id={value.internalName}
                                                                                    disabled={!value.FillInChoice}
                                                                                    value={this.state.Data[value.internalName] === "Specify your own value :" ? this.state.SpecifyOwnvalues[value.internalName] || "" : ""}
                                                                                    style={{ outline: "none" }}
                                                                                    className={styles.eDIBSpecifyOwnvalueInputdrp}
                                                                                    onChange={(event) => this.SpecifyeventHandler(event, index)}
                                                                                /> : null}

                                                                            </div>
                                                                        </div>
                                                                    )
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
                                                                                        className='inputText'
                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                    />
                                                                                </div>
                                                                            </div>
                                                                        </div>
                                                                    )
                                                                case "Number":
                                                                    return (
                                                                        <div className='fieldEditor' key={index}>
                                                                            <div>
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
                                                                                        className='inputText'
                                                                                        onChange={(event) => this.eventHandler(event, index)}
                                                                                    />
                                                                                </div>
                                                                            </div>
                                                                        </div>
                                                                    )
                                                                case "UserMulti"
                                                                    :
                                                                    return (
                                                                        <div className='fieldEditor' key={index}>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className='feildDisplay'>

                                                                                <div className='peoplepicker'>
                                                                                    <PeoplePicker
                                                                                        context={this.props.context}
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
                                </fieldset>
                            </PivotItem>
                            <PivotItem headerText="Attachments">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className='fieldEditor'>
                                        <div style={{ width: "50%" }}>
                                            <label className='label' htmlFor='Attachments'>Upload Attachments{this.props.IsAttachmentsRequired === true ? <span className='labelIcon'>*</span> : null} </label>
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
                                    </div>
                                    <div className={styles.viewfeildFilesDisplay}>
                                        <span>
                                            <ul style={{ margin: "unset", padding: "unset" }}>
                                                {/* {this.state.attachfiles.map((file, ind) => */}
                                                {this.state.getattchements.map((file, ind) =>
                                                    <div className={styles.viewfeildFilesDisplay} key={ind}>
                                                        {file.isExists ? (
                                                            <>
                                                                <span className={styles.attachmentSpan}>
                                                                    <IconButton
                                                                        onClick={() => this.onRemoveAttachments(file)}
                                                                        iconProps={{ iconName: 'Delete' }}
                                                                        styles={{
                                                                            icon: { fontSize: 18 },
                                                                            root: {
                                                                            },
                                                                            rootHovered: { backgroundColor: "white" },
                                                                            rootPressed: { backgroundColor: 'white' }
                                                                        }}
                                                                    />
                                                                </span>
                                                                <span style={{ width: "200px" }}>
                                                                    <li style={{ wordWrap: "break-word" }} key={ind}>{file.name}</li>
                                                                </span>
                                                                <span className={styles.attachmentSpan}>
                                                                    <a href={file.fileUrl} download={file.name}>
                                                                        <IconButton
                                                                            iconProps={{ iconName: 'Download' }}
                                                                            styles={{
                                                                                icon: { fontSize: 18 },
                                                                                root: {
                                                                                },
                                                                                rootHovered: { backgroundColor: "white" },
                                                                                rootPressed: { backgroundColor: 'white' }
                                                                            }}
                                                                        />
                                                                    </a>
                                                                </span>
                                                            </>
                                                        ) :
                                                            (
                                                                <>
                                                                    <span className={styles.attachmentSpan}>
                                                                        <IconButton
                                                                            onClick={() => this.onRemoveAttachments(file)}
                                                                            iconProps={{ iconName: 'Delete' }}
                                                                            styles={{
                                                                                icon: { fontSize: 18 },
                                                                                root: {
                                                                                },
                                                                                rootHovered: { backgroundColor: "white" },
                                                                                rootPressed: { backgroundColor: 'white' }
                                                                            }}
                                                                        />

                                                                    </span>

                                                                    <span style={{ width: "200px" }}>
                                                                        <li key={ind}>{file.name}</li>
                                                                    </span>
                                                                </>)}
                                                    </div>)}
                                            </ul>
                                        </span>
                                    </div>
                                </fieldset>
                            </PivotItem>
                            <PivotItem headerText="Approval">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    {this.state.approverFields.map((ele, ind) => {
                                        if (ele.internalName === "ApprFldReletedPerson10Id") {
                                            return (
                                                <>
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
                                                                        principalTypes={[PrincipalType.SharePointGroup, PrincipalType.User, PrincipalType.SecurityGroup]}
                                                                        resolveDelay={1000} />
                                                                </div>
                                                            </div>
                                                        </div>
                                                    }</>)
                                        }
                                        else {
                                            return (
                                                <>
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
                                                                        onChange={this._getPeoplePickerItems.bind(this, ele.internalName)}
                                                                        showHiddenInUI={false}
                                                                        ensureUser={true}
                                                                        principalTypes={[PrincipalType.User]}
                                                                        resolveDelay={1000} />
                                                                </div>
                                                            </div>
                                                        </div>
                                                    }</>)

                                        }

                                    })}

                                </fieldset>
                            </PivotItem>
                            <PivotItem headerText="Version History">
                                <fieldset className='sonyfieldSet'>
                                    <div className='ms-Grid'>
                                        <div className='ms-Grid-row' style={{ "marginTop": "5px", overflowY: "scroll", height: 200 }}>
                                            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                                {activityItemExamples.map((item: { key: string | number }) => (
                                                    <ActivityItem {...item} key={item.key} className={classNames.exampleRoot} />
                                                ))}
                                            </div>
                                        </div>
                                    </div>
                                </fieldset>
                                <span className={styles.buutonsCont}> <Link href={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/Versions.aspx?list=${this.props.listId.id}&ID=${getIdFromUrl()}`}>Click here for more information</Link></span>

                            </PivotItem>
                        </Pivot>
                      
                        <div className='Spanbutton'>
                            <span style={{ marginRight: "1%" }} hidden={!((this.state.Data.StatusNo === "5" || this.state.Data.StatusNo === "10000" || this.state.Data.StatusNo === "15") && this.state.RequsterEmail === this._UserEmail)}><PrimaryButton type='reset' text={(this.state.Data.StatusNo === "15" || this.state.Data.StatusNo === "10000") ? "Resubmit" : "Submit"} onClick={this.onSubmitData} /></span>
                            <span style={{ marginRight: "1%" }} hidden={!((this.state.Data.StatusNo === "5" || this.state.Data.StatusNo === "10000" || this.state.Data.StatusNo === "15") && this.state.RequsterEmail === this._UserEmail)}><PrimaryButton type='reset' text='Draft' onClick={this.onDraft} /></span>
                            <span><PrimaryButton type='reset' text='Exit' onClick={this.homePage} /></span>
                            {/* </div> */}
                        </div>
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
                                    <p>Please fill below details.</p>
                                    <ul>
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
                    </div>
                </section>
            )

        }
    }

}
