import * as React from 'react';
import styles from './SonyEdib.module.scss';
import { ISonyEdibProps } from './ISonyEdibProps';
import { ActivityItem, Dialog, DialogFooter, DialogType, Icon, IconButton, IPersonaProps, Link, mergeStyleSets, Persona, PersonaSize, Pivot, PivotItem, PrimaryButton, TextField } from '@fluentui/react';
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
import "@pnp/sp/site-users/web";
import { IAppREmails, IBoardAppr } from './Createform';
import { IAuditTrail } from './EditForm';
import BoardApprovers from './boradApproves';
// import { SPServiceWorker, ListController } from '@pnp/spfx-controls-react/lib/SPServiceWorker';

// import { SPFabricCoreStyles } from '@microsoft/sp-office-ui-fabric-core';
export interface IFieldCollection {
    Title: string,
    DataType: string,
    Required: boolean,
    Tab: string,
    internalName: string,
    Option?: string[];
    FillInChoice?: boolean;
}

export interface IApproverField {
    Title: string,
    internalName: string,
    Required: boolean,
    Email?: string[],
    levelType?: string,
    Disable?: boolean
}
export interface IPeoplePickerItems {
    id: number;
    loginName?: string;
    secondaryText?: string;

}
export interface IApproInfo {
    name: string;
    jobRole: string;
    email: string;
}
export interface IAlert {
    ApprMsg?: boolean;
    RjctMsg?: boolean
}
export type UserProfileProperties = {
    Key: string;
    Value: string;
}[];
export interface IGeFiles {
    fileName: string;
    fileUrl: string

}
export interface IData {
    [x: string]: string | number[] | string
}

export interface ISonyEdibState {
    fieldCollection: IFieldCollection[],
    Data: any,
    getItems: any,
    auditLog: IAuditTrail[],
    approverFields: IApproverField[];
    approversEmail: IAppREmails;
    userRole: string;
    UserEmail: string;
    ApprbtnHideShow: boolean;
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
    alertMsg: IAlert[];
    isAdminGroupUser: boolean;
    boardApprvers: IBoardAppr[];
    boardBtn: boolean;
    RequsterEmail: string;
    GetAttchment: IGeFiles[];
    Comments: string;
    hideComentDialog: boolean;
    hideComentTab?:boolean;


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

export default class ViewForm extends React.Component<ISonyEdibProps, ISonyEdibState, {}> {
    private _spService: spService = null;
    private _sp;
    private _userName: string;
    private _userEmail: string;
    private _UserRole: string;
    private _userpictureUrl: string;
    private _userfirstName: string;
    private _userlastName: string;
    private _folderName: string;
    private _itemId: number = Number(getIdFromUrl());
    constructor(props: ISonyEdibProps) {
        super(props);
        this.state = {
            fieldCollection: [],
            Data: {},
            getItems: {},
            auditLog: [],
            approverFields: [],
            approversEmail: {},
            userRole: "",
            UserEmail: "",
            ApprbtnHideShow: true,
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
            hideComentDialog: true,
            alertMsg: [],
            isAdminGroupUser: false,
            boardApprvers: [],
            boardBtn: true,
            RequsterEmail: "",
            GetAttchment: [],
            Comments: "",
            hideComentTab:true
        }

        this._sp = spfi().using(SPFx(this.props.context))
        this._spService = new spService(this.props.context);
        Promise.all([
            this.getSiteGroups(),
            this.GetUserProperties(),
            this.getApproverEmail(),
            this.getItems(),
            // this.approvarFields()
        ]).catch(err => console.error(err));

    }

    // get user details
    private GetUserProperties = async (): Promise<void> => {
        await this._sp.profiles.myProperties().then((result: { UserProfileProperties: UserProfileProperties; DisplayName: string; }) => {
            const props = result.UserProfileProperties;
            this._userName = result.DisplayName;
            for (let i = 0; i < props.length; i++) {
                const allProperties = props[i];
                if (allProperties.Key === "PictureURL") {
                    this._userpictureUrl = allProperties.Value;
                }
                else if (allProperties.Key === "Title") {
                    this._UserRole = allProperties.Value;
                    // console.log(reqRole);
                }
                else if (allProperties.Key === "FirstName") {
                    const frstname = allProperties.Value;
                    this._userfirstName = frstname.substring(0, 1);
                    //console.log(firstletterFN);
                }
                else if (allProperties.Key === "LastName") {
                    const lastname = allProperties.Value;
                    this._userlastName = lastname.substring(0, 1);
                    //console.log(firstletterLN);
                }
                else if (allProperties.Key === "WorkEmail") {
                    this._userEmail = allProperties.Value;
                }
            }

            this.setState({
                FullName: this._userName,
                userDepartment: "",
                pictureUrl: this._userpictureUrl,
                firstName: this._userfirstName,
                lastName: this._userlastName,
                reqRole: this._UserRole,
                imageinitials: this._userfirstName + this._userlastName,
                UserEmail: this._userEmail
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

    // get by items by items id
    public getItems = async (): Promise<void> => {
        // alert(true);
        await this._spService.getListItemsById(this.props.listName, this._itemId).then(res => {
            this.setState({
                Data: res,
                auditLog: JSON.parse(res.AuditTrail),
                boardApprvers: JSON.parse(res.BoardApprovers),
                RequsterEmail: res.Author.EMail
            });
            // console.log(res.Comments);
            this._folderName = res.Title;

            // const isExists = this.state.boardApprvers.some(obj => obj.CCT.Status === "CCT Changed" || obj.PWB.Status === "PWB Changed");
            // this.setState({ boardBtn: !isExists });
            // console.log("hello", res.BoardApprovers);
            const audit: IAuditTrail[] = JSON.parse(res.AuditTrail);
            if (audit !== null) {
                activityItemExamples.length = 0
                audit.forEach((ele) => {
                    if (ele.ActionTaken === "Submitted") {
                        activityItemExamples.unshift({
                            key: Math.floor(Math.random() * 10000),
                            activityDescription: [
                                <span className={classNames.nameText} key={1}>{ele.Actioner}</span>,
                                <span className={classNames.space} key={2}>{ele.ActionTaken}</span>,
                                <div key={3}><span key={3}>Role: {ele.Role}</span></div>,

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
                                <div key={3}><span key={3}>Role: {ele.Role}</span></div>,

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
            this.formInput().then(res => res).catch(err => err);

            this.getFilesInsideFolder(this._folderName).then(res => res).catch(err => err);

        });

        // this.approvarFields().then(res => res).catch(err => err);
    }

    // Configuration input fields
    public formInput = async (): Promise<void> => {
        const collectionfileds = this.props.collectionData;
        const fieldIfo: { Title: string; DataType: string; Required: boolean; Tab: string; Option: string[]; internalName: string; FillInChoice?: boolean }[] = []
        await this._spService.getfieldDetails(this.props.listName).then((res) => {
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
                            });
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
            const updatedFieldInfo = fieldIfo.map(x => {
                if (typeof this.state.Data[x.internalName] !== "undefined" && this.state.Data[x.internalName] !== null) {
                    if (x.DataType === "MultiChoice") {
                        // console.log(this.state.Data[x.Title]);

                        // Find the corresponding object in fieldIfo and update its Option property
                        const updatedObject = fieldIfo.find(obj => obj.internalName === x.internalName);

                        if (updatedObject) {
                            // Update the Option property with the values from this.state.Data
                            updatedObject.Option = [...this.state.Data[x.internalName]];
                        }

                        return updatedObject; // Return the updated object
                    } 
                    else if(x.DataType === "Choice"){
                        const updatedObject = fieldIfo.find(obj => obj.internalName === x.internalName);

                        if (updatedObject) {
                            // Update the Option property with the values from this.state.Data
                            updatedObject.Option = [this.state.Data[x.internalName]];
                        }

                        return updatedObject; 

                    }
                    else {
                        return x; // For other cases, return the original object
                    }
                } else {
                    return x; // For objects that don't meet the conditions, return the original object
                }
            });

            // console.log(updatedFieldInfo); // Contains the updated fieldIfo array


            // this.setState({ fieldCollection: fieldIfo });
            this.setState({ fieldCollection: updatedFieldInfo });
            // this.approvarFields().then(res => res).catch(err => err);updatedFieldInfo
        }).catch(error => {
            console.log("Something went wrong! please contact admin for more information.", error);
        });
        // this.getItems().then(res => res).catch(err => err);
    }

    // get the approers fileds
    public approvarFields = async (): Promise<void> => {
        await this._spService.apprConfigu(this.props.listName).then(res => this.setState({ approverFields: res })).catch(err => err);

    }

    // get approver Email

    public getApproverDetails = async (): Promise<void> => {
        const listTitle: string = this.props.listName;

        //     const apprTemp: string[] = []
        await this._sp.web.lists.getByTitle(listTitle).items.getById(this._itemId)
            .select(`ApprFldReletedPerson1/EMail,ApprFldReletedPerson2/EMail,ApprFldReletedPerson3/EMail,ApprFldReletedPerson4/EMail,ApprFldReletedPerson5/EMail,ApprFldReletedPerson6/EMail,ApprFldReletedPerson7/EMail,ApprFldReletedPerson8/EMail,ApprFldReletedPerson9/EMail`)
            .expand(`ApprFldReletedPerson1,ApprFldReletedPerson2,ApprFldReletedPerson3,ApprFldReletedPerson4,ApprFldReletedPerson5,ApprFldReletedPerson6,ApprFldReletedPerson7,ApprFldReletedPerson8,ApprFldReletedPerson9`)()
            .then(result => {
                if (typeof result.ApprFldReletedPerson1 !== "undefined") {
                    const tempemail1: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson1.length; i++) {
                        tempemail1.push(result.ApprFldReletedPerson1[i].EMail);

                    }
                    this.setState({ Approver1: [...this.state.Approver1, ...tempemail1] });
                }
                if (typeof result.ApprFldReletedPerson2 !== "undefined") {
                    const tempemail2: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson2.length; i++) {
                        tempemail2.push(result.ApprFldReletedPerson2[i].EMail);
                    }
                    this.setState({ Approver2: [...this.state.Approver2, ...tempemail2] });

                }
                if (typeof result.ApprFldReletedPerson3 !== "undefined") {
                    const tempemail3: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson3.length; i++) {
                        tempemail3.push(result.ApprFldReletedPerson3[i].EMail);
                    }
                    this.setState({ Approver3: [...this.state.Approver3, ...tempemail3] });

                }
                if (typeof result.ApprFldReletedPerson4 !== "undefined") {
                    const tempemail4: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson4.length; i++) {
                        tempemail4.push(result.ApprFldReletedPerson4[i].EMail);
                    }
                    this.setState({ Approver4: [...this.state.Approver4, ...tempemail4] });

                }
                if (typeof result.ApprFldReletedPerson5 !== "undefined") {
                    const tempemail5: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson5.length; i++) {
                        tempemail5.push(result.ApprFldReletedPerson5[i].EMail);
                    }
                    this.setState({ Approver5: [...this.state.Approver5, ...tempemail5] });

                }
                if (typeof result.ApprFldReletedPerson6 !== "undefined") {
                    const tempemail6: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson6.length; i++) {
                        tempemail6.push(result.ApprFldReletedPerson6[i].EMail);
                    }
                    this.setState({ Approver6: [...this.state.Approver6, ...tempemail6] });
                    // console.log("test", this.state.Approver6);

                }
                if (typeof result.ApprFldReletedPerson7 !== "undefined") {
                    const tempemail7: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson7.length; i++) {
                        tempemail7.push(result.ApprFldReletedPerson7[i].EMail);
                    }
                    this.setState({ Approver7: [...this.state.Approver7, ...tempemail7] });
                    // console.log("test", this.state.Approver7);

                }
                if (typeof result.ApprFldReletedPerson8 !== "undefined") {
                    const tempemail8: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson8.length; i++) {
                        tempemail8.push(result.ApprFldReletedPerson8[i].EMail);
                    }
                    this.setState({ Approver8: [...this.state.Approver8, ...tempemail8] });
                    // console.log("test", this.state.Approver8);
                }
                if (typeof result.ApprFldReletedPerson9 !== "undefined") {
                    const tempemail9: string[] = [];
                    for (let i = 0; i < result.ApprFldReletedPerson9.length; i++) {
                        tempemail9.push(result.ApprFldReletedPerson9[i].EMail);
                    }
                    this.setState({ Approver9: [...this.state.Approver9, ...tempemail9] });
                    // console.log("test", this.state.Approver9);

                }
            })

    }


    // get approvers
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
                this.setState(prevState => ({
                    approversEmail: {
                        ...prevState.approversEmail,
                        [`${fieldTitle}Id`]: apprTemp
                    }
                }));
            }
            this.approvarFields().then(res => res).catch(err => err)
        }
        this.getApproverDetails().then(res => res).catch(err => err);
    }

    // get current  user groups
    private getSiteGroups = async (): Promise<void> => {
        const group = await this._sp.web.currentUser.groups();
        // console.log(group);
        const gropTitles = group.map(x => x.Title);
        if (gropTitles.includes(this.props.DibGnAdminGroup)) { //here we need mention admingrop title ("Xencia Demo Apps Members") 
            this.setState({ isAdminGroupUser: true });
        }
        // console.log("group", group);
        // await this._sp.web.siteUsers.filter(`IsSiteAdmin eq false`)().then(res => console.log(res.map(x => x.Email)));
    }
    // Previous Approvers
    public previousApprvers = async (): Promise<any> => {
        const temp: number[] = [];
        const { approverFields, Data } = this.state;
        const statusNoArray = ["9000", "8000", "7000", "6000", "5000", "4000", "3000", "2000", "1000"];
        const filterdStatusNo = statusNoArray.filter(x => Number(x) < Number(this.state.Data.StatusNo));
        // console.log(filterdStatusNo);
        // let istrue = false;
        const apprverInternalName = approverFields.map(x => x.internalName);


        if (filterdStatusNo.length > 0) {
            for (let i = 0; i < filterdStatusNo.length; i++) {
                if (filterdStatusNo[i] === "9000" && this.state.Approver9.length > 0 && !this.state.approverFields[8].Disable) {
                    temp.push(...Data[apprverInternalName[8]]);
                    break;
                }
                else if (filterdStatusNo[i] === "8000" && this.state.Approver8.length > 0 && !this.state.approverFields[7].Disable) {
                    temp.push(...Data[apprverInternalName[7]]);
                    break;
                }
                else if (filterdStatusNo[i] === "7000" && this.state.Approver7.length > 0 && !this.state.approverFields[6].Disable) {
                    temp.push(...Data[apprverInternalName[6]]);
                    break;
                }
                else if (filterdStatusNo[i] === "6000" && this.state.Approver6.length > 0 && !this.state.approverFields[5].Disable) {

                    temp.push(...Data[apprverInternalName[5]]);
                    break;
                }
                else if (filterdStatusNo[i] === "5000" && this.state.Approver5.length > 0 && !this.state.approverFields[4].Disable) {
                    temp.push(...Data[apprverInternalName[4]]);
                    break;
                }
                else if (filterdStatusNo[i] === "4000" && this.state.Approver4.length > 0 && !this.state.approverFields[3].Disable) {
                    temp.push(...Data[apprverInternalName[3]]);
                    break;
                }
                else if (filterdStatusNo[i] === "3000" && this.state.Approver3.length > 0 && !this.state.approverFields[2].Disable) {
                    temp.push(...Data[apprverInternalName[2]]);
                    break;
                }
                else if (filterdStatusNo[i] === "2000" && this.state.Approver2.length > 0 && !this.state.approverFields[1].Disable) {
                    // alert("2000")
                    temp.push(...Data[apprverInternalName[1]]);
                    break;
                }

                else if (filterdStatusNo[i] === "1000" && this.state.Approver1.length > 0 && !this.state.approverFields[0].Disable) {
                    // alert("10000")
                    temp.push(...Data[apprverInternalName[0]]);
                    break;
                }
            }
        } else {
            temp.length = 0

        }
        // console.log(temp);
        return temp

    }


    // Approve button functionality
    public apprFunctionality = async (): Promise<void> => {
        const statusNoArray = ["1000", "2000", "3000", "4000", "5000", "6000", "7000", "8000", "9000"];
        const filterdStatusNo = statusNoArray.filter(x => Number(x) > Number(this.state.Data.StatusNo));
        let statusNo;
        let status;
        let isExit: boolean;
        let dibId;
        const privouseApprver: number[] = []
        const auditlog = this.state.auditLog;
        const currentStatus: string = this.state.Data.Status;
        const approvarStage: string = currentStatus.replace("Pending for", " ");
        // console.log(approvarStage);

        if (filterdStatusNo.length > 0) {
            for (let i = 0; i < filterdStatusNo.length; i++) {
                if (filterdStatusNo[i] === "2000" && this.state.Approver2.length > 0 && !this.state.approverFields[1].Disable) {
                    statusNo = "2000";
                    // status = "Pending for Related Person 2";
                    status = `Pending for ${this.state.approverFields[1].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "3000" && this.state.Approver3.length > 0 && !this.state.approverFields[2].Disable) {
                    statusNo = "3000";
                    // status = "Pending for Related Person 3";
                    status = `Pending for ${this.state.approverFields[2].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "4000" && this.state.Approver4.length > 0 && !this.state.approverFields[3].Disable) {
                    statusNo = "4000";
                    // status = "Pending for Related Person 4";
                    status = `Pending for ${this.state.approverFields[3].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "5000" && this.state.Approver5.length > 0 && !this.state.approverFields[4].Disable) {
                    statusNo = "5000";
                    // status = "Pending for Related Person 5";
                    status = `Pending for ${this.state.approverFields[4].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "6000" && this.state.Approver6.length > 0 && !this.state.approverFields[5].Disable) {
                    statusNo = "6000";
                    // status = "Pending for  group leader";
                    status = `Pending for ${this.state.approverFields[5].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "7000" && this.state.Approver7.length > 0 && !this.state.approverFields[6].Disable) {
                    statusNo = "7000";
                    // status = "Pending for Borad Leader";
                    status = `Pending for ${this.state.approverFields[6].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "8000" && this.state.Approver8.length > 0 && !this.state.approverFields[7].Disable) {

                    statusNo = "8000";
                    // status = "Pending for ASE";
                    status = `Pending for ${this.state.approverFields[7].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
                else if (filterdStatusNo[i] === "9000" && this.state.Approver9.length > 0 && !this.state.approverFields[8].Disable) {

                    statusNo = "9000";
                    // status = "Pending for Chassis Leader";
                    status = `Pending for ${this.state.approverFields[8].Title}`;
                    const obj = {
                        "Actioner": this._userName,
                        "ActionTaken": `Approved by ${approvarStage}`,
                        "Role": this._UserRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": this.state.Comments
                    };
                    isExit = true;
                    auditlog.push(obj);
                    break;
                }
            }
        }
        else {
            statusNo = "11000";
            status = "Officially Released";
            dibId = this.state.Data.Title
            const obj = {
                "Actioner": this._userName,
                "ActionTaken": `Officially Released`,
                "Role": this._UserRole,
                "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                "Comments": this.state.Comments
            };
            auditlog.push(obj);

            // if (this.props.isBoardApprovalsRequired === true) {
            //     const obj1 = {
            //         "Actioner": this._userName,
            //         "ActionTaken": "Pending for Board Approval",
            //         "Role": this._UserRole,
            //         "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
            //         "Comments": ""
            //     };
            //     status = "Pending for Board Approval";
            //     statusNo = "12000";
            //     auditlog.push(obj1);
            // }
            isExit = true;
        }
        if (isExit === true) {
            await this.previousApprvers().then(res => {
                res.map((x: any) => privouseApprver.push(x));
                // console.log(privouseApprver)
            }).catch(err => err)

        }
        const updatedListObj = {
            Status: status,
            StatusNo: statusNo,
            Comments: this.state.Comments,
            AuditTrail: JSON.stringify(auditlog),
            StartProcessing: true,
            "DIB_ID": dibId,
            PreviousApprovarsId: privouseApprver
        }

        if (isExit === true) {
            const temp: IAlert[] = []
            await this._spService.updateListItem(this.props.listName, updatedListObj, this._itemId).then(res => {
                temp.push({
                    ApprMsg: true,
                    RjctMsg: false
                });
                // this.previousApprvers().then(res => res).catch(err => err);
                this.toggleHideDialog();
                return res
            }).catch(err => err);
            this.setState({
                alertMsg: temp
            });
        }
    }
    private dialogcontent = {
        type: DialogType.normal,
        title: 'Information!'
    }
    private dialogcontentComments = {
        type: DialogType.normal,
        title: 'Comments'
    }

    // comments
    public handleTextareaChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ Comments: newValue });
        // console.log(newValue.trim());
    }


    // Reject button functionalities
    public rejectBtn = async (): Promise<void> => {
        if (this.state.Comments !== "") {
            // if (this.state.Comments !== "" && this.state.Comments.trim().length > 0) {
            const temp: IAlert[] = []
            const statusNo = "10000";
            const status = "Rejected";
            const auditlog = this.state.auditLog;
            const currentStatus: string = this.state.Data.Status;
            const approvarStage: string = currentStatus.replace("Pending for", "");
            const obj = {
                "Actioner": this._userName,
                "ActionTaken": `Rejected by ${approvarStage}`,
                "Role": this._UserRole,
                "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                "Comments": this.state.Comments
            };
            auditlog.push(obj);
            const updatedListObj = {
                Status: status,
                StatusNo: statusNo,
                Comments: this.state.Comments,
                StartProcessing: true,
                AuditTrail: JSON.stringify(auditlog)
            }

            await this._spService.updateListItem(this.props.listName, updatedListObj, this._itemId).then(res => {
                temp.push({
                    ApprMsg: false,
                    RjctMsg: true
                });
                this.toggleHideDialog();
                return res
            }).catch(err => err);
            this.setState({ alertMsg: temp });

        }
        else {
            // this.commentstoggleHideDialog();
            this.commentstoggleHideDialogCMTTab();
        }

    }

    // Dailogs for notification Approve & reject  hideComentDialog
    public toggleHideDialog = (): void => {
        this.setState({ hideDialog: false });
    }
    public commentstoggleHideDialog = (): void => {
        this.setState({ hideComentDialog: false });

    } 
    // comments tab popup
    public commentstoggleHideDialogCMTTab = (): void => {
        this.setState({ hideComentTab: false });

    } 



    // cancell
    public onCancel = (): void => {
        this.setState({ hideDialog: true });
        this.Approvalpage();
    }
    public getFilesInsideFolder = async (folderName: string): Promise<void> => {
        const files = await this._sp.web.getFolderByServerRelativePath(this.props.context.pageContext.web.serverRelativeUrl + `/${this.props.libraryName}/` + folderName).files();
        const tenantUrl = window.location.protocol + "//" + window.location.host;
        const tempFiles: { fileName: string; fileUrl: string; }[] = []
        files.forEach(values => {
            // let temp=[];
            const filesObj = {
                fileName: values.Name,
                fileUrl: tenantUrl + values.ServerRelativeUrl
            }
            tempFiles.push(filesObj);
            // console.log("temp", tempFiles);
        });
        this.setState({ GetAttchment: tempFiles });
    }
    
    // pendingpage
    public Approvalpage = (): void => {
        const pageURL: string = this.props.context.pageContext.web.absoluteUrl;
        window.location.href = `${pageURL}/SitePages/MyDIBPendingRequests.aspx`;
        this.setState({ hideDialog: true });
    }

    public homePage = (): void => {
        const pageURL: string = this.props.context.pageContext.web.absoluteUrl;
    //    here checking previoues page  if its not exist redrict to home paga
        if(window.history.length>1){
            window.history.go(-1)
        }
        else{
            window.location.href = pageURL;
        }
        this.setState({ hideDialog: true });
    }
    // edit form
    public editPage = (): void => {
        const pageURL: string = this.props.context.pageContext.web.absoluteUrl;
        // console.log(pageURL);
        window.location.href = `${pageURL}/SitePages/DibEdit.aspx?itemId=${this._itemId}`;
    }
    // comments validation
    public commentsvalidation=():void=>{
        if(this.state.Comments !==""){
            this.rejectBtn().then(res=>res).catch(err=>err);
            this.setState({hideComentTab:true});
        }
        else{
            this.commentstoggleHideDialog();
        }
    }
    // compontDidmout
    public componentDidMount = (): void => {
        // this.approvarFields().then(res => res).catch(err => err);
        // alert(this.props.isBoardApprovalsRequired);
        // this.formInput().then(res => res).catch(err => err);
        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.ready
                .then((registration) => {
                    if (registration.active) {
                        registration.active.postMessage({ action: 'cleanupIndexedDB' });
                        //   console.log('IndexedDB cleanup message sent to the service worker.');
                    }
                    //  else {
                    //   console.error('Service worker is not active.');
                    // }
                })
                .catch((error) => {
                    console.error('Error while accessing service worker:', error);
                });
        }
        //    else {
        //     console.error('Service Worker is not supported.');
        //   }
    }

    public render(): React.ReactElement<ISonyEdibProps> {
        const {

            hasTeamsContext,
        } = this.props;

        // dialogContent
        const dialogContentProps = {
            type: DialogType.normal,
            title: 'Information!',
        }

        return (
            <section className={`${styles.sonyEdib} ${hasTeamsContext ? styles.teams : ''}`} >
                <div className={styles.sonydibcontainer}>
                    <div className={styles.frmtitle}>View Form</div>
                    <Pivot aria-label="Large Link Size Pivot Example" >
                        {/*---- Section 1----- */}
                        <PivotItem headerText="Requestor Details">
                            <fieldset style={{ height: 60, "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                <div className="personarow" style={{ paddingTop: 5 }}>
                                    <div className="personaCol1">
                                        <div className=" ms-sm12 ms-lg12">
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
                                    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 personaCol2">
                                        <div className="ms-Grid-row presonaCol2Row" >
                                            <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg2 presonaCol2col1" >
                                                <span className="hdrTtle userName">Status: </span>
                                            </div>
                                            <div className="ms-Grid-col ms-sm8 ms-md9 ms-lg10 presonaCol2col2">
                                                <span className='userName'>{this.state.Data.Status}</span>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </fieldset>
                        </PivotItem>
                    </Pivot>
                    < form >
                        {/* Section 2---- */}
                        <Pivot aria-label="Large Link Size Pivot Example" >
                            <PivotItem headerText="DIB content">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className='ms-Grid'>
                                        <div className='ms-Grid-row'>
                                            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                    {this.state.fieldCollection.map((value, index) => {
                                                        if (value.Tab === "DIB Content") {
                                                            switch (value.DataType) {
                                                                case "MultiChoice":
                                                                    return (
                                                                        <div className='fieldEditor'>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className='radioCheckbox'>
                                                                                {value.Option.map((ele: string) => {
                                                                                    return (
                                                                                        <div key={index} >

                                                                                            <input type="checkbox" name={value.internalName} value={ele}
                                                                                                checked={(this.state.Data[value.internalName] !== null) ? Object.values(this.state.Data[value.internalName]).includes(ele) : false}
                                                                                                // checked={(this.state.Data[value.internalName] !== null ||this.state.Data[value.internalName] !== undefined ) ? Object.values(this.state.Data[value.internalName]).includes(ele) : false}
                                                                                                required={value.Required}
                                                                                                disabled
                                                                                            />{ele}
                                                                                        </div>
                                                                                    )
                                                                                })}
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
                                                                                {value.Option.map((ele: string) => {
                                                                                    return (
                                                                                        <div key={index}>
                                                                                            <input type="radio" name={value.internalName} value={ele}
                                                                                                checked={this.state.Data[value.internalName] !== null ? this.state.Data[value.internalName] === ele ? true : false : false}
                                                                                                required={value.Required}
                                                                                                disabled />{ele}

                                                                                        </div>
                                                                                    )
                                                                                })}
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
                                                                                {/* {this.state.Data[value.internalName]} */}
                                                                                {/* <div className='fieldGroup'> */}
                                                                                <textarea
                                                                                    name={value.internalName}
                                                                                    id={value.internalName}
                                                                                    value={this.state.Data[value.internalName] || ""}
                                                                                    style={{ border: "none", outline: "none", resize: "none", overflow: "unset" }}
                                                                                    required={value.Required}
                                                                                    rows={3}
                                                                                    readOnly
                                                                                    className="textarea"

                                                                                // onChange={(event) => this.handleTextareaChange(event, index)}
                                                                                />
                                                                                {/* <div  dangerouslySetInnerHTML={{ __html: (this.state.Data[value.internalName]) }} /> */}
                                                                            </div>
                                                                            {/* </div> */}
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
                                                                                        <div key={index} >
                                                                                            <input type="radio" name={value.internalName} value={ele}
                                                                                                // onChange={(event) => this.eventYesOrNoHandler(event, index)}
                                                                                                // checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                checked={(this.state.Data[value.internalName] !== null && this.state.Data[value.internalName].toString() === ele) ? true : false}
                                                                                                required={value.Required} disabled />
                                                                                            {/* <div className={styles.radiocheckmark}>
                                                                                                    <span className={styles.radioinsidecircle} />
                                                                                                </div> */}
                                                                                            {ele === "true" ? "Yes" : "No"}
                                                                                        </div>
                                                                                    )
                                                                                })}
                                                                            </div>
                                                                        </div>
                                                                    )

                                                                // _____________________________________________________________________________________________________________________________________-
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
                                                                                        // titleText={value.Title}
                                                                                        personSelectionLimit={3}
                                                                                        groupName={""} // Leave this blank in case you want to filter from all users
                                                                                        showtooltip={true}
                                                                                        required={value.Required}
                                                                                        disabled={true}
                                                                                        defaultSelectedUsers={Object.keys(this.state.approversEmail).includes(value.internalName) ? this.state.approversEmail[value.internalName] : null}
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
                                                                                showHiddenInUI={false}
                                                                                principalTypes={[PrincipalType.User]}
                                                                                resolveDelay={1000}
                                                                                disabled={true} />
                                                                        </div>
                                                                    )
                                                                case "DateTime":
                                                                    return (
                                                                        <div className='DibRelfieldEditor' key={index}>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className={styles.viewfielddisply}>
                                                                                <label>{(this.state.Data[value.internalName] !== undefined) ? this.state.Data[value.internalName].slice(0, 10) : null}</label>
                                                                            </div>
                                                                        </div>

                                                                    )
                                                                default:
                                                                    if (value.internalName !== "DIB_ID") {
                                                                        return (
                                                                            <div className='DibRelfieldEditor' key={index}>
                                                                                <div>
                                                                                    <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                </div>
                                                                                <div className={styles.viewfielddisply}>
                                                                                    <label>{this.state.Data[value.internalName] || ""}</label>
                                                                                </div>
                                                                            </div>
                                                                        )

                                                                    }
                                                                    else if (value.internalName === "DIB_ID" && (this.state.Data.StatusNo === "12000" || this.state.Data.StatusNo === "11000" || this.state.Data.StatusNo === "13000")) {
                                                                        return (
                                                                            <div className='DibRelfieldEditor' key={index}>
                                                                                <div>
                                                                                    <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                                </div>
                                                                                <div className={styles.viewfielddisply}>
                                                                                    <label>{this.state.Data[value.internalName] || ""}</label>
                                                                                </div>
                                                                            </div>
                                                                        )

                                                                    }

                                                            }
                                                        }

                                                    })}
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </fieldset>
                            </PivotItem>
                            {/* Section 3 */}
                            <PivotItem headerText="Related Check Items">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className='ms-Grid'>
                                        <div className='ms-Grid-row'>
                                            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                                                    {this.state.fieldCollection.map((value, index) => {
                                                        if (value.Tab === "Related check items") {
                                                            switch (value.DataType) {
                                                                case "MultiChoice":
                                                                    return (
                                                                        <div className='fieldEditor'>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className='radioCheckbox'>
                                                                                {value.Option.map((ele: string) => {
                                                                                    return (
                                                                                        <div key={index} >

                                                                                            <input type="checkbox" name={value.internalName} value={ele}
                                                                                                checked={this.state.Data[value.internalName] !== null ? Object.values(this.state.Data[value.internalName]).includes(ele) : false}
                                                                                                required={value.Required}
                                                                                                disabled
                                                                                            />{ele}
                                                                                            {/* <span className={styles.checkmark}></span> */}
                                                                                        </div>
                                                                                    )
                                                                                })}
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
                                                                                    style={{ border: "none", outline: "none", resize: "none" }}
                                                                                    required={value.Required}
                                                                                    readOnly
                                                                                    rows={3}
                                                                                    className="textarea"

                                                                                // onChange={(event) => this.handleTextareaChange(event, index)}
                                                                                />
                                                                            </div>
                                                                            {/* </div> */}
                                                                        </div>
                                                                    )
                                                                case "Choice":
                                                                    return (
                                                                        <div className='fieldEditor'>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className='radioCheckbox'>
                                                                                {value.Option.map((ele: string) => {
                                                                                    return (
                                                                                        <div key={index} >
                                                                                            <input type="radio" name={value.internalName} value={ele}
                                                                                                checked={this.state.Data[value.internalName] !== null ? this.state.Data[value.internalName] === ele ? true : false : false}
                                                                                                required={value.Required}
                                                                                                disabled />{ele}
                                                                                            {/* <span className={styles.radiocheckmark}></span> */}

                                                                                        </div>
                                                                                    )
                                                                                })}
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
                                                                                        <div key={index}>
                                                                                            <input type="radio" name={value.internalName} value={ele}
                                                                                                // checked={this.state.Data[value.internalName] === ele ? true : false}
                                                                                                checked={(this.state.Data[value.internalName] !== null && this.state.Data[value.internalName].toString() === ele) ? true : false}
                                                                                                required={value.Required} disabled />
                                                                                            {ele === "true" ? "Yes" : "No"}

                                                                                        </div>
                                                                                    )
                                                                                })}
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
                                                                                        // titleText={value.Title}
                                                                                        personSelectionLimit={3}
                                                                                        groupName={""} // Leave this blank in case you want to filter from all users
                                                                                        showtooltip={true}
                                                                                        required={value.Required}
                                                                                        disabled={true}
                                                                                        defaultSelectedUsers={Object.keys(this.state.approversEmail).includes(value.internalName) ? this.state.approversEmail[value.internalName] : null}
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
                                                                                showHiddenInUI={false}
                                                                                principalTypes={[PrincipalType.User]}
                                                                                resolveDelay={1000}
                                                                                disabled={true} />
                                                                        </div>
                                                                    )
                                                                case "DateTime":
                                                                    return (
                                                                        <div className='DibRelfieldEditor' key={index}>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className={styles.viewfielddisply}>
                                                                                <label>{(this.state.Data[value.internalName] !== undefined) ? this.state.Data[value.internalName].slice(0, 10) : null}</label>
                                                                            </div>
                                                                        </div>

                                                                    )
                                                                default:
                                                                    return (
                                                                        <div className='DibRelfieldEditor' key={index}>
                                                                            <div>
                                                                                <label className='label' htmlFor={value.internalName}>{value.Title}{value.Required ? <span className='labelIcon'>*</span> : null}</label>
                                                                            </div>
                                                                            <div className={styles.viewfielddisply}>
                                                                                <label>{this.state.Data[value.internalName] || ""}</label>
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
                                    {((this.state.Data.StatusNo === "11000" || this.state.Data.StatusNo === "12000" || this.state.Data.StatusNo === "13000")) ?
                                        <BoardApprovers
                                            listName={this.props.listName}
                                            itemId={this._itemId}
                                            boardApprJson={this.state.boardApprvers}
                                            isAdmin={this.state.isAdminGroupUser}
                                            context={this.props.context}
                                            StatusNo={this.state.Data.StatusNo}
                                            loginUser={this._userEmail}
                                            Requster={this.state.RequsterEmail}
                                            userName={this._userName}
                                            userRole={this._UserRole}
                                            change={this.getItems}
                                        // OnClick={this.getItems()}

                                        /> : null}

                                </fieldset>
                            </PivotItem>
                            {/* Section 4 */}
                            <PivotItem headerText="Attachments">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className='fieldEditor'>
                                        <div>
                                            <label className='label' htmlFor='Attachments'>Uploaded Attachments</label>
                                            {/* <span className='labelIcon'>*</span> */}
                                        </div>

                                    </div>
                                    <div className={styles.viewfeildFilesDisplay}>
                                        <span>
                                            <ul style={{ margin: "unset", padding: "unset" }}>
                                                {this.state.GetAttchment.map((file, ind) =>
                                                    <div className={styles.viewfeildFilesDisplay} key={ind}>
                                                        <span className={styles.attachmentSpan}>
                                                            <a href={file.fileUrl} download={file.fileName}>
                                                                <IconButton
                                                                    iconProps={{ iconName: 'Download' }}
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
                                                            </a>
                                                        </span>
                                                        <span>
                                                            <li key={ind}>{file.fileName}</li>
                                                        </span>

                                                    </div>)}
                                            </ul>
                                        </span>
                                    </div>

                                </fieldset>
                            </PivotItem>
                            {/* section 5 */}
                            <PivotItem headerText="Approval">
                                <fieldset style={{ "border": "1px solid #80808030", "borderRadius": "10px", "padding": "20px" }}>
                                    <div className='conationer'>
                                        {this.state.approverFields.map((ele, ind) => {
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
                                                                        disabled={true}
                                                                        defaultSelectedUsers={this.state.approversEmail[ele.internalName]}
                                                                        //  onChange={this._getPeoplePickerItems.bind(this, ele.internalName)}
                                                                        showHiddenInUI={false}
                                                                        ensureUser={true}
                                                                        principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                                        resolveDelay={1000} />
                                                                </div>
                                                            </div>
                                                        </div>
                                                    }</>)
                                        })}

                                    </div>
                                </fieldset>
                            </PivotItem>
                            {/* section 6 */}
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
                            <PivotItem headerText="Comments">
                                <fieldset className='sonyfieldSet'>
                                    {/* <div className='fieldEditor'> */}
                                    <div>
                                        <div>
                                            <label className='label' htmlFor="Comments">Comments </label>
                                        </div>
                                        <div>
                                            <TextField
                                                //   required 
                                                multiline
                                                onChange={this.handleTextareaChange}
                                                value={this.state.Data.StatusNo === "10000" ? this.state.Data.Comments : this.state.Comments}
                                                disabled={
                                                    !(this.state.Data.StatusNo === "1000" && this.state.Approver1.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "2000" && this.state.Approver2.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "3000" && this.state.Approver3.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "4000" && this.state.Approver4.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "5000" && this.state.Approver5.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "6000" && this.state.Approver6.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "7000" && this.state.Approver7.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "8000" && this.state.Approver8.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "9000" && this.state.Approver9.includes(this._userEmail)

                                                    )
                                                }
                                            />
                                        </div>
                                    </div>
                                </fieldset>
                            </PivotItem>
                        </Pivot>
                        <div className='Spanbutton'>
                            <span style={{ marginRight: "1%" }} hidden={!((this.state.Data.StatusNo === "5" || this.state.Data.StatusNo === "15" || this.state.Data.StatusNo === "10000") && this.state.RequsterEmail === this._userEmail)}><PrimaryButton text={this.state.Data.StatusNo === "5" ? "Edit" : "Edit & Resubmit"} onClick={this.editPage} /></span>
                            <span style={{ marginRight: "1%" }}
                                hidden={
                                    !(this.state.Data.StatusNo === "1000" && this.state.Approver1.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "2000" && this.state.Approver2.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "3000" && this.state.Approver3.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "4000" && this.state.Approver4.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "5000" && this.state.Approver5.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "6000" && this.state.Approver6.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "7000" && this.state.Approver7.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "8000" && this.state.Approver8.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "9000" && this.state.Approver9.includes(this._userEmail)
                                    )
                                }><PrimaryButton text='Approve' onClick={this.apprFunctionality} /></span>
                            <span style={{ marginRight: "1%" }}
                                hidden={
                                    !(this.state.Data.StatusNo === "1000" && this.state.Approver1.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "2000" && this.state.Approver2.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "3000" && this.state.Approver3.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "4000" && this.state.Approver4.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "5000" && this.state.Approver5.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "6000" && this.state.Approver6.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "7000" && this.state.Approver7.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "8000" && this.state.Approver8.includes(this._userEmail) ||
                                        this.state.Data.StatusNo === "9000" && this.state.Approver9.includes(this._userEmail)
                                        // this.state.boardBtn === false
                                    )
                                }><PrimaryButton text='Reject' onClick={this.rejectBtn} /></span>
                            <span><PrimaryButton text='Exit' onClick={this.homePage} /></span>
                        </div>
                    </form >
                </div>
                <Dialog
                    hidden={this.state.hideDialog}
                    onDismiss={this.toggleHideDialog}
                    minWidth={300}
                    dialogContentProps={dialogContentProps}>
                    <div className={styles.dialogboxTextAlginments}>
                        {this.state.alertMsg.map(x => x.ApprMsg === true ? <p className={styles.Successmsg}>Request has been approved successfully</p> : null)}
                        {this.state.alertMsg.map(x => x.RjctMsg === true ? <p className={styles.Successmsg}>Request has been rejected successfully</p> : null)}
                    </div>
                    <DialogFooter>
                        <PrimaryButton text="Ok" onClick={this.onCancel} />
                    </DialogFooter>
                </Dialog>
                {/* commmnets d */}
                <Dialog
                    hidden={this.state.hideComentDialog}
                    onDismiss={this.commentstoggleHideDialog}
                    minWidth={300}

                    dialogContentProps={this.dialogcontent
                    }>
                    <div className={styles.dialogboxTextAlginments}>

                        <p style={{ fontSize: "12px", padding: "5px", textAlign: "center" }}>Please add comments</p>
                    </div>
                    <DialogFooter>
                        <PrimaryButton text="Ok" onClick={() => this.setState({ hideComentDialog: true })} />
                    </DialogFooter>
                </Dialog>
                {/* commnets popup */}
                <Dialog
                    hidden={this.state.hideComentTab}
                    onDismiss={this.commentstoggleHideDialogCMTTab}
                    minWidth={300}

                    dialogContentProps={this.dialogcontentComments
                    }>
                    <div className={styles.dialogboxTextAlginments}>
                    <TextField
                                                //   required 
                                                multiline
                                                onChange={this.handleTextareaChange}
                                                value={this.state.Data.StatusNo === "10000" ? this.state.Data.Comments : this.state.Comments}
                                                disabled={
                                                    !(this.state.Data.StatusNo === "1000" && this.state.Approver1.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "2000" && this.state.Approver2.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "3000" && this.state.Approver3.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "4000" && this.state.Approver4.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "5000" && this.state.Approver5.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "6000" && this.state.Approver6.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "7000" && this.state.Approver7.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "8000" && this.state.Approver8.includes(this._userEmail) ||
                                                        this.state.Data.StatusNo === "9000" && this.state.Approver9.includes(this._userEmail)

                                                    )
                                                }
                                            />

                    </div>
                    <DialogFooter>
                        <PrimaryButton text="Ok" onClick={this.commentsvalidation} />
                    </DialogFooter>
                </Dialog>
            </section >
        );

    }
}
