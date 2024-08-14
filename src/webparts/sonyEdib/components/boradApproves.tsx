import * as React from 'react';
import './custom.css'
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/files";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users/web";
import { IBoardAppr } from './Createform';
import { Dialog, DialogFooter, DialogType, IconButton, PrimaryButton } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import spService from './Serivce/spService';
import { IAuditTrail } from './EditForm';
import styles from './SonyEdib.module.scss';

export interface IBoaderApproversProps {
    listName: string;
    itemId: number;
    boardApprJson: IBoardAppr[];
    isAdmin: boolean;
    context: WebPartContext;
    StatusNo: string;
    loginUser: string;
    Requster: string;
    userName: string;
    userRole: string;
    change?: any;
}

export interface IBoaderApproversState {
    hideDialog: boolean;
    SelectedBoard: IBoardAppr[];
    boardApprerInfo: IBoardAppr[];
    getItemData: IBoardAppr[],
    SubmitUpdatedjson: IBoardAppr[];
    BoardName: string;
    AuditLog: IAuditTrail[];
    currentStatus: string;
    currentStatusNo: string;
    isUpdated: boolean;
    hideCommentsinput: boolean;
    hideCommentsinputPWB: boolean;
    CommentDialogProps: string;
    RejectCommentsData: any;
    PWBRejectedComments:any;
    hideAlertDialog: boolean;

}

// const iconName:IIconProps={iconName:"Edit"}
export default class BoardApprovers extends React.Component<IBoaderApproversProps, IBoaderApproversState, {}> {
    private _spService: spService = null;
    private _selected: { Board: string; Status: string, Id: string }[] = [];
    private _selectedPWB: { Board: string; Status: string, Id: string }[] = [];

    constructor(props: IBoaderApproversProps) {
        super(props);
        this.state = {
            hideDialog: true,
            SelectedBoard: [],
            boardApprerInfo: [],
            getItemData: [],
            SubmitUpdatedjson: [],
            BoardName: "",
            AuditLog: [],
            currentStatus: "",
            currentStatusNo: '',
            isUpdated: true,
            hideCommentsinput: true,
            CommentDialogProps: "",
            RejectCommentsData: {},
            hideAlertDialog: true,
            PWBRejectedComments:{},
            hideCommentsinputPWB:true
        }
    }

    // 
    public onEdit = async (Data: IBoardAppr, ind: number): Promise<any> => {
        this.setState({ BoardName: Data.Board });
        let temp: IBoardAppr[] = []
        await this._spService.getListItemsById(this.props.listName, this.props.itemId).then(res => {
            temp = JSON.parse(res.BoardApprovers)
        });

        // filter selected board 
        const fltrlatestJson = temp.filter(obj => obj.Board === Data.Board)
        const updatedArrayBoard = fltrlatestJson.map(item => {
            if (item.Board === Data.Board) {
                return {
                    ...item,
                    isUpdated: Data.isUpdated,
                    CCT: { Status: item.CCT.Status, isChanged: item.CCT.isChanged, value: item.CCT.Status === "CCT Changed" ? 1 : (item.CCT.Status === "DIB Revised" || item.CCT.Status === "DIB Cancelled") ? 2 : item.CCT.Status === "CCT Approved" ? 3 : item.CCT.Status === "CCT Rejected" ? 4 : 0, Comments: item.CCT.Comments },
                    PWB: { Status: item.PWB.Status, isChanged: item.PWB.isChanged, value: item.PWB.Status === "PWB Changed" ? 1 : item.PWB.Status === "PWB Approved" ? 3 : item.PWB.Status === "PWB Rejected" ? 4 : 0 }
                };
            }

            else {
                return item;
            }
        });
        this.setState({ SelectedBoard: updatedArrayBoard });
        this.toggleHideDailog();
    }

    // Dropdown Selection
    public onChengeBoardCCTStatus = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value, id } = eve.target;
        let upboard = [];
        const filterCCt = this._selected.filter(obj => obj.Id === id);
        // if board  is not there we need to push the board and board avlue
        if (value === "CCT Rejected") {
            this.setState({ CommentDialogProps: name });
            this.setState(prevState => ({ RejectCommentsData: { ...prevState.RejectCommentsData, [name]: "" } }));

            // temp.push(name);

            this.toggleCommentsHide();
        }
        if (!(this._selected.includes(filterCCt[0]))) {
            upboard.push(
                {
                    Board: name,
                    Status: value,
                    Id: id
                }
            );
        } else {
            // if board is  we neeed upadate value 
            upboard = this._selected.map(items => {
                if (items.Board === name && id.match(/CCT/g)) {
                    return {
                        ...items,
                        Status: value,
                        Id: id
                    }
                }

            })
        }
        this._selected = [...upboard];
        const updatedArrayBoard = this.state.SelectedBoard.map(item => {

            if (item.Board === name && id.match(/CCT/g)) {
                return {
                    ...item,
                    isUpdated: true,
                    CCT: { Status: value, isChanged: true, value: item.CCT.value, Comments: item.CCT.Comments }
                };
            }

            else {
                return item;
            }
        });

        this.setState({ SelectedBoard: updatedArrayBoard });
    }
    //PWB Status change Handler
    public onChengeBoardPWBStatus = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value, id } = eve.target;
        if (value === "PWB Rejected") {
            this.setState({ CommentDialogProps: name });
            this.setState(prevState => ({ PWBRejectedComments: { ...prevState.PWBRejectedComments, [name]: "" } }));

            // temp.push(name);

            this.toggleCommentsHidePWB();
        }
        let upboard = [];
        const filterCCt = this._selectedPWB.filter(obj => obj.Id === id);
        if (!(this._selectedPWB.includes(filterCCt[0]))) {
            upboard.push(
                {
                    Board: name,
                    Status: value,
                    Id: id
                }

            )
        } else {
            upboard = this._selectedPWB.map(items => {
                if (items.Board === name && id.match(/PWB/g)) {
                    return {
                        ...items,
                        Status: value,
                        Id: id
                    }
                }
            });
        }
        this._selectedPWB = [...upboard];
        const updatedArrayBoard = this.state.SelectedBoard.map(item => {
            if (item.Board === name && id.match(/PWB/g)) {
                return {
                    ...item,
                    isUpdated: true,
                    PWB: { Status: value, isChanged: true, value: item.PWB.value,Comments:item.PWB.Comments }
                };
            }
            else {
                return item;
            }
        });

        this.setState({ SelectedBoard: [...updatedArrayBoard] });
    }
    // onsubmit
    public onCancel = async (): Promise<void> => {
        this.setState({ SelectedBoard: [], hideDialog: true });
        this._selected.length = 0;
        this._selectedPWB.length = 0;
    }
    //CCT reject comments validation
    public rejectedCommentsvaildation = (board: string): void => {
        if (this.state.RejectCommentsData[board] !== "") {
            this.setState({ hideCommentsinput: true });
        }
        else {
            this.toggleAlertMsg();
            this.setState({ hideCommentsinput: false });

        }
    }
    // PWB Rejected
    public PWBrejectedCommentsvaildation = (board: string): void => {
        if (this.state.PWBRejectedComments[board] !== "") {
            this.setState({ hideCommentsinputPWB: true });
        }
        else {
            this.toggleAlertMsg();
            this.setState({ hideCommentsinputPWB: false });

        }
    }
    // onSubmit dropdwn values PWBrejectedCommentsvaildation
    public onSubmit = async (): Promise<void> => {
        await this._spService.getListItemsById(this.props.listName, this.props.itemId).then(res => {
            this.setState({
                getItemData: JSON.parse(res.BoardApprovers),
                AuditLog: JSON.parse(res.AuditTrail),
                currentStatus: res.Status,
                currentStatusNo: res.StatusNo
            });
        });
        const auditlog = this.state.AuditLog;
        const SelectedValues = [...this._selected, ...this._selectedPWB];
        let updatedArrayBoard: IBoardAppr[];
        // if(SelectedValues[0].Status === "CCT  Rejected" || SelectedValues[1].Status === "PWB Rejected"){}
        for (let i = 0; i < SelectedValues.length; i++) {
            updatedArrayBoard = this.state.getItemData.map(item => {
                if (item.Board === SelectedValues[i].Board && SelectedValues[i].Id.match(/CCT/g)) {
                    const auditObj = {
                        "Actioner": this.props.userName,
                        "ActionTaken": `${SelectedValues[i].Board} CCT desgin status changed to ${SelectedValues[i].Status}`,
                        "Role": this.props.userRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments": SelectedValues[i].Status === "CCT Rejected" ? this.state.RejectCommentsData[SelectedValues[i].Board] : ""
                    };
                    auditlog.push(auditObj);
                    // value: SelectedValues[i].Status  === "CCT Changed" ? 1 :(SelectedValues[i].Status  === "DIB Revised" ||SelectedValues[i].Status  === "DIB Cancelled")?2:(SelectedValues[i].Status  === "CCT Approved"||SelectedValues[i].Status  === "CCT Rejected")?3:0
                    return {
                        ...item,
                        isUpdated: true,
                        CCT: { Status: SelectedValues[i].Status, isChanged: true, value: SelectedValues[i].Status === "CCT Changed" ? 1 : (SelectedValues[i].Status === "DIB Revised" || SelectedValues[i].Status === "DIB Cancelled") ? 2 : SelectedValues[i].Status === "CCT Approved" ? 3 : SelectedValues[i].Status === "CCT Rejected" ? 4 : 0, Comments: SelectedValues[i].Status === "CCT Rejected" ? this.state.RejectCommentsData[SelectedValues[i].Board] : "" }
                    };
                }
                else if (item.Board === SelectedValues[i].Board && SelectedValues[i].Id.match(/PWB/g)) {
                    const auditObj = {
                        "Actioner": this.props.userName,
                        "ActionTaken": `${SelectedValues[i].Board} PWB desgin status changed to ${SelectedValues[i].Status}`,
                        "Role": this.props.userRole,
                        "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
                        "Comments":SelectedValues[i].Status === "PWB Rejected" ? this.state.PWBRejectedComments[SelectedValues[i].Board] : ""
                    };
                    auditlog.push(auditObj);
                    return {
                        ...item,
                        isUpdated: true,
                        PWB: { Status: SelectedValues[i].Status, isChanged: true, value: SelectedValues[i].Status === "PWB Changed" ? 1 : SelectedValues[i].Status === "PWB Approved" ? 3 : SelectedValues[i].Status === "PWB Rejected" ? 4 : 0 ,Comments:SelectedValues[i].Status === "PWB Rejected" ?this.state.PWBRejectedComments[SelectedValues[i].Board] :""}
                    };
                }
                else {
                    return item;
                }
            });
            this.setState({ getItemData: updatedArrayBoard });

        }
        const updatedIsComleted = updatedArrayBoard.map(obj => {
            if (!(obj.CCT.Status === "CCT Changed" || obj.CCT.Status === "" || obj.CCT.Status === "CCT Rejected") && !(obj.PWB.Status === "PWB Changed" || obj.PWB.Status === "PWB Rejected" || obj.PWB.Status === "")) {

                return {
                    ...obj,
                    isCompleted: true
                }
            }
            else {
                return obj

            }
        });

        // checking is completed
        // const allBoardIsComplete = updatedIsComleted.every(x => x.isCompleted === true);
        // if (allBoardIsComplete) {
        //     const auditObj = {
        //         "Actioner": this.props.userName,
        //         "ActionTaken": "Board Approval Completed",
        //         "Role": this.props.userRole,
        //         "ActionTakenOn": new Date().toDateString() + " " + new Date().toLocaleTimeString(),
        //         "Comments": ""
        //     };
        //     auditlog.push(auditObj);

        //     this.setState({
        //         currentStatus: "Board Approval Completed",
        //         currentStatusNo: "13000"
        //     })
        // }

        const updatedItem = {
            BoardApprovers: JSON.stringify(updatedIsComleted),
            AuditTrail: JSON.stringify(auditlog),
            StartProcessing: true,
            Status: this.state.currentStatus,
            StatusNo: this.state.currentStatusNo
        }
        await this._spService.updateListItem(this.props.listName, updatedItem, this.props.itemId).then(async res => {
            // console.log(res.data.Title)
            this.setState({ hideDialog: true, boardApprerInfo: updatedIsComleted });
            this._selected.length = 0;
            this._selectedPWB.length = 0;
            // getItems();
            this.props.change();
            return res;
        }).catch(err => err);
        this.setState({ hideDialog: true });
    }

    public homePage = (): void => {
        const pageURL: string = this.props.context.pageContext.web.absoluteUrl;
        // console.log(pageURL);
        // window.location.href = pageURL + "/SitePages/Sony.aspx";
        window.location.href = pageURL;
        this.setState({ hideDialog: true });
    }
    public toggleHideDailog = (): void => {
        this.setState({ hideDialog: false })
    }
    public toggleCommentsHide = (): void => {
        this.setState({ hideCommentsinput: false });
    }
    public toggleCommentsHidePWB = (): void => {
        this.setState({ hideCommentsinputPWB: false });
    }
    public toggleAlertMsg = (): void => {
        this.setState({ hideAlertDialog: false })
    }
    public handleTextareaChange = (event: React.ChangeEvent<HTMLTextAreaElement>, ind?: number): void => {
        const { name, value } = event.target;
        this.setState(prevState => ({ RejectCommentsData: { ...prevState.RejectCommentsData, [name]: value } }));
        // console.log(this.state.RejectCommentsData, this.state.CommentDialogProps);
    }
    public handleTextareaChangePWBRehected = (event: React.ChangeEvent<HTMLTextAreaElement>, ind?: number): void => {
        const { name, value } = event.target;
        this.setState(prevState => ({ PWBRejectedComments: { ...prevState.PWBRejectedComments, [name]: value } }));
        // console.log(this.state.PWBRejectedComments, this.state.CommentDialogProps);
    }
    // public 
    public componentDidMount(): void {
        this._spService = new spService(this.props.context);
        if (this.state.isUpdated === true) {
            this.setState({ boardApprerInfo: this.props.boardApprJson });
        }
    }

    public render(): React.ReactElement<IBoaderApproversProps> {
        const dialogContentProps = {
            type: DialogType.normal,
            title: `${this.state.BoardName} Desgin status`

        }
        const dialogContentPropsComments = {
            type: DialogType.normal,
            title: "Comments"

        }
        const dialogContentPropsalrt = {
            type: DialogType.normal,
            title: "Information!"

        }
        // Information!
        return (
            <div>
                {this.state.boardApprerInfo.map((x, ind) => {
                    return (
                        <><div className='boardApprConatiner'>
                            <div className='boardapproverlabel'>
                                <span style={{ width: "50px" }}>{x.Board}</span>
                                <span hidden={!(this.props.StatusNo === "12000" && x.isCompleted === false && (this.props.isAdmin === true || this.props.Requster === this.props.loginUser))} style={{ paddingLeft: "20px" }}>
                                    <IconButton onClick={() => this.onEdit(x, ind)}
                                        title='Modify Board status'
                                        iconProps={{ iconName: 'SingleColumnEdit' }}
                                        styles={{
                                            icon: { fontSize: 18 },
                                            root: {

                                            },
                                            rootHovered: { backgroundColor: "white" },
                                            rootPressed: { backgroundColor: 'white' }
                                        }}
                                    />

                                </span>
                            </div>
                        </div>
                            <div className='borderStatusContainer'>
                                <div className='boardStatus'>
                                    <div className='boardApprConatinerlabel'>
                                        <label className='boardapproverlabel'>CCT Status</label><label className={styles.boardStatusLabel}>{x.CCT.Status}</label>
                                    </div>
                                    <div className='boardApprConatinerlabel'>
                                        <label className='boardapproverlabel'>PWB Status</label><label className={styles.boardStatusLabel}>{x.PWB.Status}</label>
                                    </div>
                                </div>
                            </div>
                        </>
                    )
                })}
                <Dialog
                    hidden={this.state.hideDialog}
                    onDismiss={this.toggleHideDailog}
                    minWidth={500}
                    dialogContentProps={dialogContentProps}>
                    <div>
                        {
                            this.state.SelectedBoard.map((obj, ind) => {
                                return (<>
                                    <div className='dailogboxfieldEditor' key={ind}>
                                        <div>
                                            <label className='dailogboxlabel' htmlFor={obj.Board}>CCT</label>
                                        </div>
                                        <div style={{ width: "100%" }}>
                                            <select className='dropdwon'
                                                id={`${obj.Board}CCT`}
                                                name={obj.Board}
                                                value={obj.CCT.Status}
                                                onChange={(event) => this.onChengeBoardCCTStatus(event, ind)}
                                                // disabled={obj.isCompleted === true?true:false}
                                                disabled={!(obj.CCT.value === 0 || obj.CCT.value === 1 || obj.CCT.value === 4)}
                                            >
                                                <option disabled />
                                                <option disabled={!(this.props.isAdmin && (obj.CCT.value === 0 || obj.CCT.value === 4))}>CCT Changed</option>
                                                <option disabled={!(this.props.Requster === this.props.loginUser && obj.CCT.value === 1)}> CCT Approved </option>
                                                <option disabled={!(this.props.Requster === this.props.loginUser && obj.CCT.value === 1)}> CCT Rejected</option>
                                                <option disabled={!((obj.CCT.value === 0 || obj.CCT.value === 4) && this.props.isAdmin)}>DIB Revised</option>
                                                <option disabled={!((obj.CCT.value === 0 || obj.CCT.value === 4) && this.props.isAdmin)}>DIB Cancelled</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div className='dailogboxfieldEditor' key={ind}>
                                        <div>
                                            <label className='dailogboxlabel' htmlFor={obj.Board}>PWB</label>
                                        </div>
                                        <div style={{ width: "100%" }}>
                                            <select className='dropdwon'
                                                id={`${obj.Board}PWB`}
                                                name={obj.Board}
                                                value={obj.PWB.Status}
                                                onChange={(event) => this.onChengeBoardPWBStatus(event, ind)}
                                                disabled={!(obj.PWB.value === 1 || obj.PWB.value === 0 || obj.PWB.value === 4)}
                                            >
                                                <option disabled />
                                                <option disabled={!(this.props.isAdmin && (obj.PWB.value === 0 || obj.PWB.value === 4))}>PWB Changed</option>
                                                <option disabled={!(this.props.Requster === this.props.loginUser && obj.PWB.value === 1)}> PWB Approved </option>
                                                <option disabled={!(this.props.Requster === this.props.loginUser && obj.PWB.value === 1)}> PWB Rejected</option>
                                            </select>
                                        </div>
                                    </div>
                                </>
                                )
                            })
                        }
                    </div>
                    <DialogFooter>
                        <span hidden={!(this._selected.length > 0 || this._selectedPWB.length > 0)}>
                            <PrimaryButton text="Submit" onClick={this.onSubmit} />
                        </span>

                        <PrimaryButton text="Cancel" onClick={this.onCancel} />
                    </DialogFooter>
                </Dialog>
                {/*  CCT Rejecct comments dailong */}
                <Dialog
                    hidden={this.state.hideCommentsinput}
                    onDismiss={this.toggleCommentsHide}
                    minWidth={500}
                    dialogContentProps={dialogContentPropsComments}>
                    <textarea
                        name={this.state.CommentDialogProps}
                        id={this.state.CommentDialogProps + "CCTRejected"}
                        value={this.state.RejectCommentsData[this.state.CommentDialogProps] || ""}
                        // style={{ border: "none", outline: "none", resize: "none" }}
                        required
                        // readOnly
                        rows={3}
                        className="textarea"

                        onChange={(event) => this.handleTextareaChange(event)}
                    />

                    <DialogFooter>
                        <span>
                            <PrimaryButton text="Ok" onClick={() => this.rejectedCommentsvaildation(this.state.CommentDialogProps)} />
                        </span>

                        {/* <PrimaryButton text="Cancel" onClick={this.onCancel} /> */}
                    </DialogFooter>
                </Dialog>
                 {/*  PWb Rejecct comments dailong */}
                 <Dialog
                    hidden={this.state.hideCommentsinputPWB}
                    onDismiss={this.toggleCommentsHidePWB}
                    minWidth={500}
                    dialogContentProps={dialogContentPropsComments}>
                    <textarea
                        name={this.state.CommentDialogProps}
                        id={this.state.CommentDialogProps + "PWBRejected"}
                        value={this.state.PWBRejectedComments[this.state.CommentDialogProps] || ""}
                        // style={{ border: "none", outline: "none", resize: "none" }}
                        required
                        // readOnly
                        rows={3}
                        className="textarea"

                        onChange={(event) => this.handleTextareaChangePWBRehected(event)}
                    />

                    <DialogFooter>
                        <span>
                            <PrimaryButton text="Ok" onClick={() => this.PWBrejectedCommentsvaildation(this.state.CommentDialogProps)} />
                        </span>

                        {/* <PrimaryButton text="Cancel" onClick={this.onCancel} /> */}
                    </DialogFooter>
                </Dialog>
                {/* alert comments */}
                <Dialog
                    hidden={this.state.hideAlertDialog}
                    onDismiss={this.toggleAlertMsg}
                    minWidth={500}
                    dialogContentProps={dialogContentPropsalrt}>
                    <div className={styles.dialogboxTextAlginments}>

                        <p style={{ fontSize: "12px", padding: "5px", textAlign: "center" }}>Please add comments</p>
                    </div>

                    <DialogFooter>
                        <span>
                            <PrimaryButton text="Ok" onClick={() => this.setState({ hideAlertDialog: true })} />
                        </span>

                        {/* <PrimaryButton text="Cancel" onClick={this.onCancel} /> */}
                    </DialogFooter>
                </Dialog>

            </div>

        );

    }
}
