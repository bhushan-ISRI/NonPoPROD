import React, { useEffect, useState } from "react";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import { useParams } from "react-router-dom";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import SPCRUDOPS from "../../service/DAL/newspcrudops";
import "./CSS/NewRequest.scss";
import { IExpenseMaster } from "../../service/INTERFACE/IExpenseMaster";

import logo from "../../assets/sona-comstarlogo.png";

import { TextField, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { useNavigate } from 'react-router-dom';

interface IRouteParams extends Record<string, string | undefined> {
    id?: string;
}

export const APTeamAcceptance = (props: ISonanonpoprodProps) => {
    const navigate = useNavigate();

    const { id } = useParams<IRouteParams>();
    const [RequestDate, setRequestDate] = React.useState<string>();
    const [EmployeeName, setEmployeeName] = React.useState<string>();
    const [Divison, setDivison] = React.useState<string>();
    const [Email, setEmail] = React.useState<string>();
    const [HOD, setHOD] = React.useState<string>();
    const [Department, setDepartment] = React.useState<string>();
    const [EmpStatus, setEmpStatus] = React.useState<string>();
    const [Status, setStatus] = React.useState<string>();
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [ApprovalMatrixdata, setApprovalMatrix] = React.useState<any[]>([]);
    const [WorkflowHistory, setWorkflowHistory] = React.useState<any[]>([]);
    const [GLDescription, setGLDescription] =React.useState<any | null>(null);
    const [CostCenterDescription, setCostCenterDescription] =React.useState<any | null>(null);
    const [NatureOfExpItems, setNatureOfExpItems] =React.useState<any | null>(null);
//const [NatureOfExpItems, setNatureOfExpItems] = React.useState<any[]>([]);
    
    const [ItemID, setItemID] = useState<any>();
    const [Comment, setComment] = React.useState("");

    const [GLCodeItems, setGLCodeItems] = React.useState<IExpenseMaster[]>([]);
     const [masterItems, setmasterItems] = React.useState<any[]>([]);
    
    const [GLCodeOptions, setGLCodeOptions] = React.useState<IDropdownOption[]>([]);
    //const [selectedGLCode, setSelectedGLCode] =
    const [selectedGLCode, setSelectedGLCode] =
    React.useState<string>("");
   // const [selectedGLCode, setSelectedGLCode] = React.useState<string | number>();


    

    const [CostCenterOptions, setCostCenterOptions] = React.useState<IDropdownOption[]>([]);
    //const [selectedCostCenter, setSelectedCostCenter] = React.useState<string | number>();
    const [selectedCostCenter, setSelectedCostCenter] = useState<string | undefined>();

    const [WorkflowJSX, setWorkflowJSX] = React.useState<JSX.Element[] | null>(null);
    const [Stage, setStageData] = React.useState(0);
    //const [GLCodeValue, setGLCodeValue] = React.useState<any[]>([]);
    const [GLCodeValue, setGLCodeValue] = React.useState<string>("");
    const [CostCenterValue, setCostCenterValue] = React.useState<string>("");

    const [ExpenseTypemMastervalues, setExpenseTypemMastervalues] = React.useState<any[]>([]);

    const [requestNo, setrequestNo] = useState<any>();// useState<string | undefined>(undefined);

    const [contactNo, setcontactNo] = React.useState<string>();
    const [Location, setLocation] = React.useState<string>();
    const [RMVal, setRMVal] = React.useState<string>("");
    const [VendorCode, setVendorCode] = React.useState<string>("");
    const [VendorName, setVendorName] = React.useState<string>("");
    const [NatureExpense, setNatureExpense] = React.useState<string>("");
    const [InvoiceNo, setInvoiceNo] = React.useState<string>("");
    const [InvoiceDate, setInvoiceDate] = React.useState<string>("");
    const [BasicAmount, setBasicAmount] = React.useState<string>("");

    const [PayRemark, setPayRemark] = React.useState<string>("");
    const [GSTAmt, setGSTAmt] = React.useState<string>("");
    const [otherCharges, setotherCharges] = React.useState<string>("");
    const [TotalAmount, setTotalAmount] = React.useState<string>("");
    const [existingFiles, setExistingFiles] = useState<IExistingFile[]>([]);
    const [loadingExisting, setLoadingExisting] = useState(false);



    // useEffect(() => {
    //     if (id) {
    //         loadItem(Number(id));
    //     }
    // }, [id]);



    React.useEffect(() => {
        if (ApprovalMatrixdata.length > 0) {
            displayWorkflow();
        }
    }, [ApprovalMatrixdata]);
    const displayWorkflow = (matrix?: any[]) => {
        const data = matrix || ApprovalMatrixdata;

        const wf: JSX.Element[] = [];
        let isActive;
        let notActive = false;

        data.forEach((m, i) => {
            if (!notActive && Stage !== 99) {
                if (Stage === i) {
                    isActive = 'activeApprover';
                    notActive = true;
                } else {
                    isActive = 'beforeactiveApprover';
                }
            } else {
                isActive = 'overrideStage';
            }

            wf.push(
                <ul className="main-menu" key={i}>
                    <li className={`${m.Role} ${isActive}`.trim()}>
                        {m.User}
                    </li>
                </ul>
            );
        });

        setWorkflowJSX(wf);
    };
    async function getData(): Promise<any[]> {

    const sp = spfi().using(SPFx(props.currentSPContext));

    const items: any[] = await sp.web.lists
        .getByTitle("ExpenseMaster")
        .items
        .select(
            "ID",
            "Title",
            "GLCode",
            "CostCenter",
            "Status",
            "ExpenseType",
            "GLDescription"
        )
        .top(100)();

    console.log(items);

    setExpenseTypemMastervalues(items);

    // IMPORTANT
    return items;
}

    // useEffect(() => {
    //     getData();
    // }, []);
    useEffect(() => {
    loadInitialData();
}, []);

const loadInitialData = async () => {
  const expenseMasterData = await getData();
    //await getData();
   if (id) {
     await loadItem(Number(id), expenseMasterData);
       // await loadItem(Number(id));
    }

    
};

    
    const th: React.CSSProperties = {
        textAlign: "left",
        padding: "8px 10px",
        borderBottom: "1px solid #e5e7eb",
        fontWeight: 700,
        fontSize: 13,
    };
    const td: React.CSSProperties = {
        padding: "8px 10px",
        borderBottom: "1px solid #e5e7eb",
        fontSize: 13,
    };
    interface IExistingFile {
        name: string;
        url: string;
        sizeBytes: number;
        lastModified?: string;
        source: "Library" | "ListAttachment";
        serverRelativeUrl?: string;
        attachmentFileName?: string;
    }
    const loadItem = async (
    itemId: number,
    expenseMasterData: any[]
) => {

    try {

        const sp = spfi().using(SPFx(props.currentSPContext));

        const item = await sp.web.lists
            .getByTitle("NonPO")
            .items.getById(itemId)
            .select("ID,Title,GLDescription,CostCenterDescription,RequestDate,Stage,EmployeeName,ApprovalMatrix,WorkflowHistory,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus,ApprovalMatrix,VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,WorkflowHistory,Status,PaymentRemarks,ApprovalRemarks,VendorCode/VendorCode,VendorCode/Id,VendorCode/Title,NatureofExpense/Title,NatureofExpense/ExpenseType,CurrentApprover/ID,CurrentApprover/EMail,CurrentApprover/Title")
            .expand("VendorCode,NatureofExpense,CurrentApprover")();

        console.log(item);
        setItemID(itemId);
            setrequestNo(item.Title || "")
            setPayRemark(item.PaymentRemarks || "")
            setStageData(item.Stage || "");

            setStatus(item.Status || "")
            setRMVal(item.RM ?? "");
            setLocation(item.Location || "");
            setRequestDate(item.RequestDate ? new Date(item.RequestDate).toLocaleDateString("en-GB") : "");
            setEmployeeName(item.EmployeeName || "");
            setDivison(item.Division || "");
            setEmail(item.Email || "");
            setHOD(item.HOD || "");
            setDepartment(item.Department || "");
            setEmpStatus(item.EmployeeStatus || "");
            setcontactNo(item.ContactNo || "");
            setVendorCode(item.VendorCode.VendorCode || "");
            setVendorName(item.VendorName || "");
            const Empitems = await sp.web.lists
                .getByTitle("EmployeeMaster")
                .items
                .filter(`EmployeeName eq '${item.EmployeeName}'`)
                .select("*,Employee/Title,Employee/Id")
                .expand("Employee")();
            const emp = Empitems[0];
            const EmpId = emp?.Employee?.Id || "";
            setCostCenterValue(Empitems[0].CostCenter);
            setCostCenterDescription(Empitems[0].CostCenterDescription)

            setInvoiceNo(item.InvoiceNumber || "");
            const invDate = new Date(item.InvoiceDate || "").toISOString().split("T")[0];
            setInvoiceDate(invDate);
            setBasicAmount(item.Basicamount || "");
            setGSTAmt(item.GST || "");
            setotherCharges(item.OtherCharges || "");
            setTotalAmount(item.Totalamount || "");
            let parsedApprovalMatrix = [];
            if (item.ApprovalMatrix) {
                try {
                    parsedApprovalMatrix = JSON.parse(item.ApprovalMatrix);
                } catch (parseError) {
                    console.error("Error parsing ApprovalMatrix JSON:", parseError);
                }
            }

            setApprovalMatrix(parsedApprovalMatrix);
            await loadFinalNonPoFile(item.Title);
            let parsedWorkflowHistory = [];
            if (item.WorkflowHistory) {
                try {
                    parsedWorkflowHistory = JSON.parse(item.WorkflowHistory);
                } catch (parseError) {
                    console.error("Error parsing WorkflowHistory JSON:", parseError);
                }
            }

            setWorkflowHistory(parsedWorkflowHistory);
            // GL code :dropdown
         setNatureExpense(
            item.NatureofExpense?.ExpenseType || ""
        );

        // FILTER HERE
        let filteredGLCode = expenseMasterData.filter(
            (m) =>
                m.ExpenseType ===
                item.NatureofExpense?.ExpenseType
        );

        console.log("FILTERED", filteredGLCode);

        // Dropdown options]
        const options: IDropdownOption[] =
            expenseMasterData.map((x) => ({
                key: x.GLCode,
                text: x.GLCode
            }));

        setGLCodeOptions(options);
        // // Bind selected value
         setSelectedGLCode(filteredGLCode[0].GLCode);
         setGLDescription(filteredGLCode[0].GLDescription);

    } catch (error) {

        console.log(error);
    }
};
    const toFolderName = (requestNo: string) => {
        return (requestNo || "NONPO")
            .replace(/[\\#%&*:{?<>|"]/g, "_") // removed /
            .replace(/\//g, "-")
            .trim();
    };
    const loadFinalNonPoFile = async (requestNo: any) => {
        //const absUrl: string = currentSPContext?.pageContext?.web?.absoluteUrl || "";
        try {
            const sp = spfi().using(SPFx(props.currentSPContext));
            const origin = window.location.origin;

            const absUrl: string = props.currentSPContext?.pageContext?.web?.absoluteUrl || "";

            // if (!props.currentSPContext || !requestNo) return;
            setLoadingExisting(true);

            const libTitle = "NonPODoc";

            // 1) Library files under /<RequestNo>/
            const libFiles: IExistingFile[] = [];
            try {
                const root = await sp.web.lists.getByTitle(libTitle).rootFolder();
                const rootUrl: string = root.ServerRelativeUrl;
                const folderPath = `${rootUrl}/${toFolderName(requestNo)}`;

                const files = await sp.web
                    .getFolderByServerRelativePath(folderPath)
                    .files.select("Name,ServerRelativeUrl,Length,TimeLastModified")();

                for (const f of files || []) {
                    libFiles.push({
                        name: f.Name,
                        url: `${origin}${f.ServerRelativeUrl}`,
                        sizeBytes: Number(f.Length || 0),
                        lastModified: f.TimeLastModified,
                        source: "Library",
                        serverRelativeUrl: f.ServerRelativeUrl,
                    });
                }
            } catch (error) {
                alert("error")
            }

            setExistingFiles([...libFiles]);
        } catch (e) {
            console.warn("[APTeamAcceptance] existing docs load failed:", e);
            setExistingFiles([]);
        } finally {
            setLoadingExisting(false);
        }

    };
    const handleSendback = async () => {
        try {

            if (!Comment || Comment.trim() === "") {
                alert("Please enter a comment before sending back.");
                return;
            }

            setIsSubmitting(true);

            if (!id) return;
            const itemId = Number(id);
            const spCrudOps = await SPCRUDOPS(props);
            const sp = spfi().using(SPFx(props.currentSPContext));
            const currentUser = await sp.web.currentUser();

            const item = await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .select("*")();



            let approvalMatrix: any[] = [];
            approvalMatrix = JSON.parse(item.ApprovalMatrix || "[]");

            const updatedIndex = approvalMatrix.findIndex(x =>
                parseInt(x.UserID) === currentUser.Id && x.Role === "APPerformer");

            const currentItem = approvalMatrix.find(x =>
                String(x.UserID) === String(currentUser.Id) && x.Role === "APPerformer"
            );

            const currentRole = currentItem.Role;
            const currentLevel = currentItem.Level;
            const normalizeId = (id: any) => parseInt(id);
            // const nowed = new Date();
            const formattedDateed = new Date().toISOString();
            if (currentItem !== -1) {
                approvalMatrix[updatedIndex] = {
                    ...approvalMatrix[updatedIndex],
                    Status: "Send Back",
                    Comment: Comment,
                    ActionDate: formattedDateed
                };
            }
            // =====================================
            // ✅ STEP 5: FINAL JSON
            // =====================================

            const updatedApprovalMatrix = JSON.stringify(approvalMatrix);

            const now = new Date();
            const formattedDate = new Date().toISOString();

            // ===============================
            // Use existing workflow history
            // ===============================

            let workflowHistory = [...WorkflowHistory];

            // ===============================
            // Add new history entry
            // ===============================

            const newHistoryEntry = {
                CurrentApprover: currentUser.Title,
                ActionTaken: `${currentRole} Send Back`,
                Comment: Comment || "",
                Date: formattedDate,
                CurrentStatus: 'Send Back'
            };

            workflowHistory.push(newHistoryEntry);

            setWorkflowHistory(workflowHistory);

            const updatedHistory = JSON.stringify(workflowHistory);

            const Empitems = await sp.web.lists
                .getByTitle("EmployeeMaster")
                .items
                .filter(`EmployeeName eq '${item.EmployeeName}'`)
                .select("*,Employee/Title,Employee/Id")
                .expand("Employee")();
            const emp = Empitems[0];
            const EmpId = emp?.Employee?.Id || "";



            await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(Number(id))
                .update({
                    Status: "Send Back",
                    PendingAt: 'Pending At Initiator',
                    CurrentApproverId: EmpId,
                    WorkflowHistory: updatedHistory,
                    ApprovalMatrix: updatedApprovalMatrix,
                    CostCentre: CostCenterValue,
                    GLCode: GLCodeValue
                });

            alert("AP team SendBack successfully.");
            setComment("");
            navigate('/APTeam');


        } catch (error) {

            console.error(error);
            alert("Something went wrong.");

        } finally {

            setIsSubmitting(false);

        }

    };
    const handleApprove = async () => {

        if (isSubmitting) return;
        console.log(selectedGLCode);
        if (!selectedGLCode) {
            alert("Please select the GL code");
            return;
        }
       console.log( NatureOfExpItems)

        if (!CostCenterValue) {
            alert("Please select the Cost center");
            return;
        }
        if (!Comment || Comment.trim() === "") {
                alert("Please enter a comment");
                return;
        }

        //  setIsSubmitting(true);

        try {

            const itemId = Number(id);

            const spCrudOps = await SPCRUDOPS(props);
            const sp = spfi().using(SPFx(props.currentSPContext));
            //  const currentUser = await sp.web.currentUser();
            const currentUser = await sp.web.currentUser();
            const item = await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .select("*")();

            let approvalMatrix: any[] = [];
            approvalMatrix = JSON.parse(item.ApprovalMatrix || "[]");
            // =====================================
            // FIND CURRENT APPROVER Index
            // =====================================

            const updatedIndex1 = approvalMatrix.findIndex(x =>
                parseInt(x.UserID) === currentUser.Id && x.Role === "APPerformer");

             const currentItem = approvalMatrix.find(x =>
                parseInt(x.UserID) === (currentUser.Id) && x.Role === "APPerformer"
            );

            const currentRole = currentItem.Role;
            const currentLevel = currentItem.Level;

            
         
            const normalizeId = (id: any) => parseInt(id);


            // const nowed = new Date();
            const formattedDateed = new Date().toISOString();
            if (currentItem !== -1) {
                approvalMatrix[updatedIndex1] = {
                    ...approvalMatrix[updatedIndex1],
                    Status: "Pending for Vouching",
                    Comment: Comment,
                    ActionDate: formattedDateed
                };
            }
            // =====================================
            // ✅ STEP 5: FINAL JSON
            // =====================================

            const updatedApprovalMatrix = JSON.stringify(approvalMatrix);

            const now = new Date();
            const formattedDate = new Date().toISOString();

            // ===============================
            // Use existing workflow history
            // ===============================

            let workflowHistory = [...WorkflowHistory];

            // ===============================
            // Add new history entry
            // ===============================

            const newHistoryEntry = {
                CurrentApprover: currentUser.Title,
                ActionTaken: `${currentRole} Approved`,
                Comment: Comment || "",
                Date: formattedDate,
                CurrentStatus: 'Pending for Vouching'
            };

            workflowHistory.push(newHistoryEntry);

            setWorkflowHistory(workflowHistory);

            const updatedHistory = JSON.stringify(workflowHistory);

            // =====================================
            // PREPARE UPDATE FIELDS
            // =====================================

            let updateFields: any = {};
            updateFields = {

                Status: "Pending for Vouching",
                PendingAt: currentUser.Title,
                ApproverStatus: "Pending for Vouching",
                //Stage: updatedIndex + 1,
                CurrentApproverId: currentUser.Id,
                WorkflowHistory: updatedHistory,
                ApprovalMatrix: updatedApprovalMatrix,
                CostCentre: CostCenterValue,
                GLCode:selectedGLCode,// GLCodeValue,
                NatureofExpenseId:NatureOfExpItems.Id,
                GLDescription:GLDescription,
                CostCenterDescription:CostCenterDescription
            };


            // =====================================
            // UPDATE SHAREPOINT ITEM
            // =====================================

            await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .update(updateFields);


            // =====================================
            // SUCCESS
            // =====================================

            alert(`Request successfully approved by the ${currentRole}`);
            setComment("");
            navigate('/APTeam');


        }
        catch (error) {
            console.error(error);
            alert("Approval failed");
        }
        finally {
            setIsSubmitting(false);
        }

    };
    const handleReject = async () => {
        try {

            if (!Comment || Comment.trim() === "") {
                alert("Please enter a comment before Reject.");
                return;
            }

            setIsSubmitting(true);

            const itemId = Number(id);
            const sp = spfi().using(SPFx(props.currentSPContext));
            const currentUser = await sp.web.currentUser();

            const item = await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .select("*")();

            // =====================================
            // ✅ Parse Approval Matrix
            // =====================================

            let approvalMatrix: any[] = [];
            try {
                approvalMatrix = JSON.parse(item.ApprovalMatrix || "[]");
            } catch {
                alert("Invalid Approval Matrix");
                return;
            }

            // =====================================
            // ✅ Find Current Approver
            // =====================================

            const currentIndex = approvalMatrix.findIndex(x =>
                parseInt(x.UserID) === currentUser.Id
            );

            if (currentIndex === -1) {
                alert("You are not authorized");
                return;
            }

            const currentRole = approvalMatrix[currentIndex].Role;


            // =====================================
            // ✅ UPDATE APPROVAL MATRIX (Mark Rejected)
            // =====================================

            const formattedDate = new Date().toISOString();

            approvalMatrix[currentIndex] = {
                ...approvalMatrix[currentIndex],
                Status: "Rejected",
                Comment: Comment,
                ActionDate: formattedDate
            };

            const updatedApprovalMatrix = JSON.stringify(approvalMatrix);

            // =====================================
            // ✅ WORKFLOW HISTORY (LIKE APPROVE)
            // =====================================

            let workflowHistory = [];

            try {
                workflowHistory = JSON.parse(item.WorkflowHistory || "[]");
            } catch {
                workflowHistory = [];
            }

            const newHistoryEntry = {
                CurrentApprover: currentUser.Title,
                ActionTaken: `${currentRole} Rejected`,
                Comment: Comment,
                Date: formattedDate,
                CurrentStatus: "Rejected"
            };

            workflowHistory.push(newHistoryEntry);

            const updatedHistory = JSON.stringify(workflowHistory);

            // =====================================
            // ✅ UPDATE SHAREPOINT ITEM
            // =====================================

            await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .update({
                    Status: "Rejected",
                    ApproverStatus: "Rejected",
                    WorkflowHistory: updatedHistory,
                    ApprovalMatrix: updatedApprovalMatrix
                });

            // =====================================
            // ✅ SUCCESS
            // =====================================

            alert(`Request rejected by ${currentRole}`);
            setComment("");

            // ✅ Dynamic redirect
            //  history.push(`/${currentRole}Dashboard`);
            navigate('/APTeam');


        } catch (error) {
            console.error("Error while rejecting:", error);
            alert("Something went wrong while rejecting.");
        } finally {
            setIsSubmitting(false);
        }
    };
    // ✅ UI must return something
    return (
        <div className='MainUplodForm' style={{ margin: "0px" }}>
            <div className='row'>
                <div className='col-md-12'>
                    <div className='Main-Boxpoup'>
                        <div className="bordered">
                            <img src={logo} />
                            <h1>Non PO Approval </h1>
                        </div>
                        <div className='displayWF'>{WorkflowJSX}</div>
                        <div className='borderedbox'>
                            <div className="heading1">
                                <label>Requestor Information</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-6'>
                                        <label htmlFor="Employee Code" className='font'>Request Date</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {RequestDate}</label>
                                    </div>
                                    <div className='col-md-6'>
                                        <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {EmployeeName}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-6'>
                                        <label htmlFor="Contact No" className='font'>Division</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {Divison}</label>
                                    </div>
                                    <div className='col-md-6'>
                                        <label htmlFor="Employee Status" className='font'>Email</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {Email}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-6'>
                                        <label htmlFor="Location" className='font'>HOD</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {HOD}</label>
                                    </div>
                                    <div className='col-md-6'>
                                        <label htmlFor="RM" className='font'>Department</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {Department}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-6'>
                                        <label htmlFor="Location" className='font'>Employee Status</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {EmpStatus}</label>
                                    </div>
                                    <div className='col-md-6'>
                                        <label htmlFor="RM" className='font'>Contact No</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {contactNo}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {Location}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {RMVal}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Status" className='font'>Status</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {Status}</label>
                                    </div>
                                </div>
                            </div>
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Invoice Details</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Code" className='font'>Vendor Code</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {VendorCode}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Name" className='font'>Vendor Name </label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {VendorName}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Contact No" className='font'>Nature of Expense</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {NatureExpense}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Status" className='font'>Invoice No</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {InvoiceNo}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font'>Invoice Date</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {InvoiceDate}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>Basic Amount</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {BasicAmount}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>GST Amount</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {GSTAmt}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>Other Charges (If any):</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {otherCharges}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>Total Amount</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {TotalAmount}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>Remarks for NON-PO Expense</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {PayRemark}</label>
                                    </div>
                                </div>
                            </div>
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Attached Documents</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-12'>
                                        {loadingExisting ? (
                                            <div className="text-muted">Loading documents…</div>
                                        ) : existingFiles.length === 0 ? (
                                            <div style={{ color: "#6b7280" }}>No documents available.</div>
                                        ) : (
                                            <div style={{ overflowX: "auto" }}>
                                                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                                                    <thead>
                                                        <tr style={{ background: "#f8fafc" }}>
                                                            <th style={th}>File</th>
                                                            <th style={{ ...th, width: 140 }}>Size</th>
                                                            <th style={{ ...th, width: 120 }}>Action</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        {existingFiles.map((f) => (
                                                            <tr key={`${f.source}-${f.name}`}>
                                                                <td style={td}>
                                                                    <a href={f.url} target="_blank" rel="noreferrer" style={{ color: "#0a66c2" }}>
                                                                        {f.name}
                                                                    </a>
                                                                </td>
                                                                <td style={{ ...td }}>
                                                                    {f.sizeBytes ? `${(f.sizeBytes / (1024 * 1024)).toFixed(2)} MB` : "—"}
                                                                </td>
                                                                <td style={td}>
                                                                    <a className="btn btn-sm btn-outline-primary" href={f.url} target="_blank" rel="noreferrer">
                                                                        Download
                                                                    </a>
                                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            </div>
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Work Flow History</label>
                            </div>
                            <div className="main-formcontainer">
                                <div className='Workflowbox'>
                                    {WorkflowHistory && WorkflowHistory.length > 0 ? (
                                        <table className="workflow-table" style={{ width: '100%', borderCollapse: 'collapse' }}>
                                            <thead>
                                                <tr>
                                                    <th style={{ padding: '8px', textAlign: 'left' }}>Date</th>
                                                    <th style={{ padding: '8px', textAlign: 'left' }}>Action By</th>
                                                    <th style={{ padding: '8px', textAlign: 'left' }}>Action Taken</th>
                                                    <th style={{ padding: '8px', textAlign: 'left' }}>Comment</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {WorkflowHistory.map((h: any, idx: number) => (
                                                    <tr key={idx}>
                                                        <td style={{ padding: '8px' }}>{h.Date ? new Date(h.Date).toLocaleString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).replace(",", "") : ""}</td>
                                                        <td style={{ padding: '8px' }}>{h.CurrentApprover || ''}</td>
                                                        <td style={{ padding: '8px' }}>{h.ActionTaken || ''}</td>
                                                        <td style={{ padding: '8px' }}>{h.Comment || ''}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    ) : (
                                        <p>No workflow history</p>
                                    )}
                                </div>
                            </div>
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Remark Section</label>
                            </div>
                            <div className="main-formcontainer">
                                <div className='row'>
                                    <div className='col-md-4'>
                                        <label className='font'>GL Code <span className='Mantorystar'>*</span></label>
                      
                                        <Dropdown options={GLCodeOptions} selectedKey={selectedGLCode}
                                         onChange={(e, option) => {
                                            const glCode = option?.key as string;
                                            setSelectedGLCode(glCode);
                                            // Find selected item immediately from state list
                                            const selectedItem =
                                                ExpenseTypemMastervalues.find(
                                                    (m: any) => m.GLCode === glCode
                                                );
                                            console.log(selectedItem);
                                            setNatureOfExpItems(selectedItem)
                                           

                                            // Set Nature of Expense
                                            setNatureExpense(
                                                selectedItem?.ExpenseType || ""
                                            );
                                            setGLDescription(
                                                selectedItem?.GLDescription || ""
                                            );
                                         }}/>
                                    </div>
                                    <div className='col-md-3'>
                                        <label className='font'>GL Description</label>
                                        <label className='fonttext'>  {GLDescription}</label>
                                    </div>
                                    <div className='col-md-2'>
                                        <label className='font'>Cost Center </label>
                                        <label className='fonttext'>  {CostCenterValue}</label>

                                        {/* <input type="text"
                                            value={CostCenterValue}
                                            onChange={(e) => setCostCenterValue(e.target.value)}
                                        /> */}

                                    </div>
                                    <div className='col-md-3'>
                                        <label className='font'>Cost Center Description</label>
                                        <label className='fonttext'>  {CostCenterDescription}</label>
                                    </div>
                                </div>

                            </div>
                            <div className='row'>
                                <div className="col-md-3">
                                    <label className="font"> Remark <span className='Mantorystar'>*</span></label>
                                    <textarea
                                        className="form-control"
                                        rows={3}
                                        value={Comment}
                                        onChange={(e) => setComment(e.target.value)}
                                        style={{ resize: "none" }}
                                    />
                                </div>
                            </div>
                            <div className='row my-3'>
                                <div className='col-md-12'>
                                    <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                        <a className="submit-btn" onClick={() => handleApprove()}>Approve</a>
                                        <a className="Reject-btn" onClick={() => handleReject()}>Reject</a>
                                        <a className="Reject-btn" onClick={() => handleSendback()}>Send Back</a>
                                        <a className="reset-btn" onClick={() => navigate("/APTeam")}>Exit</a>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};