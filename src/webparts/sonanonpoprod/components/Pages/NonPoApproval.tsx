import React, { useEffect, useState } from "react";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import { useParams } from "react-router-dom";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import SPCRUDOPS from "../../service/DAL/newspcrudops";
import "./CSS/NewRequest.scss";

import logo from "../../assets/sona-comstarlogo.png";

import { useNavigate } from 'react-router-dom';

interface IRouteParams extends Record<string, string | undefined> {
    id?: string;
}

export const NonPoApproval = (props: ISonanonpoprodProps) => {
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
    const [WorkflowJSX, setWorkflowJSX] = React.useState<JSX.Element[] | null>(null);
    const [Stage, setStageData] = React.useState(0);
    const [WorkflowHistory, setWorkflowHistory] = React.useState<any[]>([]);
    const [ItemID, setItemID] = useState<any>();
    const [Comment, setComment] = React.useState("");

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



    useEffect(() => {
        if (id) {
            loadItem(Number(id));
        }
    }, [id]);


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
    const loadItem = async (itemId: number) => {
        try {
            const sp = spfi().using(SPFx(props.currentSPContext));

            const item = await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .select("ID,Title,RequestDate,EmployeeName,Stage,ApprovalMatrix,WorkflowHistory,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus,ApprovalMatrix,VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,WorkflowHistory,Status,PaymentRemarks,ApprovalRemarks,VendorCode/VendorCode,VendorCode/Id,VendorCode/Title,NatureofExpense/Title,NatureofExpense/ExpenseType,CurrentApprover/ID,CurrentApprover/EMail,CurrentApprover/Title")
                .expand("VendorCode,NatureofExpense,CurrentApprover")();

            console.log(item);

            setItemID(itemId);
            setrequestNo(item.Title || "")
            setPayRemark(item.PaymentRemarks || "")

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
            setNatureExpense(item.NatureofExpense.ExpenseType || "");
            setInvoiceNo(item.InvoiceNumber || "");
            const invDate = new Date(item.InvoiceDate || "").toISOString().split("T")[0];

            setInvoiceDate(invDate);
            setBasicAmount(item.Basicamount || "");
            setGSTAmt(item.GST || "");
            setotherCharges(item.OtherCharges || "");
            setTotalAmount(item.Totalamount || "");
            setStageData(item.Stage || "");
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
            // 🔥 Load Final BG attachment from library



        } catch (error) {
            console.error("Error loading item:", error);
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
            const absUrl: string = props.currentSPContext?.pageContext?.web?.absoluteUrl || "";
            const origin = window.location.origin;


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

                console.error(error);
                alert("Something went wrong.");
            }

            setExistingFiles([...libFiles]);
        } catch (e) {
            console.warn("[APTeamAcceptance] existing docs load failed:", e);
            setExistingFiles([]);
        } finally {
            setLoadingExisting(false);
        }

    };
    const handleApprove = async () => {

        if (isSubmitting) return;

        if (!Comment || Comment.trim() === "") {
            alert("Please enter a comment");
            return;
        }
        setIsSubmitting(true);

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

            // =====================================
            // GET MASTER APPROVAL MATRIX FROM LIST
            // =====================================

            //  const spCrudOps = await SPCRUDOPS();

            const parentItems = await spCrudOps.getData(
                "NONPOApprovalMatrix",
                "Id,Role/RoleName,Approver/ID,Approver/Title,Approver/EMail",
                "Role,Approver",
                `Status eq 'Active'`,
                { column: "ID", isAscending: true }
            );

            // Convert to JSON structure
            const masterMatrix = parentItems.map((item: any) => {

                let pendingText = "";

                const roleName = item.Role?.RoleName || "";

                if (roleName === "CFO")
                    pendingText = "Pending At CFO";

                else if (roleName === "APPerformer")
                    pendingText = "Pending At APPerformer";

                else if (roleName === "RM")
                    pendingText = "Pending At RM";

                else if (roleName === "HOD")
                    pendingText = "Pending At HOD";

                return {
                    Role: roleName,
                    Level: item.Level?.Level || "",
                    UserID: item.Approver?.ID?.toString(),
                    User: item.Approver?.Title || "",
                    Email: item.Approver?.EMail || "",
                    PendingText: pendingText, // ✅ ADD THIS
                    Status: "Pending"
                };
            });

            // =====================================
            // PARSE APPROVAL MATRIX
            // =====================================

            let approvalMatrix: any[] = [];
            try {
                approvalMatrix = JSON.parse(item.ApprovalMatrix || "[]");
            }
            catch {
                alert("Invalid Approval Matrix");
                return;
            }

            // =====================================
            // FIND CURRENT APPROVER
            // =====================================

            const currentItem = approvalMatrix.find(
                x =>
                    String(x.UserID) === String(currentUser.Id) &&
                    x.Status === "Pending"
            );
            // const currentItem = approvalMatrix.find(x =>(
            //             String(x.UserID) === String(currentUser.Id) 

            //         ) 
            //         // ["RM", "HOD", "CFO"].includes(x.Role)
            // );
            if (!currentItem) {
                alert("You are not authorized");
                return;
            }

            const currentRole = currentItem.Role;
            const currentLevel = currentItem.Level;


            const normalizeId = (id: any) => parseInt(id);
            // =====================================
            // 🔥 UPDATE EXISTING MATRIX ONLY (FINAL)
            // =====================================

            const getLevelNumber = (val: any) => {
                if (!val) return "";
                const match = val.toString().match(/\d+/);
                return match ? match[0] : val.toString().trim();
            };

            let finalMatrix = approvalMatrix.map((oldItem: any) => {

                // ✅ DO NOT MODIFY STATIC ROLES
                if (["Initiator", "RM", "HOD"].includes(oldItem.Role)) {
                    return oldItem;
                }

                // 🔍 Match by LEVEL
                const masterItem = masterMatrix.find(
                    (m: any) =>
                        getLevelNumber(m.Role) === getLevelNumber(oldItem.Role)
                );

                if (!masterItem) return oldItem;

                const oldId = String(oldItem.UserID || "").trim();
                const newId = String(masterItem.UserID || "").trim();

                const roleChanged =
                    (masterItem.Role || "").trim() !== (oldItem.Role || "").trim();

                const userChanged = oldId !== newId;

                // ✅ UPDATE ONLY IF CHANGE
                if (roleChanged || userChanged) {

                    console.log("🔥 Updating Level:", oldItem.Level);

                    return {
                        ...oldItem,
                        Role: masterItem.Role,
                        UserID: masterItem.UserID,
                        User: masterItem.User,
                        UserEmail: masterItem.Email,
                        PendingText: masterItem.PendingText
                    };
                }

                return oldItem;
            });


            // =====================================
            // ✅ STEP 2: FIND CURRENT LEVEL (NOT USER)
            // =====================================

            //    const currentLevel = approvalMatrix[currentIndex]?.Level;

            const updatedIndex = approvalMatrix.findIndex(
                x =>
                    parseInt(x.UserID) === currentUser.Id &&
                    x.Status === "Pending"
            );
            if (updatedIndex === -1) {
                console.log("❌ INDEX NOT FOUND");
                return;
            }
            // =====================================
            // ✅ STEP 3: UPDATE CURRENT APPROVER
            // =====================================

            // const nowed = new Date();
            const formattedDateed = new Date().toISOString();

            if (updatedIndex !== -1) {
                finalMatrix[updatedIndex] = {
                    ...finalMatrix[updatedIndex],
                    Status: "Approved",
                    Comment: Comment,
                    ActionDate: formattedDateed
                };
            }
if (finalMatrix.length > updatedIndex + 1) {

    finalMatrix[updatedIndex + 1] = {
        ...finalMatrix[updatedIndex + 1],
        Status: "Pending"
    };

}

            // =====================================
            // ✅ STEP 4: GET NEXT APPROVER (SAFE)
            // =====================================

            let nextApprover = null;

            if (finalMatrix.length > updatedIndex + 1) {
                nextApprover = finalMatrix[updatedIndex + 1];
            }


            // =====================================
            // ✅ STEP 5: FINAL JSON
            // =====================================

            const updatedApprovalMatrix = JSON.stringify(finalMatrix);

            // =====================================
            // DATE FORMAT
            // =====================================

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
                CurrentStatus: nextApprover
                    ? `Submitted to ${nextApprover.Role}`
                    : `${currentRole} Approved`
            };

            workflowHistory.push(newHistoryEntry);

            setWorkflowHistory(workflowHistory);

            const updatedHistory = JSON.stringify(workflowHistory);

            // =====================================
            // PREPARE UPDATE FIELDS
            // =====================================

            let updateFields: any = {};

            if (currentRole === "CFO") {

                updateFields = {

                    Status: "Pending for Acceptance",
                    PendingAt: nextApprover ? nextApprover.PendingText : " ",
                    ApproverStatus: "Approved",
                    Stage: updatedIndex + 1,
                    CurrentApproverId: nextApprover
                        ? parseInt(nextApprover.UserID)
                        : null,
                    WorkflowHistory: updatedHistory,
                    ApprovalMatrix: updatedApprovalMatrix
                };

            }
            else {

                updateFields = {

                    Status: "Pending for Approval",
                    PendingAt: nextApprover ? nextApprover.PendingText : "",
                    ApproverStatus: "Pending",
                    Stage: updatedIndex + 1,
                    CurrentApproverId: nextApprover
                        ? parseInt(nextApprover.UserID)
                        : null,
                    WorkflowHistory: updatedHistory,
                    ApprovalMatrix: updatedApprovalMatrix
                };
            }

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
            navigate('/MyApproval');


        }
        catch (error) {
            console.error(error);
            alert("Approval failed");
        }
        finally {
            setIsSubmitting(false);
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
                parseInt(x.UserID) === currentUser.Id);

            const currentItem = approvalMatrix.find(x =>
                String(x.UserID) === String(currentUser.Id));



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

            // const Empitems = await sp.web.lists
            //     .getByTitle("EmployeeMaster")
            //     .items
            //     .filter(`EmployeeName eq '${item.EmployeeName}'`)
            //     .select("*,Employee/Title,Employee/Id")
            //     .expand("Employee")();
            // const emp = Empitems[0];
            // const EmpId = emp?.Employee?.Id || "";



            await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(Number(id))
                .update({
                    Status: "Send Back",
                    PendingAt: 'Pending At Initiator',
                    CurrentApproverId: null,
                    WorkflowHistory: updatedHistory,
                    ApprovalMatrix: updatedApprovalMatrix,

                });

            setComment("");
            alert(`Request Send Back by ${currentRole}`);
            navigate('/MyApproval');

        } catch (error) {

            console.error(error);
            alert("Something went wrong.");

        } finally {

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
            navigate('/MyApproval');


        } catch (error) {
            console.error("Error while rejecting:", error);
            alert("Something went wrong while rejecting.");
        } finally {
            setIsSubmitting(false);
        }
    };

    const handleOpenFile = (url: string) => {
        window.open(url, "_blank", "noopener,noreferrer");
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
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Requestor Information</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Code" className='font fontblock'>Request Date : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {RequestDate}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label className='font fontblock'>Employee Name  : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {EmployeeName}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label className='font fontblock'>Email : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {Email}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label className='font fontblock'>Contact No : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {contactNo}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>Department : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {Department}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label className='font fontblock'>Division : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {Divison}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font fontblock'>Location : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {Location}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>RM : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {RMVal}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font fontblock'>HOD : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {HOD}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font fontblock'>Employee Status : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {EmpStatus}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Status" className='font fontblock'>Status : &nbsp;&nbsp; </label>
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
                                        <label htmlFor="Employee Code" className='font fontblock'>Vendor Code : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {VendorCode}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Name" className='font fontblock'>Vendor Name  : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {VendorName}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Contact No" className='font fontblock'>Nature of Expense : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {NatureExpense}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Status" className='font fontblock'>Invoice No : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {InvoiceNo}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font fontblock'>Invoice Date : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {InvoiceDate}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>Basic Amount : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {BasicAmount}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>GST Amount : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {GSTAmt}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>Other Charges (If any): : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {otherCharges}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>Total Amount : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {TotalAmount}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font fontblock'>Remarks for NON-PO Expense : &nbsp;&nbsp; </label>
                                        <label className='fonttext'>  {PayRemark}</label>
                                    </div>
                                </div>
                            </div>

                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Attached Documents</label>
                            </div>
                            <div className='main-formcontainer'>
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
                                                            <a onClick={() => handleOpenFile(f.url)} style={{ color: "#0a66c2" }}>
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
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Work Flow History</label>
                            </div>
                            <div className="main-formcontainer">
                                <div className='Workflowbox'>
                                    {WorkflowHistory && WorkflowHistory.length > 0 ? (
                                        <table className="workflow-table" style={{ width: '100%', borderCollapse: 'collapse' }}>
                                            <thead>
                                                <tr>
                                                    <th style={{ padding: '8px', textAlign: 'left' }}>Action Date</th>
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
                                    <div className="col-md-12">
                                        <label className="font">Approver Remark <span className='Mantorystar'>*</span></label>
                                        <textarea
                                            className="form-control"
                                            rows={3}
                                            value={Comment}
                                            onChange={(e) => setComment(e.target.value)}
                                            style={{ resize: "none" }}
                                        />
                                    </div>
                                </div>
                            </div>
                            <div style={{ display: "flex", justifyContent: "center", gap: "5px", margin: "10px 20px" }}>
                                <a className="submit-btn" onClick={() => handleApprove()}>Approve</a>
                                <a className="sendback-btn" onClick={() => handleSendback()}>Send Back</a>
                                <a className="Reject-btn" onClick={() => handleReject()}>Reject</a>
                                <a className="reset-btn" onClick={() => navigate("/MyApproval")}>Exit</a>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>



    );
};