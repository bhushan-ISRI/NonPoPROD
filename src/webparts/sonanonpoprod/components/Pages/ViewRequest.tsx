import React, { useEffect, useState } from "react";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import { useParams } from "react-router-dom";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import SPCRUDOPS from "../../service/DAL/newspcrudops";
import "./CSS/NewRequest.scss";
import { useNavigate } from 'react-router-dom';

import logo from "../../assets/sona-comstarlogo.png";

interface IRouteParams extends Record<string, string | undefined> {
    id?: string;
}

export const ViewRequest = (props: ISonanonpoprodProps) => {
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

    const [VouAmount, setVouAmount] = React.useState<string>("");
    const [GLCode, setGLCode] = React.useState<string>("");
    const [CostCentre, setCostCentre] = React.useState<string>("");
    const [GLDescription, setGLDescription] = React.useState<any | null>(null);
    const [CostCenterDescription, setCostCenterDescription] = React.useState<any | null>(null);
    const [VoucherNumber, setVoucherNumber] = React.useState<string>("");

    const [GSTAmt, setGSTAmt] = React.useState<string>("");
    const [otherCharges, setotherCharges] = React.useState<string>("");
    const [TotalAmount, setTotalAmount] = React.useState<string>("");
    const [existingFiles, setExistingFiles] = useState<IExistingFile[]>([]);
    const [loadingExisting, setLoadingExisting] = useState(false);
    const [loading, setLoading] = useState(true);




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

        data.forEach((m, i) => {
            let isActive = "";

            if (Status?.trim().toLowerCase() === "closed") {
                // ✅ All completed when closed
                isActive = "beforeactiveApprover";
            } else {
                if (i < Stage) {
                    isActive = "beforeactiveApprover"; // GREEN
                } else if (i === Stage) {
                    isActive = "overrideStage"; // YELLOW
                } else {
                    isActive = "overrideStage"; // NORMAL
                }
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
    const displayWorkflow1 = (matrix?: any[]) => {
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

    const displayWorkflownew = (matrix?: any[]) => {
        const data = matrix || ApprovalMatrixdata;

        const wf: JSX.Element[] = [];
        let isActive;
        let notActive = false;

        data.forEach((m, i) => {

            // ✅ NEW: If Closed → no active highlight
            if (Status === "Closed") {
                isActive = "overrideStage";
            } else {
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
        setLoading(true);
        try {
            const sp = spfi().using(SPFx(props.currentSPContext));

            const item = await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(itemId)
                .select("ID,Title,GLDescription,CostCenterDescription,Amount,GLCode,CostCentre,VoucherNumber,RequestDate,EmployeeName,Stage,ApprovalMatrix,WorkflowHistory,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus,ApprovalMatrix,VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,WorkflowHistory,Status,PaymentRemarks,ApprovalRemarks,VendorCode/VendorCode,VendorCode/Title,VendorCode/Id,NatureofExpense/Title,NatureofExpense/ExpenseType,CurrentApprover/ID,CurrentApprover/EMail,CurrentApprover/Title")
                .expand("VendorCode,NatureofExpense,CurrentApprover")();

            console.log(item);
            // setLoading(false);

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
            setGLDescription(item.GLDescription || "");
            setCostCenterDescription(item.CostCenterDescription || "");

            setInvoiceDate(invDate);
            setBasicAmount(item.Basicamount || "");
            setGSTAmt(item.GST || "");
            setotherCharges(item.OtherCharges || "");
            setTotalAmount(item.Totalamount || "");
            setStageData(item.Stage || "");
            setVouAmount(item.Amount || "");
            setGLCode(item.GLCode || "");
            setCostCentre(item.CostCentre || "");
            setVoucherNumber(item.VoucherNumber || "");

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
            setLoading(false);
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
                        //  url: `${absUrl}${f.ServerRelativeUrl}`,
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

    const handleOpenFile = (url: string) => {
        window.open(url, "_blank", "noopener,noreferrer");
    };

    // ✅ UI must return something
    return (
        <div>
            <div className="bordered">
                <a><img src={logo} /></a>
                <h1>Non PO Approval</h1>
            </div>
            <div className='displayWF'>{WorkflowJSX}</div>

            <div className='borderedbox'>
                <div className="heading1">
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
                            <label className='font fontblock'>Department : &nbsp;&nbsp; </label>
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
                            <label className='font fontblock'>RM : &nbsp;&nbsp; </label>
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
                <div className="heading1">
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
                            <label className='font fontblock'>Basic Amount : &nbsp;&nbsp; </label>
                            <label className='fonttext'>  {BasicAmount}</label>
                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4'>
                            <label className='font fontblock'>GST Amount : &nbsp;&nbsp; </label>
                            <label className='fonttext'>  {GSTAmt}</label>
                        </div>
                        <div className='col-md-4'>
                            <label className='font fontblock'>Other Charges (If any): : &nbsp;&nbsp; </label>
                            <label className='fonttext'>  {otherCharges}</label>
                        </div>
                        <div className='col-md-4'>
                            <label className='font fontblock'>Total Amount : &nbsp;&nbsp; </label>
                            <label className='fonttext'>  {TotalAmount}</label>
                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4'>
                            <label className='font fontblock'>Remarks for NON-PO Expense : &nbsp;&nbsp; </label>
                            <label className='fonttext'>  {PayRemark}</label>
                        </div>
                    </div>
                </div>
                
                {Status === "Pending for Vouching" && (
                    <>
                        <div className="heading1">
                            <label>AP Performer Updates</label>
                        </div>
                        <div className='main-formcontainer'>
                            <div className='row mb-20'>
                                <div className='col-md-6'>
                                    <label className='font fontblock'>GL Code : &nbsp;&nbsp; </label>
                                    <label className='fonttext'>  {GLCode}</label>
                                </div>
                                <div className='col-md-6'>
                                    <label className='font fontblock'>GL Code Description : &nbsp;&nbsp; </label>
                                    <label className='fonttext'>  {GLDescription}</label>
                                </div>
                            </div>
                            <div className='row mb-20'>
                                <div className='col-md-6'>
                                    <label className='font fontblock'>Cost Center : &nbsp;&nbsp; </label>
                                    <label className='fonttext'>  {CostCentre}</label>
                                </div>
                                <div className='col-md-6'>
                                    <label className='font fontblock'>Cost Center Description: : &nbsp;&nbsp; </label>
                                    <label className='fonttext'>  {CostCenterDescription}</label>
                                </div>
                            </div>
                            <div className='row mb-20'>
                                <div className='col-md-6'>
                                    <label className='font fontblock'> Voucher Number : &nbsp;&nbsp; </label>
                                    <label className='fonttext'>  {VoucherNumber}</label>
                                </div>
                                <div className='col-md-6'>
                                    <label className='font fontblock'>Voucher Amount : &nbsp;&nbsp; </label>
                                    <label className='fonttext'>  {VouAmount}</label>
                                </div>


                            </div>
                        </div>
                    </>
                )}
                
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
                <div style={{ display: "flex", justifyContent: "center", gap: "5px", margin: "10px 20px" }}>
                    <button className="reset-btn" onClick={() => navigate(-1)}>Exit</button>
                </div>
            </div>
        </div>

    );
};