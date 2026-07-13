import React, { useEffect, useRef, useState } from "react";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import { useParams } from "react-router-dom";


import logo from "../../assets/sona-comstarlogo.png";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import SPCRUDOPS from "../../service/DAL/newspcrudops";
import "./CSS/NewRequest.scss";

import { TextField, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { Upload, Button } from "antd";
import { UploadOutlined } from "@ant-design/icons";

import { useNavigate } from 'react-router-dom';

interface IRouteParams extends Record<string, string | undefined> {
    id?: string;
}

export const ViewRequestSendBack = (props: ISonanonpoprodProps) => {
    const navigate = useNavigate();

    const { id } = useParams<IRouteParams>();
    const [RequestDate, setRequestDate] = React.useState<string>();
    // const [EmployeeName, setEmployeeName] = React.useState<string>();
    const [Divison, setDivison] = React.useState<string>();
    const [Email, setEmail] = React.useState<string>();
    // const approverJson = React.useRef<any[]>(null);
    const approverJson = React.useRef<any[]>([]);
    // const [Employeemasterdata, setEmployeeMasterdata] = React.useState<any[]>([]);
    const [ExpenseMasterdata, setExpenseMasterdata] = React.useState<any[]>([]);
    const [NatureOfExpenseTypeOptions, setNatureOfExpenseTypeOptions] = React.useState<IDropdownOption[]>([]);
    const EmployeeQmadata = React.useRef<any[]>([]);
    const currentUserId = React.useRef(0);
    // let EmployeeQmadata = React.useRef<any[]>(null);

    const [HOD, setHOD] = React.useState<string>();
    const [EmpStatus, setEmpStatus] = React.useState<string>();
    const [Status, setStatus] = React.useState<string>();
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [ApprovalMatrixdata, setApprovalMatrix] = React.useState<any[]>([]);
    //const approvalMatrix = React.useState<any[]>([]);// useRef<any[]>(null);

    const approvalMatrix = React.useRef<any[]>([]);

    //approvalMatrix.current = approvalMatrixJSON; // ✅ correct
    const [WorkflowJSX, setWorkflowJSX] = React.useState<JSX.Element[] | null>(null);
    const [Stage, setStageData] = React.useState(0);
    const [WorkflowHistory, setWorkflowHistory] = React.useState<any[]>([]);
    const [ItemID, setItemID] = useState<any>();
    const [Comment, setComment] = React.useState("");

    const [currentDate, setCurrentDate] = React.useState("");

    const [requestNo, setrequestNo] = useState<any>();// useState<string | undefined>(undefined);

    const [contactNo, setcontactNo] = React.useState<string>();
    const [Location, setLocation] = React.useState<string>();
    const [RMVal, setRMVal] = React.useState<string>("");
    const [VendorCode, setVendorCode] = React.useState<string>("");
    const [VendorName, setVendorName] = React.useState<string>("");
    const [NatureExpense, setNatureExpense] = React.useState<string>("");
    const [InvoiceNo, setInvoiceNo] = React.useState<string>("");
    const [InvoiceDate, setInvoiceDate] = React.useState<string>("");

    const [PayRemark, setPayRemark] = React.useState<string>("");

    const [VouAmount, setVouAmount] = React.useState<string>("");
    const [GLCode, setGLCode] = React.useState<string>("");
    const [CostCentre, setCostCentre] = React.useState<string>("");
    const [VoucherNumber, setVoucherNumber] = React.useState<string>("");



    const [BasicAmount, setBasicAmount] = useState<number>(0);
    const [GSTAmt, setGSTAmt] = useState<number>(0);
    const [otherCharges, setotherCharges] = useState<number>(0);
    const [TotalAmount, setTotalAmount] = useState<number>(0);

    //const [TotalAmount,setTotalAmount] = React.useState<string>("");
    const [existingFiles, setExistingFiles] = useState<IExistingFile[]>([]);
    const [loadingExisting, setLoadingExisting] = useState(false);
    const [originalData, setOriginalData] = useState<any>(null);
    const [selectedNatureofexpense, setSelectedNatureofexpense] = React.useState<number | undefined>();
    const [GLDescription, setGLDescription] = React.useState<any | null>(null);
    const [CostCenterDescription, setCostCenterDescription] = React.useState<any | null>(null);



    const [POAttachments, setPOAttachments] = useState<any>({
        allFiles: []
    });

    const [POTextAttachFile, setPOTextAttachFile] = useState<File | null>(null);


    // const [currentDate, setCurrentDate] = React.useState("");
    const [Employeemasterdata, setEmployeemasterdata] = React.useState<any[]>([]);

    const [EmployeeId, setEmployeeId] = React.useState("");
    const [EmployeeName, setCurrentUserName] = React.useState("");
    const [EmployeeEmail, setCurrentUserEmail] = React.useState("");
    const [EmployeeContactno, setEmployeeContactno] = React.useState("");
    const [EmployeeStatus, setEmployeeStatus] = React.useState("");
    const [EmployeeDivision, setEmployeeDivision] = React.useState("");
    const [Locationdata, setLocationdata] = React.useState("");
    const [EmployeeRM, setEmployeeRM] = React.useState("");
    const [hoddata, sethoddata] = React.useState("");
    const [Department, setDepartment] = React.useState<string>();

    const IntitorID = React.useRef<number | null>(null);
    const RMID = React.useRef<number | null>(null);
    const HODID = React.useRef<number | null>(null);
    const [CostCenter, setCostCenter] = React.useState<string>();


    const currentUserEmail = React.useRef<string>("");




    useEffect(() => {
        if (id) {
            loadItem(Number(id));
            setItemID(id)
        }
    }, [id]);

    React.useEffect(() => {
        const loadCurrentUser = async () => {
            try {
                const sp = spfi().using(SPFx(props.currentSPContext));
                const user = await sp.web.currentUser();

                currentUserId.current = user.Id;

                const today = new Date().toISOString().split("T")[0];
                setCurrentDate(today);

            } catch (error) {
                console.error("Error fetching current user:", error);
            }
        };

        loadCurrentUser();
    }, [props.currentSPContext]);


    let getdata = async (userEmail: string) => {
        await fetchExpenseMaster();
        await fetchEmployeeDataV2(userEmail);
        await fetchApprovalMatrix();
        await getapprovalmatrix(userEmail);
        await displayWorkflow();
    };

    React.useEffect(() => {
        const loadCurrentUser = async () => {
            try {
                const sp = spfi().using(SPFx(props.currentSPContext));
                const user = await sp.web.currentUser();

                currentUserId.current = user.Id;
                currentUserEmail.current = user.Email;
                // currentUserEmail.current = "damodar@sonacomstar.com";

                // Call after email is available
                await getdata(user.Email);

                const today = new Date().toISOString().split("T")[0];
                setCurrentDate(today);

            } catch (error) {
                console.error("Error fetching current user:", error);
            }
        };

        loadCurrentUser();
    }, [props.currentSPContext]);


    const fetchEmployeeDataV2 = async (userEmail: string) => {
        try {

            const toTitleCase = (str: string): string => {
                if (!str) return "";

                return str
                    .toLowerCase()
                    .split(" ")
                    .filter(Boolean)
                    .map((w: string) => w.charAt(0).toUpperCase() + w.slice(1))
                    .join(" ");
            };

            const cleanLocationForDisplay = (location: string): string => {
                if (!location) return "N/A";

                return location.replace(/^re\s+/i, "").trim();
            };

            // const FLOW_URL =
            //     "https://defaultcb1edbfe8080457d9cae51528f3643.3f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e2bb522aa41443179a72b701b9613471/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=q8b8ADCtK2eKr2f6p3MX7gxmJymPeJbm0mq2M69Rk8E";

            // const fetchPage = async (pageNumber: number) => {

            //     const response = await fetch(FLOW_URL, {
            //         method: "POST",
            //         headers: {
            //             "Content-Type": "application/json"
            //         },
            //         body: JSON.stringify({
            //             PageSize: 500,
            //             PageNumber: pageNumber
            //         })
            //     });

            //     if (!response.ok) {
            //         throw new Error("Failed to fetch employee data");
            //     }

            //     return response.json();
            // };
            const fetchPage = async (pageNumber: number) => {

                const username = "0le867nyvalvfo249e6sj4ri";
                const password = "2mpvr7r19amf7o01hr0qncr861hmtsb7o9ap51hwar72405atj3y73mndkmokg5i";

                const auth = btoa(`${username}:${password}`);

                const responsesevices = await fetch(
                    "https://mservices.zinghr.com/etl/api/v2/Auth/GenerateJWTToken?apiPermission=GEMD",
                    {
                        method: "GET",
                        headers: {
                            "Authorization": `Basic ${auth}`,
                            "Accept": "application/json",
                            "Content-Type": "application/json",
                        },
                    }
                );

                const ServicedataToken = await responsesevices.json();
                console.log(ServicedataToken.data);

                const response = await fetch("https://mservices.zinghr.com/etl/api/v2/Employee/GetEmployeeDetails", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "Authorization": `Bearer ${ServicedataToken.data}`,
                        "ClientSecret": "2mpvr7r19amf7o01hr0qncr861hmtsb7o9ap51hwar72405atj3y73mndkmokg5i"
                    },
                    body: JSON.stringify({
                        PageSize: 500,
                        PageNumber: pageNumber,
                    }),
                });
                // const data = await response.json();
                // console.log("Employee data:", data);
                if (!response.ok) {
                    throw new Error("Failed to fetch employee data");
                }

                return response.json();
            };
            let allEmployees: any[] = [];
            let page = 1;
            let hasMoreData = true;

            while (hasMoreData) {

                const res = await fetchPage(page);

                const employees = res?.data?.employees || [];

                allEmployees = [...allEmployees, ...employees];

                if (employees.length < 500) {
                    break;
                }

                page++;
            }

            const mappedEmployees = allEmployees.map((item: any) => {

                const locationAttr = (item.attributes || []).find(
                    (a: any) => a.attributeTypeDescription === "Location"
                );

                const departmentAttr = (item.attributes || []).find(
                    (a: any) =>
                        a.attributeTypeDescription?.toLowerCase() === "department"
                );

                const hodAttr = (item.attributes || []).find(
                    (a: any) =>
                        a.attributeTypeDescription?.toLowerCase() === "hod name"
                );

                const hodemailAttr = (item.attributes || []).find(
                    (a: any) =>
                        a.attributeTypeDescription?.toLowerCase() === "hod_email"
                );

                const CostCenterAttr = (item.attributes || []).find(
                    (a: any) =>
                        a.attributeTypeDescription?.toLowerCase() === "cost center"
                );


                // const user = sp.web.ensureUser(item.email);

                const today = new Date().toISOString().split("T")[0];
                setCurrentDate(today);
                return {
                    EmployeeCode: item.employeeCode || "",
                    EmployeeName: toTitleCase(item.employeeName || ""),
                    Email: item.email || "",
                    MobileNo: item.mobileNo || "",
                    EmployeeStatus: item.employeeStatus || "",
                    Department:
                        departmentAttr?.attributeTypeUnitDescription || "",
                    Location:
                        cleanLocationForDisplay(
                            locationAttr?.attributeTypeUnitDescription
                        ),
                    CostCenter: CostCenterAttr?.attributeTypeUnitDescription || "",
                    RMName: item.reportingManagerName,// === "Piyush Airan" ? "Prince Gupta" : item.reportingManagerName || "",
                    RMEmail: item.reportingManagerEmail,// === "piyush.airan@sonacomstar.com" ? "prince.gupta@sonacomstar.com" : item.reportingManagerEmail || "",
                    HODName: hodAttr?.attributeTypeUnitDescription,//, === "Piyush Airan" ? "Prince Gupta" : hodAttr?.attributeTypeUnitDescription || "",
                    HODEmail: hodemailAttr?.attributeTypeUnitDescription,// === "piyush.airan@sonacomstar.com" ? "prince.gupta@sonacomstar.com" : hodemailAttr?.attributeTypeUnitDescription || "",


                    // HODName: item.reportingManagerName2 || "",
                    // HODEmail: item.reportingManagerEmail2 || ""
                };
            });

            // set all employees
            setEmployeemasterdata(mappedEmployees);
            EmployeeQmadata.current = mappedEmployees;

            // const userEmail = "damodar@sonacomstar.com"
            // const userEmail = "uday.p@sonacomstar.com"

            // const emp = mappedEmployees.find(
            //   x => x.Email?.toLowerCase() === userEmail
            // );


            const emp = mappedEmployees.find(
                x =>
                    x.Email?.trim().toLowerCase() ===
                    userEmail?.trim().toLowerCase()
            );

            // if (!sp) return;

            if (emp) {

                const sp = spfi().using(SPFx(props.currentSPContext));

                setEmployeeId(emp.EmployeeCode);
                setCurrentUserName(emp.EmployeeName);
                setCurrentUserEmail(emp.Email);
                const IntitorUser = await sp.web.ensureUser(emp.Email);

                IntitorID.current = IntitorUser.Id;
                setEmployeeContactno(emp.MobileNo);
                setEmployeeStatus(emp.EmployeeStatus);
                setEmployeeDivision(emp.Department);
                setLocationdata(emp.Location);
                setEmployeeRM(emp.RMName);
                setDepartment(emp.Department);
                setCostCenter(emp.CostCenter);
                const rmUser = await sp.web.ensureUser(emp.RMEmail);
                RMID.current = rmUser.Id;
                const HodUser = await sp.web.ensureUser(emp.HODEmail);
                HODID.current = HodUser.Id;
                sethoddata(emp.HODName);
            }

        } catch (error) {
            console.error("fetchEmployeeDataV2 error:", error);
        }
    };

    const fetchApprovalMatrix = async () => {
        try {
            const sp = spfi().using(SPFx(props.currentSPContext));

            const parentItems = await sp.web.lists
                .getByTitle("NONPOApprovalMatrix")
                .items
                .select(
                    "*",
                    "Id,ApprovalType,Role/ID,Role/RoleName,Level/ID,Level/Level,Status,Approver/ID,Approver/Title,Approver/EMail")
                .expand("Role", "Level", "Approver")
                .filter("Status eq 'Active'")()

            const approvalMatrixJSON = parentItems.map((item: any) => ({
                Role: item.Role?.RoleName || "",
                Approver: item.Approver?.Title || "",
                ApproverID: item.Approver?.ID?.toString() || "",
                ApproverEmail: item.Approver?.EMail || ""
            }));
            approvalMatrix.current = approvalMatrixJSON;


        } catch (error) {
            console.error("Error fetching EmployeeMaster:", error);
        }
    };

    let getapprovalmatrix = async (userEmail: string) => {

        const approverArray: any[] = [];

        const matchedEmployee = EmployeeQmadata.current.find(
            x =>
                x.Email?.trim().toLowerCase() ===
                userEmail?.trim().toLowerCase()
        );

        if (!matchedEmployee) {
            console.log("Current user not found");
            return [];
        }

        // Initiator
        approverArray.push({
            Role: "Initiator",
            User: matchedEmployee.EmployeeName,
            UserID: IntitorID.current,
            UserEmail: matchedEmployee.Email,
            Status: "Approved"
        });

        // RM
        approverArray.push({
            Role: "RM",
            User: matchedEmployee.RMName,
            UserID: RMID.current,
            UserEmail: matchedEmployee.RMEmail,
            PendingText: "Pending At RM",
            Status: "Pending"
        });

        // HOD only if different from RM
        if (String(RMID.current) !== String(HODID.current)) {

            approverArray.push({
                Role: "HOD",
                User: matchedEmployee.HODName,
                UserID: HODID.current,
                UserEmail: matchedEmployee.HODEmail,
                PendingText: "Pending At HOD",
                Status: "Waiting"
            });

        }

        // Add other approvers from ApprovalMatrix list
        approvalMatrix.current.forEach((item: any) => {

            const role = item.Role || "";

            const approverId = item.ApproverID?.toString();

            if (!approverId) return;

            if (role === "RM" || role === "HOD") return;

            let pendingText = "";

            if (role === "CFO")
                pendingText = "Pending At CFO";

            else if (role === "APPerformer")
                pendingText = "Pending At AP Performer";

            // Duplicate check
            // const alreadyExists =
            //   approverArray.some(
            //     x => String(x.UserID) === String(approverId)
            //   );

            // if (alreadyExists) return;

            approverArray.push({
                Role: role,
                User: item.Approver || "",
                UserID: approverId,
                UserEmail: item.ApproverEmail || "",
                PendingText: pendingText,
                Status: "Waiting"
            });

        });

        approverJson.current = approverArray;
    }





    // let getdata = async () => {
    //     // await userProfile();
    //     await fetchExpenseMaster();
    //     await fetchEmployeeMaster();
    //     await fetchApprovalMatrix();
    //     await getapprovalmatrix();
    //     await displayWorkflow();
    // };

    // React.useEffect(() => {

    //     if (!Email || Employeemasterdata.length === 0) return;

    //     // Match current user email with EmployeeMaster email
    //     const matchedEmployee = Employeemasterdata.find(
    //         emp =>
    //             emp.Employee.Id &&
    //             emp.Employee.Id === currentUserId.current
    //     );

    //     if (matchedEmployee) {

    //         // ✅ ADD THIS LINE (MOST IMPORTANT)
    //         RMID.current = matchedEmployee.ReportingManager?.ID || null;
    //     }

    // }, [Email, Employeemasterdata]);



    const fetchExpenseMaster = async () => {
        try {
            const sp = spfi().using(SPFx(props.currentSPContext));

            const Expitems = await sp.web.lists
                .getByTitle("ExpenseMaster")
                .items
                .select("*")
                .filter("Status eq 'Active'")()
            // EmployeeQmadata.current = Expitems;
            setExpenseMasterdata(Expitems);
            const options: IDropdownOption[] = Expitems.map(item => ({
                key: item.Id,          // ✅ IMPORTANT → Id
                text: item.ExpenseType      // display text
            }));

            setNatureOfExpenseTypeOptions(options);

        } catch (error) {
            console.error("Error fetching EmployeeMaster:", error);
        }
    };

    const handleAmountChange =
        (name: "BasicAmount" | "GSTAmt" | "OtherCharges") =>
            (e: React.ChangeEvent<HTMLInputElement>) => {
                const value = Number(e.target.value) || 0;

                const newBasic = name === "BasicAmount" ? value : BasicAmount;
                const newGST = name === "GSTAmt" ? value : GSTAmt;
                const newOther = name === "OtherCharges" ? value : otherCharges;

                if (name === "BasicAmount") setBasicAmount(value);
                if (name === "GSTAmt") setGSTAmt(value);
                if (name === "OtherCharges") setotherCharges(value);
                const tot = Number(newBasic) +
                    Number(newGST) +
                    Number(newOther);
                setTotalAmount(tot);
            };
    // const fetchApprovalMatrix = async () => {
    //     try {
    //         const sp = spfi().using(SPFx(props.currentSPContext));

    //         const parentItems = await sp.web.lists
    //             .getByTitle("NONPOApprovalMatrix")
    //             .items
    //             .select(
    //                 "*",
    //                 "Id,ApprovalType,Role/ID,Role/RoleName,Level/ID,Level/Level,Status,Approver/ID,Approver/Title,Approver/EMail")
    //             .expand("Role", "Level", "Approver")
    //             .filter("Status eq 'Active'")()

    //         const approvalMatrixJSON = parentItems.map((item: any) => ({
    //             Role: item.Role?.RoleName || "",
    //             Approver: item.Approver?.Title || "",
    //             ApproverID: item.Approver?.ID?.toString() || "",
    //             ApproverEmail: item.Approver?.EMail || ""
    //         }));
    //         approvalMatrix.current = approvalMatrixJSON;


    //     } catch (error) {
    //         console.error("Error fetching EmployeeMaster:", error);
    //     }
    // };

    // let getapprovalmatrix = async () => {

    //     const approverArray: any[] = [];

    //     const matchedEmployee = EmployeeQmadata.current.find(
    //         emp => String(emp.EmployeeId) === String(currentUserId.current)
    //     );

    //     // const matchedEmployee = ApprovalMatrixdata.find(
    //     //     emp => String(emp.EmployeeId) === String(currentUserId.current)
    //     // );

    //     approverArray.push({
    //         Role: "Initiator",
    //         User: matchedEmployee.Employee.Title,
    //         UserID: matchedEmployee.Employee.Id,
    //         UserEmail: matchedEmployee.Employee.EMail,
    //         Status: "Approved"
    //     });

    //     approverArray.push({
    //         Role: "RM",
    //         User: matchedEmployee.ReportingManager.Title,
    //         UserID: matchedEmployee.ReportingManager.ID,
    //         UserEmail: matchedEmployee.ReportingManager.EMail,
    //         PendingText: "Pending At RM",
    //         Status: "Pending"
    //     });

    //     approverArray.push({
    //         Role: "HOD",
    //         User: matchedEmployee.HOD.Title,
    //         UserID: matchedEmployee.HOD.ID,
    //         UserEmail: matchedEmployee.HOD.EMail,
    //         PendingText: "Pending At HOD",
    //         Status: "Pending"
    //     });


    //     // Add other approvers from ApprovalMatrix list
    //     approvalMatrix.current.forEach((item: any) => {

    //         const role = item.Role || "";

    //         const approverId = item.ApproverID?.toString();

    //         if (!approverId) return;

    //         if (role === "RM" || role === "HOD") return;

    //         let pendingText = "";

    //         if (role === "CFO")
    //             pendingText = "Pending At CFO ";

    //         else if (role === "APPerformer")
    //             pendingText = "Pending At APPerformer ";

    //         const alreadyExists =
    //             approverArray.some(x => x.ApproverID === approverId);
    //         if (alreadyExists) return;

    //         approverArray.push({
    //             Role: role,
    //             User: item.Approver || "",
    //             UserID: approverId,
    //             UserEmail: item.ApproverEmail || "",
    //             PendingText: pendingText,
    //             Status: "Pending"
    //         });
    //     });
    //     approverJson.current = approverArray;
    // }
    React.useEffect(() => {
        if (ApprovalMatrixdata.length > 0) {
            displayWorkflow();
        }
    }, [ApprovalMatrixdata]);

    const displayWorkflow = () => {
        const wf: JSX.Element[] = [];

        // const _wf = approverJson.filter((item) => item.required === true);
        let isActive;
        let notActive = false;
        approverJson.current.forEach((m, i) => {
            //if (m.required === true) {
            if (notActive === false && Stage !== 99) {
                if (Stage === i) {
                    isActive = 'activeApprover';
                    notActive = true;
                }
                else {
                    isActive = 'beforeactiveApprover';
                }
            }
            else {
                isActive = 'overrideStage';
            }

            wf.push(
                <ul className="main-menu">
                    <li className={`${m.Role} ${isActive}`.trim()}>
                        {m.User}
                    </li>
                </ul>
            );
            //}
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
                .select("ID,Title,GLDescription,CostCenterDescription,RequestDate,Amount,GLCode,CostCentre,VoucherNumber,EmployeeName,Stage,ApprovalMatrix,WorkflowHistory,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus,ApprovalMatrix,VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,WorkflowHistory,AttachmentFiles,Status,PaymentRemarks,ApprovalRemarks,VendorCode/VendorCode,VendorCode/Title,VendorCode/Id,NatureofExpense/Title,NatureofExpense/Id,CurrentApprover/ID,CurrentApprover/EMail,CurrentApprover/Title")
                .expand("VendorCode,AttachmentFiles,NatureofExpense,CurrentApprover")();

            console.log(item);

            setItemID(itemId);
            setrequestNo(item.Title || "")
            setPayRemark(item.PaymentRemarks || "")

            setStatus(item.Status || "")
            // setRMVal(item.RM ?? "");
            // setLocation(item.Location || "");
            setRequestDate(item.RequestDate ? new Date(item.RequestDate).toLocaleDateString("en-GB") : "");
            // setEmployeeName(item.EmployeeName || "");
            // setDivison(item.Division || "");
            // setEmail(item.Email || "");
            // setHOD(item.HOD || "");
            // setDepartment(item.Department || "");
            // setEmpStatus(item.EmployeeStatus || "");
            // setcontactNo(item.ContactNo || "");
            setVendorCode(item.VendorCode.VendorCode || "");
            setVendorName(item.VendorName || "");
            //setNatureExpense(item.NatureofExpense.Title || "");
            // setSelectedNatureofexpense(item.NatureofExpense.Id);
            setSelectedNatureofexpense(item.NatureofExpense?.Id);
            setInvoiceNo(item.InvoiceNumber || "");
            const invDate = new Date(item.InvoiceDate || "").toISOString().split("T")[0];

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
            setGLDescription(item.GLDescription || "");
            setCostCenterDescription(item.CostCenterDescription || "");

            setOriginalData({
                Basicamount: item.Basicamount,
                GST: item.GST,
                OtherCharges: item.OtherCharges,
                Totalamount: item.Totalamount,
            });
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

            await loadFinalPOFile(itemId);

        } catch (error) {
            console.error("Error loading item:", error);
        }
    };



    const loadFinalPOFile = async (itemId: number) => {
        try {

            const sp = spfi().using(SPFx(props.currentSPContext));

            const folders = await sp.web.lists
                .getByTitle("NonPoDoc")
                .items
                .filter(`MainID eq '${itemId}'`)
                .select("FileRef", "FileLeafRef")
                ();

            if (!folders.length) {
                return;
            }

            const folderPath = folders[0].FileRef;

            console.log("Folder Path:", folderPath);

            // Fetch files inside folder
            const files = await sp.web
                .getFolderByServerRelativePath(folderPath)
                .files();

            console.log("Files Inside Folder:", files);

            const formattedFiles = files.map((file: any) => ({
                FileLeafRef: file.Name,
                FileRef: file.ServerRelativeUrl
            }));

            setPOAttachments({
                allFiles: formattedFiles
            });

        } catch (error) {
            console.error("Error loading files:", error);
        }
    };



    const toFolderName = (requestNo: string) => {
        return (requestNo || "NONPO")
            .replace(/[\\#%&*:{?<>|"]/g, "_") // removed /
            .replace(/\//g, "-")
            .trim();
    };

    const handleUpdate = async () => {
        const sp = spfi().using(SPFx(props.currentSPContext));
        if (!BasicAmount) {
            alert("Please Enter Basic Amount");
            return;
        }
        if (!InvoiceNo) {
            alert("Please Enter Invoice No");
            return;
        }
        if (!InvoiceDate) {
            alert("Please Select Invoice Date");
            return;
        }
        try {
            const item = await sp.web.lists
                .getByTitle("NonPO")
                .items.getById(ItemID)
                .select("*")();
            let stage = 1;

            let approverId = RMID.current;
            let pendingText = "Pending At RM";
            let Statuschanges = "Pending for Approval";
            let updatedHistory;
            if (Status == 'Send Back') {
                const changed = isDataChanged();
                let workflowHistory = [...WorkflowHistory];
                const lastEntry = workflowHistory[workflowHistory.length - 1];
                const lastApproverName = lastEntry?.CurrentApprover;
                const currentUserName = props.currentSPContext.pageContext.user.displayName || props.currentSPContext.pageContext.user.email;
                const performer = approverJson.current.find(x => x.Role === "APPerformer");
                const isSentBackByPerformer = performer && lastApproverName === performer.User;
                if (!changed) {
                    if (isSentBackByPerformer) {

                        const performer = approverJson.current.find(x => x.Role === "APPerformer");
                        if (performer) {
                            stage = approverJson.current.findIndex(x => x.Role === "APPerformer");
                            approverId = parseInt(performer.UserID);
                            pendingText = performer.PendingText || "Pending At AP Performer";
                            Statuschanges = "Pending for Acceptance";

                            const newHistoryEntry = {
                                CurrentApprover: props.currentSPContext.pageContext.user.displayName
                                    || props.currentSPContext.pageContext.user.email
                                    || '',
                                ActionTaken: 'Request submitted',
                                Comment: PayRemark,
                                Date: `${new Date()}`,
                                CurrentStatus: Statuschanges
                            };
                            workflowHistory.push(newHistoryEntry);
                            setWorkflowHistory(workflowHistory);
                            updatedHistory = JSON.stringify(workflowHistory);

                        }
                    }
                    else {
                        // 👉 Changed → restart from RM
                        stage = 1;
                        approverId = RMID.current;
                        pendingText = "Pending At RM";
                        Statuschanges = "Pending for Approval";
                        const newHistoryEntry = {
                            CurrentApprover: props.currentSPContext.pageContext.user.displayName
                                || props.currentSPContext.pageContext.user.email
                                || '',
                            ActionTaken: 'Request submitted',
                            Comment: '',
                            Date: `${new Date()}`,
                            CurrentStatus: Statuschanges
                        };

                        workflowHistory.push(newHistoryEntry);
                        setWorkflowHistory(workflowHistory);
                        updatedHistory = JSON.stringify(workflowHistory);
                    }
                }
                else {
                    // 👉 Changed → restart from RM
                    stage = 1;
                    approverId = RMID.current;
                    pendingText = "Pending At RM";
                    Statuschanges = "Pending for Approval";
                    const newHistoryEntry = {
                        CurrentApprover: props.currentSPContext.pageContext.user.displayName
                            || props.currentSPContext.pageContext.user.email
                            || '',
                        ActionTaken: 'Request submitted',
                        Comment: '',
                        Date: `${new Date()}`,
                        CurrentStatus: Statuschanges
                    };

                    workflowHistory.push(newHistoryEntry);
                    setWorkflowHistory(workflowHistory);
                    updatedHistory = JSON.stringify(workflowHistory);
                }

                await sp.web.lists
                    .getByTitle("NonPO")
                    .items.getById(ItemID)
                    .update({
                        Basicamount: BasicAmount.toString(),
                        GST: GSTAmt.toString(),
                        OtherCharges: otherCharges.toString(),
                        Totalamount: TotalAmount.toString(),
                        ApproverStatus: "Pending",
                        ApprovalMatrix: JSON.stringify(approverJson.current),
                        Status: Statuschanges,
                        WorkflowHistory: updatedHistory,
                        Stage: stage,
                        CurrentApproverId: approverId,
                        PendingAt: pendingText,
                        InvoiceNumber: InvoiceNo.toString(),
                        InvoiceDate: InvoiceDate,
                        PaymentRemarks: PayRemark,
                        NatureofExpenseId: selectedNatureofexpense
                        //  Stage: 1
                    });


                if (POTextAttachFile) {

                    const folderItem = await sp.web.lists
                        .getByTitle("NonPoDoc")
                        .items
                        .filter(`MainID eq '${ItemID}'`)
                        .select("FileRef")
                        .top(1)();

                    if (folderItem.length > 0) {

                        const folderPath = folderItem[0].FileRef;

                        const folder = sp.web.getFolderByServerRelativePath(folderPath);

                        const fileBuffer = await POTextAttachFile.arrayBuffer();

                        await folder.files.addUsingPath(
                            POTextAttachFile.name,
                            fileBuffer,
                            { Overwrite: true }
                        );
                    }
                }

                alert("SendBack request submitted successfully.");
                navigate('/');
            }

        }
        catch (error) {

            console.error("Error updating item:", error);
            alert("Update failed. Please try again.");

        }

    }

    const isDataChanged = () => {

        if (!originalData) return true;

        return (

            originalData.Basicamount !== BasicAmount ||
            originalData.GST !== GSTAmt ||
            originalData.OtherCharges !== otherCharges ||
            originalData.Totalamount !== TotalAmount
        );
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



    const deleteAttachment = async (file: any) => {

        if (!window.confirm("Are you sure you want to delete this file?")) {
            return;
        }

        try {

            const sp = spfi().using(SPFx(props.currentSPContext));

            // Delete file from document library folder
            await sp.web
                .getFileByServerRelativePath(file.FileRef)
                .delete();

            // Remove from UI
            setPOAttachments((prev: any) => ({
                ...prev,
                allFiles:
                    prev.allFiles?.filter(
                        (f: any) => f.FileRef !== file.FileRef
                    ) || []
            }));

            alert("File deleted successfully");

        } catch (error) {
            console.error("Delete failed:", error);
            alert("Failed to delete file");
        }
    };

    // ✅ UI must return something
    return (
        <>

            <div className="bordered">
                <a><img src={logo} /></a>
                <h1>Non PO Approval</h1>
            </div>
            <div className='displayWF'>{WorkflowJSX}</div>
            <div className='borderedbox'>
                <div className="heading1" style={{ marginTop: "10px" }}>
                    <label>Requestor Information</label>
                </div>
                <div className='main-formcontainer'>
                    <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Request Date </label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {RequestDate}</label>
                        </div>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Employee Name </label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {EmployeeName}</label>

                        </div>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Employee Email </label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {EmployeeEmail}</label>

                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Contact No</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {EmployeeContactno}</label>

                        </div>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Employee Status</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {EmployeeStatus}</label>

                        </div>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Division</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {EmployeeDivision}</label>

                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Location</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {Locationdata}</label>

                        </div>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>RM</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {EmployeeRM}</label>

                        </div>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>HOD</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {hoddata}</label>

                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Department</label> &nbsp; : &nbsp;
                            <label className='fonttext'>  {Department}</label>
                            {/* <Field readOnly name="Department" className="fonttext" /> */}
                        </div>

                        <div className='col-md-4 alignbox'>
                            <label className='font fontpadd'>Status</label> &nbsp; : &nbsp;
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
                            <label className='font'>Vendor Code : &nbsp;&nbsp;</label>
                            <input value={VendorCode} className='form-control readonly' />
                        </div>
                        <div className='col-md-4'>
                            <label className='font'>Vendor Name : &nbsp;&nbsp;</label>
                            <input value={VendorName} className='form-control readonly' />
                        </div>
                        <div className='col-md-4'>
                            <label className='font'>Nature of Expense : &nbsp;&nbsp;</label>
                            <Dropdown options={NatureOfExpenseTypeOptions} selectedKey={selectedNatureofexpense} onChange={(e, option) => { setSelectedNatureofexpense(option?.key as number); }} />
                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4'>
                            <label htmlFor="Employee Status" className='font fontblock'>Invoice No : &nbsp;&nbsp;</label>
                            <input id="invoiceNo" className="form-control" value={InvoiceNo} onChange={(e) => setInvoiceNo(e.target.value)} />
                        </div>
                        <div className='col-md-4'>
                            <label htmlFor="Location" className='font fontblock'>Invoice Date : &nbsp;&nbsp;</label>
                            <input id="invoicedate" className="form-control" value={InvoiceDate} onChange={(e) => setInvoiceDate(e.target.value)} />
                        </div>
                        <div className='col-md-4'>
                            <label htmlFor="RM" className='font fontblock'>Basic Amount : &nbsp;&nbsp;</label>
                            <input id="basicAmt" className="form-control" type="number" value={BasicAmount} onChange={handleAmountChange("BasicAmount")} />
                        </div>
                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4'>
                            <label htmlFor="RM" className='font fontblock'>GST Amount : &nbsp;&nbsp;</label>
                            <input id="gstAmt" className="form-control" type="number" value={GSTAmt} onChange={handleAmountChange("GSTAmt")} />
                        </div>
                        <div className='col-md-4'>
                            <label htmlFor="RM" className='font fontblock'>Other Charges (If any) : &nbsp;&nbsp;</label>
                            <input id="othercharge" className="form-control" type="number" value={otherCharges} onChange={handleAmountChange("OtherCharges")} />
                        </div>
                        <div className='col-md-4'>
                            <label htmlFor="RM" className='fontfontblock '>Total Amount : &nbsp;&nbsp;</label>
                            <input id="totAmt" className="form-control" type="text" value={TotalAmount} />
                        </div>

                    </div>
                    <div className='row mb-20'>
                        <div className='col-md-4'>
                            <label htmlFor="RM" className='font fontblock'>Remarks for NON-PO Expense : &nbsp;&nbsp;</label>
                            <input id="othercharge" className="form-control" type="text" value={PayRemark} onChange={(e) => setPayRemark(e.target.value)} />
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
                    <div className='row mb-20'>
                        <div className="col-md-4">
                            <label className='font'>Attachment </label>


                            <div>
                                {POAttachments.allFiles?.length > 0 &&
                                    POAttachments.allFiles.map((file: any, index: number) => (
                                        <div
                                            key={index}
                                            className="d-flex align-items-center gap-2 mt-1"
                                        >
                                            <a
                                                href={file.FileRef}
                                                target="_blank"
                                                rel="noopener noreferrer"
                                            >
                                                {file.FileLeafRef}
                                            </a>

                                            <span
                                                style={{
                                                    color: "red",
                                                    fontWeight: "bold",
                                                    cursor: "pointer",
                                                    fontSize: "18px"
                                                }}
                                                onClick={() => deleteAttachment(file)}
                                            >
                                                ×
                                            </span>
                                        </div>
                                    ))}

                                <Upload
                                    beforeUpload={(file) => {
                                        setPOTextAttachFile(file);
                                        return false;
                                    }}
                                    onRemove={() => {
                                        setPOTextAttachFile(null);
                                    }}
                                    maxCount={1}
                                    className="upload-full-width"
                                >
                                    <Button
                                        className="upload-btn-full"
                                        icon={<UploadOutlined />}
                                        iconPosition="end"
                                    />
                                </Upload>
                                {POTextAttachFile && (
                                    <div className="mt-1 text-success">
                                        Selected: {POTextAttachFile.name}
                                    </div>
                                )}
                            </div>
                        </div>


                        {/* <div className='col-md-12'>
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
                                                    <td style={{ ...td }}>setInvoiceNo
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
                        </div> */}
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
                    <a className="submit-btn" onClick={() => handleUpdate()}>Submit</a>
                    <a className="reset-btn" onClick={() => navigate(-1)}>Exit</a>
                </div>
            </div>
        </>

    );
};
