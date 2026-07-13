import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import { Formik, Form, Field, FormikProps } from "formik";
import useFullscreenForm from "../../hook/useFullscreenForm";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@fortawesome/fontawesome-free/css/all.min.css";

import logo from "../../assets/sona-comstarlogo.png";

import {
  useLocation,
  useNavigate,
  useSearchParams
} from "react-router-dom";


// PnPjs 
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";

// import { useNavigate, useSearchParams } from "react-router-dom";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";

/** ===================== SITE & LISTS ===================== */
const SITE_ABS_URL = "https://isriglobal.sharepoint.com/sites/SonaNONPO";
const NONPO_LIST_TITLE = "NonPO";
const APPLIATION_LIST_TITLE = "ApplicationNumber";

const VENDOR_MASTER = "VendorMaster";
const NOE_MASTER = "ExpenseMaster";
const COUNTRY_MASTER = "CountryMaster";
const EMPLIST = "EmployeeMaster";

/** ===================== ATTACHMENT LIMIT ===================== */
const ATTACH_LIB: string = "NonPODoc";
const MAX_FILE_SIZE_MB = 25;
const MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024;

/** ===================== Helpers ===================== */
const toISO = (d: Date) => d.toISOString().split("T")[0];
const toISOInput = (d: any): string => {
  if (!d) return "";
  const dt = new Date(d);
  if (Number.isNaN(dt.getTime())) return "";
  return dt.toISOString().split("T")[0];
};
const safeNum = (s: string): number => {
  const n = Number((s || "0").toString().replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
};

const fyFromDate = (d: Date): string => {
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const yy = y % 100;
  if (m >= 4) {
    const next = (yy + 1) % 100;
    return `${String(yy).padStart(2, "0")}-${String(next).padStart(2, "0")}`;
  } else {
    const prev = (yy - 1 + 100) % 100;
    return `${String(prev).padStart(2, "0")}-${String(yy).padStart(2, "0")}`;
  }
};
const tryParseLastNumber = (title: string): number => {
  if (!title) return 0;
  const parts = title.split("/");
  const last = parts[parts.length - 1] || "";
  const n = parseInt(last, 10);
  return isNaN(n) ? 0 : n;
};
const buildRequestId = (fy: string, seq: number): string =>
  `NON PO/${fy}/${String(seq).padStart(5, "0")}`;

// const loadApprovalRoles = async (spx: SPFI) => {
//   const rows = await spx.web.lists.getByTitle("NONPOApprovalMatrix")
//     .items.select("Title", "Role", "Approver/Id", "Approver/Title", "Approver/EMail")
//     .expand("Approver")();

//   return rows.map(r => ({
//     Role: r.Role,
//     User: r.Approver?.Title || "",
//     UserID: r.Approver?.Id || null,
//     UserEmail: r.Approver?.EMail || "",
//     PendingText:
//       r.Role === "CFO" ? "Pending At Level 2 CEO" :
//         r.Role === "Performer" ? "Pending Level 3 Performer"
//           : ""
//   }));
// };



// const loadApprovalRoles = async (spx: SPFI) => {
//   const rows = await spx.web.lists.getByTitle("NONPOApprovalMatrix")
//     .items.select(
//       "Id",
//       "Title",
//       "ApprovalType",
//       "Status",
//       "Role/Id",
//       "Role/Title",
//       "Level/Id",
//       "Level/Title",
//       "Approver/Id",
//       "Approver/Title",
//       "Approver/EMail"
//     )
//     .expand("Role", "Level", "Approver")();

//   return rows.map(r => ({
//     Role: r.Role?.Title || "",
//     Level: r.Level?.Title || "",
//     User: r.Approver?.Title || "",
//     UserID: r.Approver?.Id || null,
//     UserEmail: r.Approver?.EMail || "",
//     PendingText:
//       r.Role?.Title === "CFO"
//         ? "Pending At Level 2 CEO"
//         : r.Role?.Title === "APPerformer" || r.Role?.Title === "Performer"
//           ? "Pending Level 3 Performer"
//           : ""
//   }));
// };




// Approval Matrix

// const loadApprovalRoles = async (spx: SPFI) => {
//   const rows = await spx.web.lists
//     .getByTitle("NONPOApprovalMatrix")
//     .items.select(
//       "Id",
//       "Title",
//       "ApprovalType",
//       "Status",
//       "Role/Id",
//       "Role/RoleName",
//       "Approver/Id",
//       "Approver/Title",
//       "Approver/EMail"
//     )
//     .expand("Role", "Approver")
//     .filter("Status eq 'Active'")
//     ();

//   return rows.map(r => {
//     const role = (r.Role?.RoleName || "").toLowerCase();

//     let pendingText = "";
//     if (role.includes("cfo")) {
//       pendingText = "Pending At  CFO";
//     }
//     else if (
//       role.includes("apperformer") ||
//       role.includes("ap performer") ||
//       role.includes("performer")
//     ) {
//       pendingText = "Pending At AP Performer";
//     }

//     return {
//       Role: r.Role?.RoleName || "",
//       User: r.Approver?.Title || "",
//       UserID: r.Approver?.Id || null,
//       UserEmail: r.Approver?.EMail || "",
//       PendingText: pendingText
//     };
//   });
// };



// const loadApprovalRoles = async (spx: SPFI) => {
//   const rows = await spx.web.lists
//     .getByTitle("NONPOApprovalMatrix")
//     .items
//     .select(
//       "Id",
//       "Title",
//       "ApprovalType",
//       "Status",
//       "Role/Id",
//       "Role/RoleName",
//       "Approver/Id",
//       "Approver/Title",
//       "Approver/EMail"
//     )
//     .expand("Role", "Approver")
//     .filter("Status eq 'Active'")();

//   // ✅ Create Approval Matrix JSON
//   const approvalMatrixJSON = rows.map((r: any) => {
//     const role = (r.Role?.RoleName || "").toLowerCase();

//     let pendingText = "";

//     if (role.includes("cfo")) {
//       pendingText = "Pending At CFO";
//     } else if (
//       role.includes("apperformer") ||
//       role.includes("ap performer") ||
//       role.includes("performer")
//     ) {
//       pendingText = "Pending At AP Performer";
//     }

//     return {
//       Role: r.Role?.RoleName || "",
//       Approver: r.Approver?.Title || "",
//       ApproverID: r.Approver?.Id?.toString() || "",
//       ApproverEmail: r.Approver?.EMail || "",
//       PendingText: pendingText
//     };
//   });

//   // ✅ Store in ref
//   approvalMatrix.current = approvalMatrixJSON;

//   return approvalMatrixJSON;
// };



// const loadApprovalRoles = async (spx: SPFI) => {
//   const rows = await spx.web.lists
//     .getByTitle("NONPOApprovalMatrix")
//     .items
//     .select(
//       "Id",
//       "Title",
//       "ApprovalType",
//       "Status",
//       "Role/Id",
//       "Role/RoleName",
//       "Approver/Id",
//       "Approver/Title",
//       "Approver/EMail"
//     )
//     .expand("Role", "Approver")
//     .filter("Status eq 'Active'")
//     ();

//   return rows.map(r => {
//     const role = (r.Role?.RoleName || "").toLowerCase();

//     let pendingText = "";
//     if (role.includes("cfo")) {
//       pendingText = "Pending At CFO";
//     } else if (
//       role.includes("apperformer") ||
//       role.includes("ap performer") ||
//       role.includes("performer")
//     ) {
//       pendingText = "Pending At AP Performer";
//     }

//     return {
//       Role: r.Role?.RoleName || "",
//       User: r.Approver?.Title || "",
//       UserID: r.Approver?.Id || null,
//       UserEmail: r.Approver?.EMail || "",
//       PendingText: pendingText
//     };
//   });
// };


interface IVendorOption {
  id: number;
  display: any;
  VendorName?: string;
  title?: string;
  VendorCode?: string;
}
interface ILookupOption {
  id: number;
  title: string;
}

interface IExistingFile {
  name: string;
  url: string;
  sizeBytes: number;
  lastModified?: string;
  source: "Library" | "ListAttachment";
  serverRelativeUrl?: string;
  attachmentFileName?: string;
}

/** ===================== Form Values ===================== */
interface INonPoFormValues {
  Title: string;

  RequestDate: string;
  EmployeeCode: string;
  EmployeeName: string;
  Email: string;
  ContactNo: string;
  Division: string;
  Department: string;
  Location: string;
  RM: string;
  HOD: string;
  EmployeeStatus: string;

  VendorCode: number | null;
  VendorName: string;
  NatureofExpense: number | null;

  InvoiceNumber: string;
  InvoiceDate: string;
  Basicamount: string;
  GST: string;
  OtherCharges: string;
  Totalamount: string;

  PaymentRemarks: string;
  ApprovalRemarks: string;
  GLCode: string;


  VoucherNumber: string;
  Amount: string;

  attachments: File[];

  Status?: string;
}

const NonPoRequest: React.FC<ISonanonpoprodProps> = (props) => {

  //  const location = useLocation();
  const navigate = useNavigate();

  const [currentDate, setCurrentDate] = React.useState("");
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

  useFullscreenForm();
  // const navigate = useNavigate();
  const [params] = useSearchParams();

  /** ------------ Router params ------------ */
  const idParam = params.get("id");
  const modeParam = (params.get("mode") || "view").toLowerCase();
  const numericId = useMemo(() => {
    const n = Number(idParam);
    return Number.isFinite(n) ? n : undefined;
  }, [idParam]);
  const isNew = modeParam === "new" || !numericId;
  const isEdit = modeParam === "edit" || isNew;
  const modeLabel = isNew ? "New" : isEdit ? "Edit" : "View";

  /** ------------ SP instance ------------ */
  const [sp, setSp] = useState<SPFI | null>(null);

  useEffect(() => {
    if (!props.currentSPContext) return;
    setSp(spfi().using(SPFx(props.currentSPContext)));
  }, [props.currentSPContext]);

  /** ------------ Masters & UI ------------ */
  const [vendors, setVendors] = useState<IVendorOption[]>([]);
  const [noeOptions, setNoeOptions] = useState<ILookupOption[]>([]);
  const [ccOptions, setCcOptions] = useState<ILookupOption[]>([]);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");
  const [WorkflowJSX, setWorkflowJSX] = React.useState<JSX.Element[] | null>(null);
  const approverJson = React.useRef<any[]>([]);
  // const [Employeemasterdata, setEmployeeMasterdata] = React.useState<any[]>([]);
  const EmployeeQmadata = React.useRef<any[]>([]);
  const currentUserId = React.useRef(0);

  const [NextNo, setNextNo] = useState<any>("");

  const [VendorLookupID, setVendorLookupID] = useState<number | null>(null);
  const [VendorNameValue, setVendorNameValue] = React.useState<string>();
  const [IncrimentalId, setIncrimentalId] = useState<any>("");

  // let EmployeeQmadata = React.useRef<any[]>(null);

  const [HOD, setHOD] = React.useState<string>();
  const [EmpStatus, setEmpStatus] = React.useState<string>();

  const [ApprovalMatrixdata, setApprovalMatrix] = React.useState<any[]>([]);
  //const approvalMatrix = React.useState<any[]>([]);// useRef<any[]>(null);

  const approvalMatrix = React.useRef<any[]>([]);
  const [Stage, setStageData] = React.useState(0);

  /** ------------ Existing documents------------ */
  const [existingFiles, setExistingFiles] = useState<IExistingFile[]>([]);
  const [loadingExisting, setLoadingExisting] = useState(false);

  /** ------------ Requester defaults------------ */
  const [profileLoaded, setProfileLoaded] = useState(false);
  // const [reqDefaults, setReqDefaults] = useState({
  //   EmployeeName: props.currentSPContext?.pageContext?.user?.displayName || "",
  //   Email: props.currentSPContext?.pageContext?.user?.email || "",
  //   Department: "",
  //   Division: "",
  //   Location: "",
  //   ContactNo: "",
  //   EmployeeCode: "",
  //   RM: "",
  //   HOD: "",
  //   EmployeeStatus: "Existing",
  // });

  const [reqDefaults, setReqDefaults] = useState({
    EmployeeName: "",
    Email: "",
    Department: "",
    Division: "",
    Location: "",
    ContactNo: "",
    EmployeeCode: "",
    RM: "",
    RMEmail: "",
    HOD: "",
    HODEmail: "",
    EmployeeStatus: "Existing",
  });

  // React.useEffect(() => {
  //   getdata();
  // }, []);



  //SNEHAL
  let getdata = async (userEmail: string) => {
    // await userProfile();
    // await fetchEmployeeMaster();
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

      const FLOW_URL =
        "https://defaultcb1edbfe8080457d9cae51528f3643.3f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e2bb522aa41443179a72b701b9613471/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=q8b8ADCtK2eKr2f6p3MX7gxmJymPeJbm0mq2M69Rk8E";

      const fetchPage = async (pageNumber: number) => {

        const response = await fetch(FLOW_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            PageSize: 500,
            PageNumber: pageNumber
          })
        });

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
          RMName: item.reportingManagerName === "Piyush Airan" ? "Prince Gupta" : item.reportingManagerName || "",
          RMEmail: item.reportingManagerEmail === "piyush.airan@sonacomstar.com" ? "prince.gupta@sonacomstar.com" : item.reportingManagerEmail || "",
          HODName: hodAttr?.attributeTypeUnitDescription === "Piyush Airan" ? "Prince Gupta" : hodAttr?.attributeTypeUnitDescription || "",
          HODEmail: hodemailAttr?.attributeTypeUnitDescription === "piyush.airan@sonacomstar.com" ? "prince.gupta@sonacomstar.com" : hodemailAttr?.attributeTypeUnitDescription || "",

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

  // let getapprovalmatrix = async () => {

  //   const approverArray: any[] = [];

  //   // =====================================================
  //   // ✅ BUILD APPROVER MATRIX JSON
  //   // =====================================================

  //   //const approverJson: any[] = [];


  //   // ✅ Get emails from Employeemasterdata

  //   const matchedEmployee = EmployeeQmadata.current.find(
  //     emp => String(emp.EmployeeId) === String(currentUserId.current)
  //   );


  //   approverArray.push({
  //     Role: "Initiator",
  //     User: matchedEmployee.Employee.Title,
  //     UserID: matchedEmployee.Employee.Id,
  //     UserEmail: matchedEmployee.Employee.EMail
  //   });

  //   approverArray.push({
  //     Role: "RM",
  //     User: matchedEmployee.ReportingManager.Title,
  //     UserID: matchedEmployee.ReportingManager.ID,
  //     UserEmail: matchedEmployee.ReportingManager.EMail,
  //     PendingText: "Pending At RM"
  //   });

  //   approverArray.push({
  //     Role: "HOD",
  //     User: matchedEmployee.HOD.Title,
  //     UserID: matchedEmployee.HOD.ID,
  //     UserEmail: matchedEmployee.HOD.EMail,
  //     PendingText: "Pending At HOD"
  //   });


  //   // Add other approvers from ApprovalMatrix list
  //   approvalMatrix.current.forEach((item: any) => {

  //     const role = item.Role || "";

  //     const approverId = item.ApproverID?.toString();

  //     if (!approverId) return;

  //     if (role === "RM" || role === "HOD") return;

  //     let pendingText = "";

  //     if (role === "CFO")
  //       pendingText = "Pending At CFO ";

  //     else if (role === "APPerformer")
  //       pendingText = "Pending At APPerformer ";

  //     const alreadyExists =
  //       approverArray.some(x => x.ApproverID === approverId);
  //     if (alreadyExists) return;

  //     approverArray.push({
  //       Role: role,
  //       User: item.Approver || "",
  //       UserID: approverId,
  //       UserEmail: item.ApproverEmail || "",
  //       PendingText: pendingText
  //     });
  //   });
  //   approverJson.current = approverArray;
  // }

  // const fetchEmployeeMaster = async () => {
  //   try {
  //     const sp = spfi().using(SPFx(props.currentSPContext));

  //     const Empitems = await sp.web.lists
  //       .getByTitle("EmployeeMaster")
  //       .items
  //       .select(
  //         "*",
  //         "Employee/Title",
  //         "Employee/Id",
  //         "Employee/EMail",
  //         "ReportingManager/ID",
  //         "ReportingManager/Title",
  //         "ReportingManager/EMail",
  //         "HOD/ID",
  //         "HOD/Title",
  //         "HOD/EMail",
  //         "EmployeeEmail"
  //       )
  //       .expand("Employee", "ReportingManager", "HOD")();
  //     EmployeeQmadata.current = Empitems;
  //     setEmployeeMasterdata(Empitems);

  //   } catch (error) {
  //     console.error("Error fetching EmployeeMaster:", error);
  //   }
  // };



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

  useEffect(() => {
    (async () => {
      if (!sp) return;
      try {
        const me = await sp.web.currentUser();
        const myId: number | undefined = me?.Id;
        const myEmail = (me?.Email || props.currentSPContext?.pageContext?.user?.email || "").trim().toLowerCase();

        let items: any[] = [];
        // by EmployeeId
        try {
          if (myId) {
            items = await sp.web.lists
              .getByTitle(EMPLIST)
              .items.select(
                "ID", "Title", "EmployeeName", "EmployeeEmail", "ContactNo", "EmployeeStatus",
                "Division", "Location", "Department", "EmployeeCode", "Status",
                "ReportingManager/Title", "ReportingManager/EMail", "HOD/Title", "HOD/EMail"
              )
              .expand("ReportingManager", "HOD")
              .filter(`EmployeeId eq ${myId}`)
              .top(5)();
          }
        } catch (Error) { alert("Error") }
        // by EmployeeEmail
        if (!items?.length) {
          try {
            const esc = (s: string) => s.replace(/'/g, "''");
            items = await sp.web.lists
              .getByTitle(EMPLIST)
              .items.select(
                "ID", "Title", "EmployeeName", "EmployeeEmail", "ContactNo", "EmployeeStatus",
                "Division", "Location", "Department", "EmployeeCode", "Status",
                "ReportingManager/Title", "ReportingManager/EMail", "HOD/Title", "HOD/EMail"
              )
              .expand("ReportingManager", "HOD")
              .filter(`EmployeeEmail eq '${esc(myEmail)}'`)
              .top(5)();
          } catch (Error) { alert("Error") }
        }

        const row = items?.[0];
        if (row) {
          setReqDefaults((prev) => ({
            ...prev,
            EmployeeName: row.EmployeeName || row.Title || prev.EmployeeName,
            Email: (row.EmployeeEmail || prev.Email || "").toLowerCase(),
            ContactNo: row.ContactNo || "",
            EmployeeStatus: row.EmployeeStatus || "Existing",
            Division: row.Division || "",
            Location: row.Location || "",
            Department: row.Department || "",
            EmployeeCode: row.EmployeeCode || "",

            RM: row?.ReportingManager?.Title || "",
            RMEmail: row?.ReportingManager?.EMail || "",
            HOD: row?.HOD?.Title || "",
            HODEmail: row?.HOD?.EMail || "",

          }));
        }
      } catch (Error) { alert("Error") }
      finally { setProfileLoaded(true); }
    })();
  }, [sp, props.currentSPContext]);



  /** ------------ Base initial values (used in new mode) ------------ */
  const baseInitialValues: INonPoFormValues = useMemo(
    () => ({
      Title: "NON-PO Request",
      RequestDate: toISO(new Date()),
      EmployeeCode: reqDefaults.EmployeeCode,
      EmployeeName: reqDefaults.EmployeeName,
      Email: reqDefaults.Email,
      ContactNo: reqDefaults.ContactNo,
      Division: reqDefaults.Division,
      Department: reqDefaults.Department,
      Location: reqDefaults.Location,
      RM: reqDefaults.RM,
      HOD: reqDefaults.HOD,
      EmployeeStatus: reqDefaults.EmployeeStatus,

      VendorCode: null,
      VendorName: "",
      NatureofExpense: null,

      InvoiceNumber: "",
      InvoiceDate: "",
      Basicamount: "",
      GST: "",
      OtherCharges: "",
      Totalamount: "",
      PaymentRemarks: "",
      ApprovalRemarks: "",
      GLCode: "",


      VoucherNumber: "",
      Amount: "",

      attachments: [],

      // NEW
      Status: "",
    }),
    [reqDefaults]
  );

  const [initialValues, setInitialValues] = useState<INonPoFormValues>(baseInitialValues);

  useEffect(() => {
    if (!numericId) setInitialValues(baseInitialValues);
  }, [baseInitialValues, numericId]);

  /** ------------ Masters ------------ */
  useEffect(() => {
    (async () => {
      if (!sp) return;
      try {
        const vItems = await sp.web.lists
          .getByTitle(VENDOR_MASTER)
          .items.select("ID,Title,VendorCode,VendorName,Status")
          .filter(`(Status eq 'Active') or (Status eq 'active')`)
          .top(5000)();
        const opts: IVendorOption[] = (vItems || []).map((v: any) => {
          const code = v.VendorCode || v.Title || "";
          // const code = v.VendorCode;
          const VendorName = v.VendorName || v.Title || "";
          // const display = [code, VendorName].filter(Boolean).join(" — ");
          const display = [VendorName];

          return { id: v.ID, VendorCode: code, VendorName, title: v.Title, display: display || `#${v.ID}` };


        });
        setVendors(opts);
      } catch { setVendors([]); }

      if (NOE_MASTER) {
        try {
          const nItems = await sp.web.lists.getByTitle(NOE_MASTER)
            .items.select("ID,Title,Status,ExpenseType")
            .filter(`(Status eq 'Active') or (Status eq 'active')`).top(5000)();
          setNoeOptions((nItems || []).map((x: any) => ({ id: x.ID, title: x.ExpenseType })));
        } catch { setNoeOptions([]); }
      }

      if (COUNTRY_MASTER) {
        try {
          const cItems = await sp.web.lists.getByTitle(COUNTRY_MASTER)
            .items.select("ID,Title,Status")
            .filter(`(Status eq 'Active') or (Status eq 'active')`).top(5000)();
          setCcOptions((cItems || []).map((x: any) => ({ id: x.ID, title: x.Title })));
        } catch { setCcOptions([]); }
      }
    })();
  }, [sp]);

  /** ------------ prefill  item ------------ */
  useEffect(() => {
    (async () => {
      if (!sp || !numericId) return;
      try {
        setLoading(true); setErr("");
        const it = await sp.web.lists.getByTitle(NONPO_LIST_TITLE).items
          .getById(numericId)
          .select(
            "Id,Title,RequestDate,EmployeeCode,EmployeeName,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus," +
            "VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,PaymentRemarks,ApprovalRemarks,GLCode," +
            "VoucherNumber,Amount,Status," +
            "VendorCode/Id,VendorCode/Title,NatureofExpense/Id,NatureofExpense/Title"
          )
          .expand("VendorCode", "NatureofExpense")();

        setInitialValues({
          Title: it?.Title || "NON-PO Request",
          RequestDate: toISOInput(it?.RequestDate) || toISO(new Date()),
          EmployeeCode: it?.EmployeeCode || "",
          EmployeeName: it?.EmployeeName || "",
          Email: it?.Email || "",
          ContactNo: it?.ContactNo || "",
          Division: it?.Division || "",
          Department: it?.Department || "",
          Location: it?.Location || "",
          RM: it?.RM || "",
          HOD: it?.HOD || "",
          EmployeeStatus: it?.EmployeeStatus || "Existing",

          VendorCode: it?.VendorCode?.Id ?? null,
          VendorName: it?.VendorName || "",
          NatureofExpense: it?.NatureofExpense?.Id ?? null,

          InvoiceNumber: it?.InvoiceNumber || "",
          InvoiceDate: toISOInput(it?.InvoiceDate) || "",
          Basicamount: it?.Basicamount || "",
          GST: it?.GST || "",
          OtherCharges: it?.OtherCharges || "",
          Totalamount: it?.Totalamount || "",
          PaymentRemarks: it?.PaymentRemarks || "",
          ApprovalRemarks: it?.ApprovalRemarks || "",
          GLCode: it?.GLCode || "",


          VoucherNumber: it?.VoucherNumber || "",
          Amount: it?.Amount || it?.Totalamount || "",

          attachments: [],

          // NEW
          Status: it?.Status || "",
        });
      } catch (e: any) {
        console.error("[NonPoRequest] prefill fetch failed:", e);
        setErr(e?.message || "Failed to load item");
      } finally { setLoading(false); }
    })();
  }, [sp, numericId]);

  /** ------------ Load existing documents ------------ */
  const toFolderName = (requestNo: string) => {
    return (requestNo || "NONPO")
      .replace(/[\\#%&*:{?<>|"]/g, "_") // removed /
      .replace(/\//g, "-")
      .trim();
  };
  const getLibraryFolderPathForRequest = async (spx: SPFI, libTitle: string, requestNo: string) => {
    const root = await spx.web.lists.getByTitle(libTitle).rootFolder();
    const rootUrl: string = root.ServerRelativeUrl;
    const folderName = toFolderName(requestNo);
    const folderPath = `${rootUrl}/${folderName}`;
    return { folderPath, folderName };
  };

  useEffect(() => {
    (async () => {
      if (!sp || !numericId) return;
      setLoadingExisting(true);
      try {
        // Title (Request No)
        const meta = await sp.web.lists.getByTitle(NONPO_LIST_TITLE)
          .items.getById(numericId)
          .select("Title")();
        const reqNo = meta?.Title || "NONPO";

        // Library files under /<RequestNo>/
        const libFiles: IExistingFile[] = [];
        try {
          const { folderPath } = await getLibraryFolderPathForRequest(sp, ATTACH_LIB, reqNo);
          const files = await sp.web.getFolderByServerRelativePath(folderPath)
            .files.select("Name,ServerRelativeUrl,Length,TimeLastModified")();

          for (const f of files || []) {
            const url = `${SITE_ABS_URL}${f.ServerRelativeUrl}`;
            libFiles.push({
              name: f.Name,
              url,
              sizeBytes: Number(f.Length || 0),
              lastModified: f.TimeLastModified,
              source: "Library",
              serverRelativeUrl: f.ServerRelativeUrl,
            });
          }
        } catch (Error) { alert("Error") }

        setExistingFiles([...libFiles]);
      } catch (e) {
        console.warn("[NONPO] load existing docs failed:", e);
        setExistingFiles([]);
      } finally {
        setLoadingExisting(false);
      }
    })();
  }, [sp, numericId]);

  /** ------------ Form helpers ------------ */
  const formikRef = useRef<FormikProps<INonPoFormValues>>(null);

  const updateTotals = () => {
    const f = formikRef.current; if (!f) return;
    const total = safeNum(f.values.Basicamount) + safeNum(f.values.GST) + safeNum(f.values.OtherCharges);
    f.setFieldValue("Totalamount", total.toString());
    f.setFieldValue("Amount", total.toString());
  };
  const onVendorChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sel = Number(e.target.value) || null;
    const f = formikRef.current; if (!f) return;
    f.setFieldValue("VendorName", sel);
    const chosen = vendors.find((v) => v.id === sel);
    //  f.setFieldValue("VendorName", chosen?.VendorName || chosen?.title || "");
    //  f.setFieldValue("VendorCode", chosen?.VendorCode || chosen?.title || "");

    f.setFieldValue("VendorCode", chosen?.VendorCode || "");
    setVendorLookupID(chosen?.id || null);
    setVendorNameValue(chosen?.VendorName || '')

    // IMPORTANT:
    // For SharePoint lookup field use InternalName + "Id"
    f.setFieldValue("VendorCodeId", chosen?.id || null);
    // Save lookup ID separately
    f.setFieldValue("VendorCodeId", chosen?.id || null);
  };
  const onNoeChange = (e: React.ChangeEvent<HTMLSelectElement>) =>
    formikRef.current?.setFieldValue("NatureofExpense", Number(e.target.value) || null);


  const getCurrentUserId = async (spx: SPFI): Promise<number | undefined> => {
    try {
      const me = await spx.web.currentUser();
      return me?.Id;
    } catch {
      return undefined;
    }
  };

  const ensureRequestFolder = async (spx: SPFI, libTitle: string, requestNo: string, mainId: number) => {
    const root = await spx.web.lists.getByTitle(libTitle).rootFolder();
    const rootUrl: string = root.ServerRelativeUrl;
    const folderName = toFolderName(requestNo);
    const folderPath = `${rootUrl}/${folderName}`;

    try {
      await spx.web.folders.addUsingPath(folderPath);
    } catch (Error) { alert("Error") }

    const folderItem = await spx.web
      .getFolderByServerRelativePath(folderPath)
      .getItem();

    await folderItem.update({
      MainID: mainId.toString()
    });

    return { folderPath, folderName };
  };

  const uploadAttachmentsToLibrary = async (
    spx: SPFI,
    libTitle: string,
    requestNo: string,
    itemId: number,
    files: File[],
    uploaderRole: string = "Requester"
  ) => {
    if (!files?.length) return;

    const { folderPath } = await ensureRequestFolder(spx, libTitle, requestNo, itemId);
    const userId = await getCurrentUserId(spx);

    for (const file of files) {
      await spx.web.getFolderByServerRelativePath(folderPath).files.addChunked(file.name, file, undefined);

      try {
        const uploaded = await spx.web.getFileByServerRelativePath(`${folderPath}/${file.name}`);
        const item = await uploaded.getItem();
        await item.update({
          RequestNo: requestNo,
          NonPOItemId: itemId,
          MainID: itemId.toString(),
          DocUploadedByRole: uploaderRole,
        });
      } catch (e) {
        console.warn("[NonPO] file metadata update skipped:", e);
      }
    }
  };

  const blockMinusAndAlpha: React.KeyboardEventHandler<HTMLInputElement> = (e) => {
    const badKeys = ["-", "e", "E", "+"];
    if (badKeys.includes(e.key)) {
      e.preventDefault();
    }
  };

  const onAmountPaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
    const txt = e.clipboardData.getData("text");
    if (!txt) return;

    const trimmed = txt.trim();

    if (trimmed === "-" || /^-/.test(trimmed)) {
      e.preventDefault();
      alert("Minus ( - ) not allowed. Please enter a non-negative amount.");
      return;
    }

    const cleaned = trimmed.replace(/[^\d.]/g, "");
    if (cleaned !== trimmed) {
      e.preventDefault();

      const input = e.target as HTMLInputElement;

      const start: number = input.selectionStart ?? input.value.length;
      const end: number = input.selectionEnd ?? input.value.length;

      const old = input.value;
      const next = old.slice(0, start) + cleaned + old.slice(end);

      input.value = next;
      input.dispatchEvent(new Event("input", { bubbles: true }));

      const newCaret = start + cleaned.length;
      input.setSelectionRange(newCaret, newCaret);
    }
  };

  // const recalcTotal = (
  //   vals: INonPoFormValues,
  //   overrides?: Partial<Pick<INonPoFormValues, "Basicamount" | "GST" | "OtherCharges">>
  // ): string => {
  //   const basic = safeNum((overrides?.Basicamount ?? vals.Basicamount) as string);
  //   const gst = safeNum((overrides?.GST ?? vals.GST) as string);
  //   const other = safeNum((overrides?.OtherCharges ?? vals.OtherCharges) as string);
  //   return String(basic + gst + other);
  // };

  const recalcTotal = (
    vals: INonPoFormValues,
    overrides?: Partial<Pick<INonPoFormValues, "Basicamount" | "GST" | "OtherCharges">>
  ): string => {

    const basic = Math.round(
      safeNum((overrides?.Basicamount ?? vals.Basicamount) as string) * 100
    );

    const gst = Math.round(
      safeNum((overrides?.GST ?? vals.GST) as string) * 100
    );

    const other = Math.round(
      safeNum((overrides?.OtherCharges ?? vals.OtherCharges) as string) * 100
    );

    const total = (basic + gst + other) / 100;

    return total.toFixed(2);
  };

  const handleAmountChange = (name: "Basicamount" | "GST" | "OtherCharges") =>
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = formikRef.current;
      if (!f) return;

      const raw = (e.target.value ?? "").toString().trim();

      if (raw === "-" || /^-/.test(raw)) {
        alert("Minus ( - ) not allowed . Please enter a non-negative amount.");
        f.setFieldValue(name, "", false);
        const nextTotal = recalcTotal(f.values, { [name]: "" } as any);
        f.setFieldValue("Totalamount", nextTotal, false);
        f.setFieldValue("Amount", nextTotal, false);
        return;
      }

      const cleaned = raw.replace(/[^\d.]/g, "");
      // Allow only one decimal point and maximum 2 digits after decimal
      let normalized = cleaned;

      if (!/^\d*\.?\d{0,2}$/.test(cleaned)) {
        return; // Ignore invalid input
      }


      const parts = cleaned.split(".");
      if (parts.length > 2) {
        normalized = `${parts[0]}.${parts[1]}`;
      }
      // const normalized = parts.length > 2 ? `${parts[0]}.${parts.slice(1).join("")}` : cleaned;

      f.setFieldValue(name, normalized, false);

      const nextTotal = recalcTotal(f.values, { [name]: normalized } as any);
      f.setFieldValue("Totalamount", nextTotal, false);
      f.setFieldValue("Amount", nextTotal, false);
    };

  const blockNonDigitForInvoice: React.KeyboardEventHandler<HTMLInputElement> = (e) => {
    const badKeys = ["-", "e", "E", "+", "."];
    if (badKeys.includes(e.key)) {
      e.preventDefault();
    }
  };

  const onInvoicePaste: React.ClipboardEventHandler<HTMLInputElement> = (e) => {
    const txt = e.clipboardData.getData("text");
    if (!txt) return;

    const trimmed = txt.trim();

    if (trimmed === "-" || /^-/.test(trimmed)) {
      e.preventDefault();
      alert("Invoice No. me minus ( - ) not allowed . Only positive values are allowed.");
      return;
    }

    const cleaned = trimmed.replace(/\D/g, "");
    if (cleaned !== trimmed) {
      e.preventDefault();
      const input = e.target as HTMLInputElement;

      const start: number = input.selectionStart ?? input.value.length;
      const end: number = input.selectionEnd ?? input.value.length;
      const next = input.value.slice(0, start) + cleaned + input.value.slice(end);

      input.value = next;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      const newCaret = start + cleaned.length;
      input.setSelectionRange(newCaret, newCaret);
    }
  };

  const handleInvoiceChange: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    const f = formikRef.current;
    if (!f) return;

    const val = (e.target.value ?? "").toString().trim();

    if (val === "-" || /^-/.test(val)) {
      alert("Invoice No. no minus ( - ) not allowed .only positive values are allowed.");
      f.setFieldValue("InvoiceNumber", "");
      return;
    }

    const digitsOnly = val.replace(/\D/g, "");
    f.setFieldValue("InvoiceNumber", digitsOnly);
  };

  const ALLOWED_TYPES = [
    "application/pdf",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "image/png",
    "image/jpeg"
  ];
  const onAttachPick = (files: FileList | null) => {
    const f = formikRef.current;
    if (!f || !files) return;

    const tooLarge: string[] = [];
    const invalidType: string[] = [];
    const valid: File[] = [];

    Array.from(files).forEach((file) => {

      // ✅ 1. Size check
      if (file.size > MAX_FILE_SIZE) {
        tooLarge.push(`${file.name} (${(file.size / 1048576).toFixed(2)} MB)`);
        return;
      }

      // ✅ 2. Type check (ADD HERE 👇)
      if (!ALLOWED_TYPES.includes(file.type) && !hasValidExtension(file.name)) {
        invalidType.push(file.name);
        return;
      }

      // ✅ Valid file
      valid.push(file);
    });

    if (tooLarge.length) {
      alert(`File size must be ≤ ${MAX_FILE_SIZE_MB} MB:\n` + tooLarge.join("\n"));
    }

    if (invalidType.length) {
      alert(`Only PDF, Word, Excel, Image files allowed:\n` + invalidType.join("\n"));
    }

    f.setFieldValue("attachments", [...(f.values.attachments || []), ...valid]);
  };
  const onAttachPick2 = (files: FileList | null) => {
    const f = formikRef.current;
    if (!f || !files) return;

    const tooLarge: string[] = [];
    const invalidType: string[] = [];
    const valid: File[] = [];

    Array.from(files).forEach((file) => {

      // ✅ Size check (25 MB)
      if (file.size > MAX_FILE_SIZE) {
        tooLarge.push(`${file.name} (${(file.size / 1048576).toFixed(2)} MB)`);
        return;
      }

      // ✅ Type check
      if (!ALLOWED_TYPES.includes(file.type)) {
        invalidType.push(file.name);
        return;
      }

      valid.push(file);
    });

    if (tooLarge.length) {
      alert(`File size must be ≤ ${MAX_FILE_SIZE_MB} MB:\n` + tooLarge.join("\n"));
    }

    if (invalidType.length) {
      alert(`Only PDF, Word, Excel, Image files allowed:\n` + invalidType.join("\n"));
    }

    f.setFieldValue("attachments", [...(f.values.attachments || []), ...valid]);
  };
  const ALLOWED_EXT = [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".png", ".jpg", ".jpeg"];

  const hasValidExtension = (fileName: string) =>
    ALLOWED_EXT.some(ext => fileName.toLowerCase().endsWith(ext));

  const onAttachPick1 = (files: FileList | null) => {
    const f = formikRef.current; if (!f || !files) return;
    const tooLarge: string[] = []; const valid: File[] = [];
    Array.from(files).forEach((file) => {
      if (file.size > MAX_FILE_SIZE) tooLarge.push(`${file.name} (${(file.size / 1048576).toFixed(2)} MB)`);
      else valid.push(file);
    });
    if (tooLarge.length) alert(`File size must be ≤ ${MAX_FILE_SIZE_MB} MB:\n` + tooLarge.join("\n"));
    f.setFieldValue("attachments", [...(f.values.attachments || []), ...valid]);
  };
  const removeAttachmentAt = (idx: number) => {
    const f = formikRef.current; if (!f) return;
    const next = [...f.values.attachments]; next.splice(idx, 1); f.setFieldValue("attachments", next);
  };

  /** ------------ Validation  ------------ */
  const validate = (v: INonPoFormValues) => {
    if (!isEdit) return {};
    const e: Partial<Record<keyof INonPoFormValues, string>> = {};
    if (!v.VendorCode) e.VendorCode = "Vendor Code is required";
    if (NOE_MASTER && noeOptions.length > 0 && !v.NatureofExpense) e.NatureofExpense = "Nature of Expense is required";
    if (!v.InvoiceNumber?.trim()) e.InvoiceNumber = "Invoice Number is required";
    if (!v.Basicamount?.trim()) e.Basicamount = "Basic Amount is required";
    if (!v.InvoiceDate) e.InvoiceDate = "Invoice Date is required";

    const isMinus = (s?: string) => {
      const t = (s ?? "").trim();
      return t === "-" || /^-/.test(t);
    };
    if (isMinus(v.Basicamount)) e.Basicamount = "Minus not allowed";
    if (isMinus(v.GST)) e.GST = "Minus not allowed";
    if (isMinus(v.OtherCharges)) e.OtherCharges = "Minus not allowed";

    return e;
  };

  /** ------------ Create NEW item ------------ */
  const createNew = async (values: INonPoFormValues) => {
    const list = sp!.web.lists.getByTitle(NONPO_LIST_TITLE);
    const list1 = sp!.web.lists.getByTitle(APPLIATION_LIST_TITLE);

    // === Fetch RM & HOD SharePoint User IDs ===


    let rmUser: any = null;
    let hodUser: any = null;


    const rmEmail = reqDefaults.RMEmail || "";
    const hodEmail = reqDefaults.HODEmail || "";

    try {
      if (rmEmail) rmUser = await sp!.web.siteUsers.getByEmail(rmEmail)();
    } catch (Error) { alert("Error") }

    try {
      if (hodEmail) hodUser = await sp!.web.siteUsers.getByEmail(hodEmail)();
    } catch (Error) { alert("Error") }
    const approvalList = [];

    const getFinancialYear = () => {
      const today = new Date();
      const year = today.getFullYear();
      const month = today.getMonth(); // 0 = Jan

      const startYear = month < 3 ? year - 1 : year;
      const endYear = startYear + 1;

      const shortStart = startYear.toString().slice(-2);
      const shortEnd = endYear.toString().slice(-2);

      return `${shortEnd}`;
    };

    const payload: any = {
      Title: `TMP-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      RequestDate: values.RequestDate ? new Date(values.RequestDate) : null,
      EmployeeCode: EmployeeId || "",
      EmployeeName: EmployeeName || "",
      Email: EmployeeEmail || "",
      ContactNo: EmployeeContactno || "",
      Division: "Railway Business",
      Department: Department || "",
      Location: Locationdata || "",
      RM: EmployeeRM || "",
      HOD: hoddata || "",
      CostCentre: CostCenter || "",
      EmployeeStatus: EmployeeStatus || "",
      VendorCodeId: Number(VendorLookupID || 0),
      VendorName: VendorNameValue,// values.VendorName || "",
      NatureofExpenseId: values.NatureofExpense,
      InvoiceNumber: values.InvoiceNumber || "",
      InvoiceDate: values.InvoiceDate ? new Date(values.InvoiceDate) : null,
      Basicamount: (values.Basicamount || "").toString(),
      GST: (values.GST || "").toString(),
      OtherCharges: (values.OtherCharges || "").toString(),
      Totalamount: (values.Totalamount || "").toString(),
      PaymentRemarks: values.PaymentRemarks || "",
      ApprovalRemarks: values.ApprovalRemarks || "",
      GLCode: values.GLCode || "",
      VoucherNumber: values.VoucherNumber || "",
      PendingAt: "Pending At RM",
      ApproverStatus: "Pending",
      Stage: 1,
      Status: "Pending for Approval",
      CurrentApproverId: RMID.current || null,
      ApprovalMatrix: JSON.stringify(approverJson.current),
      // ApprovalMatrix: JSON.stringify(approvalList),
      // WorkflowHistory: JSON.stringify(workflowHistory),
      WorkflowHistory: JSON.stringify([
        {
          CurrentApprover: props.currentSPContext.pageContext.user.displayName
            || props.currentSPContext.pageContext.user.email
            || '',
          ActionTaken: 'Request submitted',
          Comment: values.PaymentRemarks,
          Date: new Date().toISOString(),
          CurrentStatus: 'Submitted to RM'
        }
      ]),

    };

    const addRes = await list.items.add(payload);
    let itemId = addRes?.data?.Id || addRes?.data?.ID;
    if (!itemId) {
      const found = await list.items.select("Id", "Title").filter(`Title eq '${payload.Title.replace(/'/g, "''")}'`).top(1)();
      if (found?.length) itemId = found[0].Id;
    }
    if (!itemId) throw new Error("Item not created (no Id returned).");

    //  next Title 
    const reqDt = values.RequestDate ? new Date(values.RequestDate) : new Date();
    const fy = fyFromDate(reqDt);
    const prefix = `NON PO/${fy}/`;

    const getLastSeqForFY = async (): Promise<number> => {
      const rows = await list.items
        .select("Id", "Title")
        .filter(`startswith(Title, '${prefix.replace(/'/g, "''")}')`)
        .orderBy("Id", false)
        .top(5)();
      let maxSeq = 0;
      for (const r of rows || []) maxSeq = Math.max(maxSeq, tryParseLastNumber(r.Title || ""));
      return maxSeq;
    };

    const MAX_RETRIES = 6;
    let setOk = false;
    for (let i = 0; i < MAX_RETRIES; i++) {
      const nextSeq = (await getLastSeqForFY()) + 1;
      // const newTitle = buildRequestId(fy, nextSeq);
      //SET TITLE
      const fy1 = getFinancialYear();

      const sp = spfi().using(SPFx(props.currentSPContext));
      const counterItem = await sp.web.lists
        .getByTitle("ApplicationNumber")
        .items
        .select("ID", "IDNo")
        .top(100)();

      console.log(counterItem);
      if (!counterItem || counterItem.length === 0) {
        console.error("Counter row not found in ApplicationNumber list.");
        return;
      }

      const itemId1 = counterItem[0].ID;
      const currentNo = Number(counterItem[0].IDNo) || 0;

      const nextNo = currentNo + 1;
      setNextNo(nextNo);
      setIncrimentalId(itemId1);
      const paddedNumber = nextNo.toString().padStart(5, "0");
      const formattedNumber = `NONPO/${fy1}/${paddedNumber}`;



      try {
        await list.items.getById(itemId).update({ Title: formattedNumber });
        await list1.items.getById(itemId1).update({ IDNo: nextNo });
        setOk = true;
        break;
      } catch (e: any) {
        const msg = (e?.message || "").toLowerCase();
        if (e?.status === 409 || msg.includes("unique") || msg.includes("duplicate") || msg.includes("exists")) {
          await new Promise((r) => setTimeout(r, 120 + Math.random() * 160));
          continue;
        } else {
          throw e;
        }
      }
    }
    if (!setOk) {
      const lastSeq = await getLastSeqForFY();
      await list.items.getById(itemId).update({ Title: buildRequestId(fy, lastSeq + 1) });
    }

    //  Upload attachments 
    if (values.attachments?.length) {
      const att = list.items.getById(itemId).attachmentFiles;
      for (const file of values.attachments) {
        await att.add(file.name, file);
      }
    }

    return itemId;
  };

  /** ------------ Submit (create or update) ------------ */
  const onSubmit = async (values: INonPoFormValues) => {
    if (!sp) { alert("SharePoint context not ready."); return; }
    try {
      setLoading(true);

      if (isNew) {
        const itemId = await createNew(values);
        alert("NON-PO Request created successfully.");

        let reqNoFinal = "NONPO";
        try {
          const meta = await sp.web.lists.getByTitle(NONPO_LIST_TITLE)
            .items.getById(itemId).select("Title")();
          reqNoFinal = meta?.Title || reqNoFinal;
        } catch (Error) { alert("Error") }

        await uploadAttachmentsToLibrary(sp, ATTACH_LIB, reqNoFinal, itemId, values.attachments || [], "Requester");
        navigate("/");
        return;
      }

      // UPDATE existing (Status not changed here)
      if (numericId) {
        const payload: any = {
          VendorCodeId: Number(VendorLookupID || 0),//Number(VendorLookupID),//values.VendorCode,
          VendorName: setVendorNameValue,// values.VendorName || "",
          NatureofExpenseId: values.NatureofExpense,

          InvoiceNumber: values.InvoiceNumber || "",
          InvoiceDate: values.InvoiceDate ? new Date(values.InvoiceDate) : null,

          Basicamount: (values.Basicamount || "").toString(),
          GST: (values.GST || "").toString(),
          OtherCharges: (values.OtherCharges || "").toString(),
          Totalamount: (values.Totalamount || "").toString(),

          PaymentRemarks: values.PaymentRemarks || "",
          ApprovalRemarks: values.ApprovalRemarks || "",
          GLCode: values.GLCode || "",


          VoucherNumber: values.VoucherNumber || "",
          //Amount: (values.Amount || "").toString(),
        };

        await sp.web.lists.getByTitle(NONPO_LIST_TITLE).items.getById(numericId).update(payload);

        if (values.attachments?.length) {
          let reqNo = "NONPO";
          try {
            const meta = await sp.web.lists.getByTitle(NONPO_LIST_TITLE)
              .items.getById(numericId)
              .select("Title")();
            reqNo = meta?.Title || reqNo;
          } catch (Error) { alert("Error") }

          await uploadAttachmentsToLibrary(sp, ATTACH_LIB, reqNo, numericId, values.attachments);
        }

        alert("NON-PO Request updated successfully.");
        return;
      }
    } catch (err: any) {
      console.error("Submit error:", err);
      let msg = err?.message || "Failed to submit";
      try {
        const body = err?.data?.responseBody || err?.data?.error?.message || err?.responseBody;
        if (body) {
          const parsed = typeof body === "string" ? JSON.parse(body) : body;
          msg = parsed?.["odata.error"]?.message?.value || parsed?.error?.message?.value || msg;
        }
      } catch (Error) { alert(msg) }
      alert(msg);
    } finally {
      setLoading(false);
    }
  };

  /** ------------ UI ------------ */
  if (!profileLoaded || !sp) {
    return (
      <div className="p-3">
        <div className="alert alert-secondary mb-0">Loading…</div>
      </div>
    );
  }


  return (
    <Formik
      initialValues={initialValues}
      innerRef={formikRef}
      enableReinitialize={true}
      validate={validate}
      onSubmit={onSubmit}
    >
      {({ values, errors, touched, validateForm, submitForm, setTouched }) => (
        <Form>


          <div className='MainUplodForm' style={{ margin: "0px" }}>
            <div className='row'>
              <div className='col-md-12'>
                <div className='Main-Boxpoup'>

                  {/* 🔹 Header */}
                  <div className="bordered">
                    <img src={logo} />
                    <h1>NON PO Request </h1>
                  </div>
                  <div className='displayWF'>{WorkflowJSX}</div>
                  <div className='borderedbox'>
                    {/* 🔹 Section Title */}
                    <div className="heading1">
                      <label>Requester Information</label>
                    </div>
                    <div className='main-formcontainer'>
                      <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Request Date </label> &nbsp; : &nbsp;
                          <Field type="date" readOnly name="RequestDate" className="fonttext" />
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Employee Name </label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {EmployeeName}</label>
                          {/* <Field readOnly name="EmployeeName" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Employee Email </label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {EmployeeEmail}</label>
                          {/* <Field readOnly name="Email" className="fonttext" /> */}
                        </div>
                      </div>
                      <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Contact No</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {EmployeeContactno}</label>
                          {/* <Field readOnly name="ContactNo" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Employee Status</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {EmployeeStatus}</label>
                          {/* <Field readOnly name="EmployeeStatus" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Division</label> &nbsp; : &nbsp;
                          <label className='fonttext'>Railway Business </label>
                          {/* <Field readOnly name="Division"  {EmployeeDivision} className="fonttext" /> */}
                        </div>
                      </div>
                      <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Location</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {Locationdata}</label>
                          {/* <Field readOnly name="Location" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>RM</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {EmployeeRM}</label>
                          {/* <Field readOnly name="RM" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>HOD</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {hoddata}</label>
                          {/* <Field readOnly name="HOD" className="fonttext" /> */}
                        </div>
                      </div>
                      <div className='row mb-20'>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Department</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {Department}</label>
                          {/* <Field readOnly name="Department" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Employee Code</label> &nbsp; : &nbsp;
                          <label className='fonttext'>  {EmployeeId}</label>
                          {/* <Field readOnly name="EmployeeCode" className="fonttext" /> */}
                        </div>
                        <div className='col-md-4 alignbox'>
                          <label className='font fontpadd'>Status</label> &nbsp; : &nbsp;
                          {/* <Field readOnly name="Status" className="fonttext" /> */}
                        </div>
                      </div>
                    </div>
                    <div className="heading1">
                      <label>Invoice Details</label>
                    </div>
                    <div className='main-formcontainer'>
                      <div className='row mb-20'>
                        <div className='col-md-4'>
                          <label className='font'>Vendor Name <span className='Mantorystar'>*</span></label>
                          <select
                            className={`form-select ${touched.VendorName && errors.VendorName ? "is-invalid" : ""}`}
                            value={values.VendorName || ""}
                            onChange={onVendorChange}
                            disabled={!isEdit}
                          >
                            <option value="">Select vendor</option>
                            {vendors.map((v) => (
                              <option key={v.id} value={v.id}>{v.display}</option>
                            ))}
                          </select>
                          {touched.VendorName && errors.VendorName && <div className="invalid-feedback">{`${errors.VendorCode}`}</div>}
                        </div>
                        <div className='col-md-4'>
                          <label className='font'>Vendor Code <span className='Mantorystar'>*</span></label>
                          <Field
                            name="VendorCode"
                            className="form-control readonly"
                            placeholder="Auto from Vendor"
                            readOnly
                          />
                        </div>
                        <div className='col-md-4'>
                          <label className='font'>Nature of Expenses <span className='Mantorystar'>*</span></label>
                          <select
                            className={`form-select ${touched.NatureofExpense && errors.NatureofExpense ? "is-invalid" : ""}`}
                            value={values.NatureofExpense || ""}
                            onChange={onNoeChange}
                            disabled={!isEdit || !NOE_MASTER || noeOptions.length === 0}
                          >
                            <option value="">{NOE_MASTER ? "Select nature" : "Master not configured"}</option>
                            {noeOptions.map((n) => (
                              <option key={n.id} value={n.id}>{n.title}</option>
                            ))}
                          </select>
                          {touched.NatureofExpense && errors.NatureofExpense && <div className="invalid-feedback">{`${errors.NatureofExpense}`}</div>}
                        </div>
                      </div>
                      <div className='row mb-20'>
                        <div className='col-md-4'>
                          <label className='font'>Invoice No. <span className='Mantorystar'>*</span></label>
                          <Field
                            name="InvoiceNumber"
                            className={`form-control ${touched.InvoiceNumber && errors.InvoiceNumber ? "is-invalid" : ""}`}
                            placeholder="Enter invoice number"
                            disabled={!isEdit}
                          />
                          {touched.InvoiceNumber && errors.InvoiceNumber && <div className="invalid-feedback">{`${errors.InvoiceNumber}`}</div>}
                        </div>
                        <div className='col-md-4'>
                          <label className='font'>Invoice Date <span className='Mantorystar'>*</span></label>
                          <Field name="InvoiceDate" type="date" className={`form-control ${touched.InvoiceDate && errors.InvoiceDate ? "is-invalid" : ""}`} disabled={!isEdit} />
                          {touched.InvoiceDate && errors.InvoiceDate && <div className="invalid-feedback">{`${errors.InvoiceDate}`}
                          </div>}
                        </div>
                        <div className='col-md-4'>
                          <label className='font'>Basic Amount <span className='Mantorystar'>*</span></label>
                          <Field name="Basicamount">
                            {() => (
                              <input
                                type="text"
                                className={`form-control ${touched.Basicamount && errors.Basicamount ? "is-invalid" : ""}`}
                                value={values.Basicamount}
                                onChange={handleAmountChange("Basicamount")}
                                onKeyDown={blockMinusAndAlpha}
                                onPaste={onAmountPaste}
                                placeholder="0"
                                disabled={!isEdit}
                                inputMode="decimal"
                              />

                            )}
                          </Field>
                          {touched.Basicamount && errors.Basicamount && (
                            <div className="invalid-feedback">{String(errors.Basicamount)}</div>
                          )}
                        </div>


                      </div>
                      <div className='row mb-20'>
                        <div className='col-md-4'>
                          <label className='font'>GST Amount </label>
                          <Field name="GST">
                            {() => (
                              <input
                                type="text"
                                className={`form-control ${touched.GST && errors.GST ? "is-invalid" : ""}`}
                                value={values.GST}
                                onChange={handleAmountChange("GST")}
                                onKeyDown={blockMinusAndAlpha}
                                onPaste={onAmountPaste}
                                placeholder="0"
                                disabled={!isEdit}
                                inputMode="decimal"
                              />
                            )}
                          </Field>
                          {touched.GST && errors.GST && (
                            <div className="invalid-feedback">{String(errors.GST)}</div>
                          )}
                        </div>
                        <div className='col-md-4'>
                          <label className='font'>Other Charges (if any) </label>
                          <Field name="OtherCharges">
                            {() => (
                              <input
                                type="text"
                                className={`form-control ${touched.OtherCharges && errors.OtherCharges ? "is-invalid" : ""}`}
                                value={values.OtherCharges}
                                onChange={handleAmountChange("OtherCharges")}
                                onKeyDown={blockMinusAndAlpha}
                                onPaste={onAmountPaste}
                                placeholder="0"
                                disabled={!isEdit}
                                inputMode="decimal"
                              />
                            )}
                          </Field>
                          {touched.OtherCharges && errors.OtherCharges && (
                            <div className="invalid-feedback">{String(errors.OtherCharges)}</div>
                          )}
                        </div>
                        <div className='col-md-4'>
                          <label className='font'>Total Amount : Basic + GST + Other Charges </label>
                          <Field name="Totalamount" className="form-control readonly" />
                        </div>
                      </div>
                      <div className='row mb-20'>
                        <div className='col-md-12'>
                          <label className='font'>Remarks for NON-PO Expense</label>
                          <Field name="PaymentRemarks" className="form-control" type="text" disabled={!isEdit} />
                        </div>
                      </div>
                    </div>
                    <div className="heading1">
                      <label>Upload Documents</label>
                    </div>
                    <div className='main-formcontainer'>
                      <div className="row mb-20">
                        <div className="col-md-12">
                          <label className="font">Attach</label>
                          <div className="d-flex align-items-center gap-2">
                            <label className={`btn btn-${isEdit ? "outline-dark" : "secondary"} btn-sm mb-0 ${!isEdit && "disabled"}`}>
                              <i className="fas fa-paperclip me-2" />
                              Select files
                              <input type="file" multiple disabled={!isEdit} accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg" style={{ display: "none" }} onChange={(e) => onAttachPick((e.target as HTMLInputElement).files)} />
                            </label>
                            <small className={`text-${values.attachments.length ? "muted" : "secondary"}`}>
                              {values.attachments.length ? `${values.attachments.length} file(s) selected` : "Optional"}
                            </small>
                          </div>
                          {values.attachments.length > 0 && (
                            <div className="mt-2">
                              <div style={{ overflowX: "auto" }}>
                                <table className="custom-table table table-sm table-bordered mb-0">
                                  <thead className="table-light">
                                    <tr>
                                      <th className="px-4 py-2">File</th>
                                      <th className="px-4 py-2">Size</th>
                                      <th className="px-4 py-2">Action</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {values.attachments.map((f, idx) => (
                                      <tr key={`${f.name}-${idx}`}>
                                        <td className="px-4 py-2">{f.name}</td>
                                        <td className="px-4 py-2">{(f.size / (1024 * 1024)).toFixed(2)} MB</td>
                                        <td className="px-4 py-2">
                                          <button type="button" className="btn btn-sm btn-outline-danger" onClick={() => removeAttachmentAt(idx)} disabled={!isEdit}>Remove</button>
                                        </td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className='row my-3'>
                      <div className='col-md-12'>
                        <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                          {isEdit && (
                            // <a onClick={async () => {
                            //   const f = formikRef.current; if (!f) return;
                            //   await Promise.all([
                            //     f.setFieldTouched("VendorCode", true, true),
                            //     f.setFieldTouched("NatureofExpense", true, true),
                            //     f.setFieldTouched("InvoiceNumber", true, true),
                            //     f.setFieldTouched("InvoiceDate", true, true),
                            //   ]);
                            //   const errs = await f.validateForm();
                            //   if (Object.keys(errs || {}).length > 0) return;
                            //   f.submitForm();
                            // }} className="Submit-btn">
                            //   {isNew ? "Submit" : "Save"}
                            // </a>
                            <a
                              onClick={async () => {
                                const f = formikRef.current;
                                if (!f || f.isSubmitting) return; // 🚫 रोक दो multiple clicks

                                await Promise.all([
                                  f.setFieldTouched("VendorCode", true, true),
                                  f.setFieldTouched("NatureofExpense", true, true),
                                  f.setFieldTouched("InvoiceNumber", true, true),
                                  f.setFieldTouched("InvoiceDate", true, true),
                                  f.setFieldTouched("Basicamount", true, true),
                                ]);

                                const errs = await f.validateForm();
                                if (Object.keys(errs || {}).length > 0) return;

                                await f.submitForm();
                              }}
                              className={`Submit-btn ${formikRef.current?.isSubmitting ? "disabled" : ""}`}
                              style={{ pointerEvents: formikRef.current?.isSubmitting ? "none" : "auto", opacity: formikRef.current?.isSubmitting ? 0.6 : 1 }}
                            >
                              {formikRef.current?.isSubmitting ? "Submitting..." : (isNew ? "Submit" : "Save")}
                            </a>
                          )}

                          {/* <Link to="#" onClick={() => { history.goBack() }} className="reset-btn">
                            Exit
                          </Link> */}
                          <a onClick={() => { navigate(-1); }} className="Exit-btn">
                            Exit
                          </a>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

        </Form>
      )}
    </Formik>
  );
};

export default NonPoRequest;