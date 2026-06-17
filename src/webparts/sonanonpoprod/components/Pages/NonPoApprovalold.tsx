import React, { useEffect, useMemo, useState } from "react";
import { useNavigate, useSearchParams } from "react-router-dom";
import useFullscreenForm from "../../hook/useFullscreenForm";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import SPCRUDOPS from "../../service/DAL/spcrudops";

const ROUTE_APPROVED = "/MyApprovalApproved";   
const ROUTE_REJECTED = "/MyApprovalRejected";  
const ROUTE_APTEAM  = "/APTeamDashboard";       

const HISTORY_LIST = "NonPOApprovalHistory";
const HISTORY_PARENT_FIELD = "ParentId";

/** ======  helpers ====== */
const fmtDate = (d?: string | Date | null): string => {
  if (!d) return "—";
  const dt = new Date(d);
  if (Number.isNaN(dt.getTime())) return "—";
  const dd = dt.getDate().toString().padStart(2, "0");
  const mm = (dt.getMonth() + 1).toString().padStart(2, "0");
  const yyyy = dt.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
};
const toText = (v: any) => (v === null || v === undefined || v === "" ? "—" : String(v));
const ci = (s?: string) => (s ?? "").trim().toLowerCase();

interface HistoryRow {
  id: number;
  approver: string;
  actionTaken: string;
  remarks: string;
  actionDate: string;
}

type Props = ISonanonpoprodProps & {
  title?: string;
  listTitle?: string; 
};


//export const NonPoApproval = (props: ISonanonpoprodProps) => {
const NonPoApproval: React.FC<Props> = ({
  currentSPContext,
  title = "NON PO Approval",
  listTitle = "NonPO",
}) => {
  useFullscreenForm();
  const navigate = useNavigate();
  const [params] = useSearchParams();

  const idParam = params.get("id");
  const modeParam = (params.get("mode") || "view").toLowerCase();
  const isEdit = modeParam === "edit";

  const numericId = useMemo(() => {
    const n = Number(idParam);
    return Number.isFinite(n) ? n : undefined;
  }, [idParam]);

  // UI state
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string>("");

  // Requester 
  const [requestNo, setRequestNo] = useState<string | undefined>(undefined);
  const [requestDate, setRequestDate] = useState("—");
  const [employeeName, setEmployeeName] = useState("—");
  const [division, setDivision] = useState("—");
  const [email, setEmail] = useState("—");
  const [hod, setHod] = useState("—");
  const [department, setDepartment] = useState("—");
  const [employeeStatus, setEmployeeStatus] = useState("—");
  const [contactNo, setContactNo] = useState("—");
  const [locationVal, setLocationVal] = useState("—");
  const [rm, setRm] = useState("—");

  // Invoice (editable in edit mode)
  const [vendorCode, setVendorCode] = useState("—"); 
  const [vendorName, setVendorName] = useState("");
  const [natureOfExpenses, setNatureOfExpenses] = useState("—");
  const [invoiceNo, setInvoiceNo] = useState("");
  const [invoiceDate, setInvoiceDate] = useState("—");
  const [basicAmount, setBasicAmount] = useState("");
  const [gst, setGst] = useState("");
  const [otherCharges, setOtherCharges] = useState("");
  const [totalAmount, setTotalAmount] = useState("");
  const [ApprovalMatrixdata, setApprovalMatrix] = React.useState<any[]>([]);
  const [WorkflowHistorydata, setWorkflowHistory] = React.useState<any[]>([]);

  // Remarks
  const [paymentRemarks, setPaymentRemarks] = useState("");
  const [approvalRemarks, setApprovalRemarks] = useState("");
  const [HODApprover, setHODApprover] = React.useState(false);
  const [CFOApprover, setCFOApprover] = React.useState(false);
  const [PerformerApprover, setPerformerApprover] = React.useState(false);

  // History
  const [history, setHistory] = useState<HistoryRow[]>([]);

  /** =====history ===== */
  const loadItem = async () => {
    const spCrudOps = await SPCRUDOPS(currentSPContext);

    const select =
      "Id,Title,RequestDate,EmployeeName,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus," +
      "ApprovalMatrix,VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,WorkflowHistory,Status,PaymentRemarks,ApprovalRemarks," +
      "VendorCode/Title,NatureofExpense/Title,CurrentApprover/ID,CurrentApprover/EMail,CurrentApprover/Title";
    const expand = "VendorCode,NatureofExpense,CurrentApprover";
    const orderBy = { column: "Id", isAscending: false };

    const rows =
      (await spCrudOps.getTopData(
        listTitle, select, expand, `Id eq ${numericId}`, orderBy, 1,
        { currentSPContext } as ISonanonpoprodProps
      )) ?? [];

    if (!rows.length) throw new Error(`Item not found for ID: ${numericId}`);
    const it = rows[0];

    setRequestNo(it?.Title || undefined);

    // Requester (read-only)
    setRequestDate(fmtDate(it?.RequestDate));
    setEmployeeName(toText(it?.EmployeeName));
    setDivision(toText(it?.Division));
    setEmail(toText(it?.Email));
    setHod(toText(it?.HOD));
    setDepartment(toText(it?.Department));
    setEmployeeStatus(toText(it?.EmployeeStatus));
    setContactNo(toText(it?.ContactNo));
    setLocationVal(toText(it?.Location));
    setRm(toText(it?.RM));

    // Invoice
    setVendorCode(toText(it?.VendorCode?.Title));
    setVendorName(toText(it?.VendorName) === "—" ? "" : String(it?.VendorName || ""));
    setNatureOfExpenses(toText(it?.NatureofExpense?.Title));
    setInvoiceNo(toText(it?.InvoiceNumber) === "—" ? "" : String(it?.InvoiceNumber || ""));
    setInvoiceDate(fmtDate(it?.InvoiceDate));
    setBasicAmount(toText(it?.Basicamount) === "—" ? "" : String(it?.Basicamount || ""));
    setGst(toText(it?.GST) === "—" ? "" : String(it?.GST || ""));
    setOtherCharges(toText(it?.OtherCharges) === "—" ? "" : String(it?.OtherCharges || ""));
    setTotalAmount(toText(it?.Totalamount) === "—" ? "" : String(it?.Totalamount || ""));
    setPaymentRemarks(toText(it?.PaymentRemarks) === "—" ? "" : String(it?.PaymentRemarks || ""));
    setApprovalRemarks(toText(it?.ApprovalRemarks) === "—" ? "" : String(it?.ApprovalRemarks || ""));

      let parsedApprovalMatrix = [];
                if (it.ApprovalMatrix) {
                    try {
                        parsedApprovalMatrix = JSON.parse(it.ApprovalMatrix);
                    } catch (parseError) {
                        console.error("Error parsing ApprovalMatrix JSON:", parseError);
                    }
                }

                setApprovalMatrix(parsedApprovalMatrix);

                   

      let parsedWorkflowHistory = [];
                if (it.WorkflowHistory) {
                    try {
                        parsedWorkflowHistory = JSON.parse(it.WorkflowHistory);
                    } catch (parseError) {
                        console.error("Error parsing WorkflowHistory JSON:", parseError);
                    }
                }

                setWorkflowHistory(parsedWorkflowHistory);          
  };

  const loadHistory = async () => {
    try {
      const spCrudOps = await SPCRUDOPS(currentSPContext);

      const selectH = "Id,Title,ActionTaken,Remarks,ActionDate,Approver/Title";
      const expandH = "Approver";
      const orderH = { column: "Id", isAscending: false };

      let rowsH =
        (await spCrudOps.getRootData(
          HISTORY_LIST, selectH, expandH, `${HISTORY_PARENT_FIELD} eq ${numericId}`, orderH,
          { currentSPContext } as ISonanonpoprodProps
        )) ?? [];

      if (!rowsH.length && requestNo) {
        rowsH =
          (await spCrudOps.getRootData(
            HISTORY_LIST, selectH, expandH, `Title eq '${String(requestNo).replace(/'/g, "''")}'`, orderH,
            { currentSPContext } as ISonanonpoprodProps
          )) ?? [];
      }

      const mapped: HistoryRow[] = rowsH.map((h: any) => ({
        id: h?.Id,
        approver: h?.Approver?.Title || "—",
        actionTaken: toText(h?.ActionTaken),
        remarks: toText(h?.Remarks),
        actionDate: fmtDate(h?.ActionDate),
      }));
      setHistory(mapped);
    } catch (e) {
      console.warn("[NonPoApproval] history fetch warning:", e);
      setHistory([]);
    }
  };

  useEffect(() => {
    let abort = false;

    (async () => {
      try {
        setLoading(true); setErr("");
        if (!numericId) {
          setErr("Invalid or missing item ID in the URL.");
          return;
        }
        await loadItem();
        await loadHistory();
      } catch (e: any) {
        console.error("[NonPoApproval] fetch error:", e);
        setErr(e?.message || "Failed to load item");
      } finally {
        if (!abort) setLoading(false);
      }
    })();

    return () => { abort = true; };
  }, [currentSPContext, listTitle, numericId]);

  const handleSave = async () => {
    try {
      if (!isEdit || !numericId) return;
      const spCrudOps = await SPCRUDOPS(currentSPContext);
      const data: any = { ApprovalRemarks: approvalRemarks ?? "" };
      await spCrudOps.updateData(listTitle, numericId, data, { currentSPContext } as ISonanonpoprodProps);
      alert("Approval Remarks saved.");
    } catch (e: any) {
      alert("Save failed: " + (e?.message || e));
    }
  };


  const updateStatus = async (newStatus: "Approved" | "Rejected") => {
    if (!numericId) return;
    const spCrudOps = await SPCRUDOPS(currentSPContext);

    const meName: string =
      (currentSPContext?.pageContext?.user?.displayName as string) ||
      (currentSPContext?.pageContext?.legacyPageContext?.userDisplayName as string) ||
      "";
    const isRM  = ci(meName) === ci(rm);
    const isHOD = ci(meName) === ci(hod);

    let payload: any = { ApprovalRemarks: approvalRemarks ?? "" };

    if (newStatus === "Rejected") {
      payload = { ...payload, Status: "Rejected" };
    } else {
      if (isHOD) {
        payload = { ...payload, Status: "Pending for Acceptance" };
      } else if (isRM) {
        payload = { ...payload, Status: "Pending for Approval" }; // stays pending for HOD
      } else {
        payload = { ...payload, Status: "Pending for Approval" };
      }
    }

    await spCrudOps.updateData(
      listTitle,
      numericId,
      payload,
      { currentSPContext } as ISonanonpoprodProps
    );

    try {
      await spCrudOps.insertData(HISTORY_LIST, {
        Title: requestNo || "",
        [HISTORY_PARENT_FIELD]: numericId,
        ActionTaken: newStatus === "Approved" ? (isHOD ? "Approved (HOD)" : "Approved (RM)") : "Rejected",
        Remarks: approvalRemarks ?? "",
        ActionDate: new Date().toISOString(),
        ApproverName: meName, 
      }, { currentSPContext } as ISonanonpoprodProps);
    } catch (e) {
      console.warn("[NonPoApproval] history insert skipped/failed]:", e);
    }

    // 3) Navigate
    if (newStatus === "Rejected") {
      navigate(`${ROUTE_REJECTED}?from=approval&status=Rejected`);
    } else {
      if (isHOD) {
        navigate(`${ROUTE_APTEAM}?from=approval&status=Pending%20for%20Acceptance`);
      } else {
        navigate(`${ROUTE_APPROVED}?from=approval&status=Approved`);
      }
    }
  };

  /** ===== UI ===== */
  return (
    <div className="approval-root" style={{ padding: 20, background: "#f5f5f5" }}>
      <h2 className="page-title" style={{ margin: 0, marginBottom: 10 }}>
        {title} {requestNo ? `— ${requestNo}` : ""}
      </h2>

      {err && <div style={{ color: "#b91c1c", marginBottom: 8 }}>{err}</div>}
      {loading && <div style={{ marginBottom: 8 }}>Loading…</div>}

      {/* ===== Requester Information  ===== */}
      <section className="section-box" style={sectionStyle}>
        <h3 className="section-title" style={sectionHead}>Requester Information</h3>
        <div className="grid-2" style={grid2}>
          <p><b>Request Date:</b> {requestDate}</p>
          <p><b>Employee Name:</b> {employeeName}</p>

          <p><b>Division:</b> {division}</p>
          <p><b>Email:</b> {email}</p>

          <p><b>HOD:</b> {hod}</p>
          <p><b>Department:</b> {department}</p>

          <p><b>Employee Status:</b> {employeeStatus}</p>
          <p><b>Contact No:</b> {contactNo}</p>

          <p><b>Location:</b> {locationVal}</p>
          <p><b>RM:</b> {rm}</p>
        </div>
      </section>

      {/* ===== Invoice Details ===== */}
      <section className="section-box" style={sectionStyle}>
        <h3 className="section-title" style={sectionHead}>Invoice Details</h3>
        <div className="grid-3" style={grid3}>
          <div><b>Vendor Code*:</b> <br />
            <input value={vendorCode} onChange={e => setVendorCode(e.target.value)} disabled style={input} />
          </div>
          <div>
            <b>Vendor Name:</b><br />
            <input value={vendorName} onChange={e => setVendorName(e.target.value)} disabled style={input} />
          </div>
          <div><b>Nature of Expenses:</b> <br />
            <input value={natureOfExpenses} onChange={e => setNatureOfExpenses(e.target.value)} disabled style={input} />
          </div>

          <div>
            <b>Invoice No.*:</b><br />
            <input value={invoiceNo} onChange={e => setInvoiceNo(e.target.value)} disabled style={input} />
          </div>
          <div><b>Invoice Date:</b> <br />
            <input value={invoiceDate} onChange={e => setInvoiceDate(e.target.value)} disabled style={input} />
          </div>
          <div>
            <b>Basic Amount:</b><br />
            <input value={basicAmount} onChange={e => setBasicAmount(e.target.value)} disabled style={input} />
          </div>

          <div>
            <b>GST Amount:</b><br />
            <input value={gst} onChange={e => setGst(e.target.value)} disabled style={input} />
          </div>
          <div>
            <b>Other Charges (If any):</b><br />
            <input value={otherCharges} onChange={e => setOtherCharges(e.target.value)} disabled style={input} />
          </div>
          <div>
            <b>Total Amount:</b><br />
            <input value={totalAmount} onChange={e => setTotalAmount(e.target.value)} disabled  style={input} />
          </div>
        </div>
      </section>

      {/* ===== Remarks ===== */}
      <section className="section-box" style={sectionStyle}>
        <div className="grid-2" style={grid2}>
          <div>
            <p><b>Payment Remarks</b></p>
            <input value={paymentRemarks} onChange={e => setPaymentRemarks(e.target.value)} disabled  style={input} />
          </div>
          <div>
            <p><b>Approval Remarks</b></p>
            <input value={approvalRemarks} onChange={e => setApprovalRemarks(e.target.value)} disabled={!isEdit} style={input} />
          </div>
        </div>
      </section>

      {/* ===== Approval History ===== */}
      <section className="history-box" style={historyBox}>
        <h3 className="history-title" style={{ marginTop: 0, marginBottom: 10, fontWeight: "bold" }}>
          Approval History
        </h3>

        {history.length === 0 ? (
          <div style={{ color: "#6b7280" }}>No history available</div>
        ) : (
          history.map(h => (
            <div key={h.id} className="history-entry" style={historyEntry}>
              <p><b>Approval By:</b> {h.approver}</p>
              <p><b>Action Taken:</b> {h.actionTaken}</p>
              <p><b>Remarks:</b> {h.remarks}</p>
              <p><b>Action Date:</b> {h.actionDate}</p>
              <hr />
            </div>
          ))
        )}
      </section>

      {/* ===== Action Buttons ===== */}
      <div className="button-row" style={{ display: "flex", gap: 12, marginTop: 10 }}>
        {isEdit && (
          <>
            <button className="btn" style={btnApprove} onClick={() => updateStatus("Approved")}>Approve</button>
            <button className="btn" style={btnReject}  onClick={() => updateStatus("Rejected")}>Reject</button>
            <button className="btn" style={btnExit}    onClick={() => navigate("/MyApproval")}>Exit</button>
          </>
        )}

        {!isEdit && (
          <>
            {/* <button className="btn" style={btnPrimary} onClick={handleSave}>💾 Save Changes</button> */}
            <button className="btn" style={btnExit} onClick={() => navigate("/MyApproval")}>Exit</button>
          </>
        )}
      </div>
    </div>
  );
};

export default NonPoApproval;

/** Tiny styles */
const sectionStyle: React.CSSProperties = { background: "#fff", padding: "18px 20px", borderRadius: 6, marginBottom: 18, border: "1px solid #ddd" };
const sectionHead: React.CSSProperties = { margin: 0, marginBottom: 12, fontSize: 14, fontWeight: 700 };
const grid2: React.CSSProperties = { display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: "8px 20px" };
const grid3: React.CSSProperties = { display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "8px 20px" };
const input: React.CSSProperties = { width: "100%", padding: 8, border: "1px solid #ddd", borderRadius: 4 };

const historyBox: React.CSSProperties = { border: "2px solid #cc0000", borderRadius: 6, padding: 16, background: "#fff", marginBottom: 20 };
const historyEntry: React.CSSProperties = { fontSize: 13, lineHeight: 1.4, marginBottom: 8 };

const btnBase = { padding: "8px 20px", borderRadius: 6, border: "none", color: "#fff", fontWeight: 700, cursor: "pointer" } as React.CSSProperties;
const btnApprove = { ...btnBase, background: "#4CAF50" };
const btnReject  = { ...btnBase, background: "#c62828" };
const btnExit    = { ...btnBase, background: "#333" };
const btnPrimary = { ...btnBase, background: "#2563eb" };