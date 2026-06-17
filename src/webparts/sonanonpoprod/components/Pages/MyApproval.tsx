import React, { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";

import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import SPCRUDOPS from "../../service/DAL/spcrudops";
import Edit from "../../components/Pages/Image/Pencil.png"
import View from "../../components/Pages/Image/Eye.png";

import logo from "../../assets/sona-comstarlogo.png";

import { Link } from "react-router-dom";

export type NonPoStatus =
  | "Pending for Approval"
  | "Sent back"
  | "Approved"
  | "Rejected"
  | "Draft"
  | "Pending for Acceptance"
  | "Pending for Vouching"
  | "Closed"
  | string;

export interface ApprovalRow {
  requestNo: string;
  requestDate: string;
  buyerName: string;
  vendorCode: string;
  vendorName: string;
  invoiceNo: string;
  amount: number;
  status: NonPoStatus;
  _spId?: number;
  CurrentApproverEmail: string;
}

interface MyApprovalOwnProps {
  listTitle?: string;
  title?: string;
  onView?: (row: ApprovalRow) => void;
}

type MyApprovalProps = ISonanonpoprodProps & MyApprovalOwnProps;

/** ------ Helpers ------ */
const fmtDate = (d?: string | Date | null): string => {
  if (!d) return "";
  const dt = new Date(d);
  if (Number.isNaN(dt.getTime())) return "";
  const dd = dt.getDate().toString().padStart(2, "0");
  const mm = (dt.getMonth() + 1).toString().padStart(2, "0");
  const yyyy = dt.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
};
const safeNum = (n: any): number => {
  const v = Number((n ?? "").toString().replace(/,/g, "").trim());
  return Number.isFinite(v) ? v : 0;
};
const buildSelectExpand = () => {
  const select =
    "Id,Title,RequestDate,EmployeeName,VendorName,InvoiceNumber,Totalamount,Amount,Status,VendorCode/VendorCode,VendorCode/Id,VendorCode/Title,CurrentApprover/Title,CurrentApprover/EMail";
  const expand = "VendorCode,CurrentApprover";
  return { select, expand };
};

/** ------ Component ------ */
const MyApproval: React.FC<MyApprovalProps> = ({
  listTitle = "NonPO",
  title = "NON PO Approval Dashboard",
  onView,
  ...props
}) => {
  const navigate = useNavigate();

  const [query, setQuery] = useState("");
  const [rows, setRows] = useState<ApprovalRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  // Pagination
  const PAGE_SIZE = 10;
  const [currentPage, setCurrentPage] = useState(1);

  /** ---------- Fetch approvals ---------- **/
  useEffect(() => {
    let abort = false;
    (async () => {
      try {
        setLoading(true);
        setError("");

        const spCrudOps = await SPCRUDOPS(props.currentSPContext);
        const { select, expand } = buildSelectExpand();

        const filter = `Status eq 'Pending for Approval'`;
        const orderBy = { column: "Id", isAscending: false };

        const items =
          (await spCrudOps.getRootData(listTitle, select, expand, filter, orderBy, props)) ?? [];

        if (abort) return;

        const mapped: ApprovalRow[] = items.map((it: any) => ({
          _spId: it?.Id,
          requestNo: it?.Title || "",
          requestDate: fmtDate(it?.RequestDate),
          buyerName: it?.EmployeeName || "",
          vendorCode: it?.VendorCode?.VendorCode || "",
          vendorName: it?.VendorName || "",
          invoiceNo: it?.InvoiceNumber || "",
          amount: it?.Totalamount ? safeNum(it?.Totalamount) : safeNum(it?.Amount),
          CurrentApproverEmail:it?.CurrentApprover?.EMail,
          status: (it?.Status as NonPoStatus) || "Pending for Approval",
        }));
        let filteredmapped = mapped.filter((m) => (m.CurrentApproverEmail === props.userEmail))
        //setRows(mapped);
        setRows(filteredmapped)
      } catch (e: any) {
        console.error("[MyApproval] fetch error:", e);
        setError(e?.message || "Failed to load data");
      } finally {
        if (!abort) setLoading(false);
      }
    })();
    return () => {
      abort = true;
    };
  }, [listTitle, props.currentSPContext]);

  /** ---------- Search ---------- **/
  const filtered = useMemo(() => {
    if (!query.trim()) return rows;
    const q = query.toLowerCase();
    return rows.filter((r) =>
      [
        r.requestNo,
        r.requestDate,
        r.buyerName,
        r.vendorCode,
        r.vendorName,
        r.invoiceNo,
        String(r.amount),
        r.status,
      ]
        .join(" ")
        .toLowerCase()
        .includes(q)
    );
  }, [rows, query]);

  /** ---------- Reset page when filter changes ---------- **/
  useEffect(() => setCurrentPage(1), [query, rows]);

  // Pagination values
  const totalItems = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  const page = Math.min(currentPage, totalPages);
  const startIdx = (page - 1) * PAGE_SIZE;
  const endIdx = Math.min(startIdx + PAGE_SIZE, totalItems);
  const paged = filtered.slice(startIdx, endIdx);

  /** ---------- Handlers ---------- **/
  const handleView = (row: ApprovalRow) => {
    if (onView) return onView(row);
    if (!row?._spId) return alert("Item ID missing for this row.");
    navigate(`/NonPoApproval?id=${row._spId}&mode=view`, {
      state: { from: "/MyApproval", fullscreen: true },
    });
  };
  const handleEdit = (row: ApprovalRow) => {
    if (!row?._spId) return alert("Item ID missing for this row.");
    navigate(`/NonPoApproval?id=${row._spId}&mode=edit`, {
      state: { from: "/MyApproval", fullscreen: true },
    });
  };

  /** ---------- Render ---------- **/
  return (
    <div>
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2> {title} </h2>
          </div>
        </div>
      </div>
      <div className='col-md-12 px-2 d-flex justify-content-between align-items-center flex-wrap' style={{margin: "5px"}}>
        <div>
          <input
            type="text" placeholder="Search"
            className="form-control" style={{ width: "250px" }}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
      </div>
      {error && <div className="alert">{error}</div>}
      {loading && <div className="alert">Loading...</div>}
      <main className="Main-Dash mx-2">
        <div className="overflow-x-auto">
          <div className="table-vert-scroll">
            <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
              <thead style={{ backgroundColor: "#3c3e45" }} className="text-white">
                <tr>
                  <th className="px-4 py-2">View</th>
                  <th className="px-4 py-2">Request No.</th>
                  
                  <th className="px-4 py-2">Request Date</th>
                  <th className="px-4 py-2">Request Name</th>
                  <th className="px-4 py-2">Vendor Name </th> 
                  <th className="px-4 py-2">Vendor Code</th>
                  
                  <th className="px-4 py-2">Invoice No.</th>
                  <th className="px-4 py-2">Amount</th>
                  <th className="px-4 py-2">Status</th>
                  
                </tr>
              </thead>
              <tbody>
                {!loading && totalItems === 0 ? (
                  <tr>
                    <td colSpan={9} className="empty">No matching records</td>
                  </tr>
                ) : (
                  paged.map((row) => {
                    const canEdit = row.status === "Pending for Approval";
                    return (
                      <tr key={row._spId ?? `${row.requestNo}-${row.invoiceNo}`}>
                        <td>
                          <div style={{ display: "flex", gap: 8, border: "0px solid #f8f3f3" }}>
                            <Link to={`/NonPoApproval/${row._spId}`}>
                            <img src={Edit} width={15} />
                          </Link>
                          &nbsp;&nbsp;
                          <Link to={`/ViewRequest/${row._spId}`}>
                            <img src={View} width={15} />
                          </Link>
                          </div>
                        </td>
                        <td>{row.requestNo}</td>
                       
                        <td>{row.requestDate}</td>
                         <td>{row.buyerName}</td>
                     <td>{row.vendorName}</td>
                        <td>{row.vendorCode}</td>
                        
                        <td>{row.invoiceNo}</td>
                        <td className="num">{row.amount.toLocaleString("en-IN")}</td>
                        <td>{row.status}</td>
                        
                       


                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>

      {/* Pagination footer */}
      <div className="myApproval-paginationRow" role="navigation" aria-label="Pagination">
        <div className="myApproval-pageInfo">
          {totalItems > 0 ? (
            <>Showing <b>{startIdx + 1}</b>–<b>{endIdx}</b> of <b>{totalItems}</b></>
          ) : (
            <>Showing 0 of 0</>
          )}
        </div>
        <div className="myApproval-pagerButtons">
          <button
            className="btnPager"
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            disabled={currentPage <= 1}
            aria-label="Previous page"
          >
            ← Prev
          </button>
          <span className="myApproval-pageCurrent">Page {currentPage} / {totalPages}</span>
          <button
            className="btnPager"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={currentPage >= totalPages}
            aria-label="Next page"
          >
            Next →
          </button>
        </div>
      </div>

      <style>{css}</style>
    </div>
  );
};

/** ---------------- CSS ---------------- **/
const css = `
.myApproval-root{padding:12px}
.myApproval-titleRow{padding:4px 2px 8px;border-bottom:1px solid #e5e7eb;margin-bottom:8px;display:flex;align-items:center}
.myApproval-title{margin:0;font-weight:700}
.myApproval-searchRow{margin:10px 0}
.myApproval-searchWrap{position:relative;width:260px}
.myApproval-searchWrap input{width:100%;padding:8px 30px;border:1px solid #e5e7eb;border-radius:8px}
.myApproval-searchIcon{position:absolute;left:8px;top:6px}
.myApproval-tableWrap{border:1px solid #e5e7eb;border-radius:10px;overflow:auto}
.myApproval-grid{width:100%;min-width:900px;border-collapse:separate}
.myApproval-grid thead th{background:#000;color:#fff;font-weight:700}
.myApproval-grid tbody td{padding:10px;border-bottom:1px solid #e5e7eb}
.num{text-align:right}
.empty{text-align:center;padding:20px}
.myApproval-paginationRow{display:flex;align-items:center;gap:12px;justify-content:space-between;margin:10px}
.myApproval-pageInfo{color:#374151}
.myApproval-pagerButtons{display:flex;align-items:center;gap:10px}
.btnPager{padding:6px 10px;border:1px solid #e5e7eb;border-radius:8px;background:#fff;color:#111827;cursor:pointer}
.btnPager:disabled{opacity:0.5;cursor:not-allowed}
.myApproval-pageCurrent{color:#4b5563}
`;

export default MyApproval;