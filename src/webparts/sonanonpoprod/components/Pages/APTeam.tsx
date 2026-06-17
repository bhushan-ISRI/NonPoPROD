import React, { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";

import { Link } from "react-router-dom";
import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";

import SPCRUDOPS from "../../service/DAL/spcrudops";
import useFullscreenForm from "../../hook/useFullscreenForm";
import Edit from "../../components/Pages/Image/Pencil.png"
import View from "../../components/Pages/Image/Eye.png";

export interface APTeamRow {
  _spId: number;
  requestNo: string;
  requestDate: string;
  buyerName: string;
  vendorCode: string;
  vendorName: string;
  invoiceNo: string;
  amount: number;
  status?: string;
  CurrentApproverEmail: string;
}

interface APTeamProps extends ISonanonpoprodProps {
  listTitle?: string;
  title?: string;
}

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

const APTeam: React.FC<APTeamProps> = ({
  currentSPContext,
  listTitle = "NonPO",
  title = "NON PO Acceptance Dashboard",
  //onView,
  ...props
}) => {
  useFullscreenForm();
  const navigate = useNavigate();

  const [query, setQuery] = useState("");
  const [rows, setRows] = useState<APTeamRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  /** ---------- Pagination (10 per page) ---------- **/
  const PAGE_SIZE = 10;
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    let abort = false;

    (async () => {
      try {
        setLoading(true);
        setError("");

        const spCrudOps = await SPCRUDOPS(currentSPContext);

        // Columns 
        const select =
          "Id,Title,RequestDate,EmployeeName,VendorCode/VendorCode,VendorCode/Id,VendorName,InvoiceNumber,Totalamount,Amount,Status,VendorCode/Title,CurrentApprover/Id,CurrentApprover/EMail,CurrentApprover/Title";
        const expand = "VendorCode,CurrentApprover";
        const orderBy = { column: "Id", isAscending: false };

        const filter = `Status eq 'Pending for Acceptance'`;

        const items =
          (await spCrudOps.getRootData(
            listTitle,
            select,
            expand,
            filter,
            orderBy,
            { currentSPContext } as ISonanonpoprodProps
          )) ?? [];

        if (abort) return;

        const mapped: APTeamRow[] = items.map((it: any) => ({
          _spId: it?.Id,
          requestNo: it?.Title || "",
          requestDate: fmtDate(it?.RequestDate),
          buyerName: it?.EmployeeName || "",
          vendorCode: it?.VendorCode?.VendorCode || "",
          vendorName: it?.VendorName || "",
          invoiceNo: it?.InvoiceNumber || "",
          CurrentApproverEmail: it?.CurrentApprover?.EMail,
          amount:
            it?.Totalamount !== undefined && it?.Totalamount !== null && it?.Totalamount !== ""
              ? safeNum(it?.Totalamount)
              : safeNum(it?.Amount),
          status: it?.Status || "",
        }));

        let filteredmapped = mapped.filter((m) => (m.CurrentApproverEmail === props.userEmail))
        setRows(filteredmapped);
      } catch (e: any) {
        console.error("[APTeam] fetch error:", e);
        setError(e?.message || "Failed to load AP Team data");
      } finally {
        if (!abort) setLoading(false);
      }
    })();

    return () => { abort = true; };
  }, [listTitle, currentSPContext]);

  /** ---------- Search ---------- **/
  const filtered = useMemo(() => {
    const onlyPendingForAcceptance = rows.filter(
      r => (r.status || "").trim() === "Pending for Acceptance"
    );

    if (!query.trim()) return onlyPendingForAcceptance;

    const q = query.toLowerCase();
    return onlyPendingForAcceptance.filter((row) =>
      [
        row.requestNo,
        row.requestDate,
        row.buyerName,
        row.vendorCode,
        row.vendorName,
        row.invoiceNo,
        String(row.amount),
        row.status || "",
      ]
        .join(" ")
        .toLowerCase()
        .includes(q)
    );
  }, [rows, query]);

  /** ---------- Reset page on data/search change ---------- **/
  useEffect(() => {
    setCurrentPage(1);
  }, [rows, query]);

  /** ---------- Pagination derived values ---------- **/
  const totalItems = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  const page = Math.min(currentPage, totalPages);
  const startIdx = (page - 1) * PAGE_SIZE;
  const endIdx = Math.min(startIdx + PAGE_SIZE, totalItems);
  const paged = filtered.slice(startIdx, endIdx);

  /**  View / Edit navigate to APTeamAcceptance */
  const handleView = (row: APTeamRow) => {
    if (!row?._spId) {
      alert("Item ID missing for this row.");
      return;
    }
    navigate(`/APTeamAcceptance?id=${row._spId}&mode=view`, { state: { from: "/APTeam", fullscreen: true } });
  };

  const handleEdit = (row: APTeamRow) => {
    if (!row?._spId) {
      alert("Item ID missing for this row.");
      return;
    }
    navigate(`/APTeamAcceptance?id=${row._spId}&mode=edit`, { state: { from: "/APTeam", fullscreen: true } });
  };

  return (
    <div>
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2> {title} </h2>
          </div>
        </div>
      </div>
      <div className='col-md-12 px-2 d-flex justify-content-between align-items-center flex-wrap' style={{ margin: "5px" }}>
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
                  <th className="px-4 py-2">Vendor Name</th>
                  <th className="px-4 py-2">Vendor Code </th>
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
                  paged.map((row) => (
                    <tr key={row._spId ?? `${row.requestNo}-${row.invoiceNo}`}>
                      <td>
                        
                          <Link to={`/APTeamAcceptance/${row._spId}`}>
                            <img src={Edit} width={15} />
                          </Link>
                          &nbsp;&nbsp;
                          <Link to={`/ViewRequest/${row._spId}`}>
                            <img src={View} width={15} />
                          </Link>
                        
                      </td>
                      <td>{row.requestNo}</td>
                      <td>{row.requestDate}</td>
                      <td>{row.buyerName}</td>
                      <td>{row.vendorName}</td>
                      <td>{row.vendorCode}</td>

                      <td>{row.invoiceNo}</td>
                      <td className="num">{row.amount.toLocaleString("en-IN")}</td>
                      <td>{"Pending for Acceptance"}</td> {/* force-label, to avoid showing anything else */}

                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>
      {/*  Pagination footer */}
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
            disabled={page <= 1}
            aria-label="Previous page"
          >
            ← Prev
          </button>
          <span className="myApproval-pageCurrent">Page {page} / {totalPages}</span>
          <button
            className="btnPager"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={page >= totalPages}
            aria-label="Next page"
          >
            Next →
          </button>
        </div>
      </div>

      <style>{css}</style>
    </div >
  );
};

/** ======================= CSS ======================== **/
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
.myApproval-grid thead th {
  background: #000 !important;
  color: #fff !important;
  font-weight: 700;
}
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

export default APTeam;