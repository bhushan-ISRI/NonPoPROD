import React, { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";

import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import SPCRUDOPS from "../../service/DAL/spcrudops";
import type { ISiteUserInfo } from "@pnp/sp/site-users/types";

export type NonPoStatus =
  | "Pending for Approval"
  | "Sent Back"
  | "Approved"
  | "Rejected"
  | "Closed"
  | string;

export interface NonPoRow {
  requestNo: string;
  requestDate: string;
  vendorCode: string;
  vendorName: string;
  invoiceNo: string;
  amount: number;
  status: NonPoStatus;
  _spId?: number;
}

interface BaseProps {
  listTitle?: string;
  onAddNon?: () => void;
  onViewRow?: (row: NonPoRow) => void;
  title?: string;

  statusEquals?: NonPoStatus | NonPoStatus[];

  disableAuthorFilter?: boolean;
}

type InitiatorBaseTableProps = ISonanonpoprodProps & BaseProps;

async function resolveCurrentUserId(ctx: any): Promise<number | undefined> {
  const legacyId: number | undefined = ctx?.pageContext?.legacyPageContext?.userId;
  if (legacyId) return legacyId;

  try {
    const { spfi } = await import("@pnp/sp");
    const { SPFx } = await import("@pnp/sp/behaviors/spfx");
    const sp = spfi().using(SPFx(ctx));
    const me: ISiteUserInfo = await sp.web.currentUser();
    if (me?.Id) return me.Id;
  } catch (Error) { alert("Error") }

  try {
    const loginName: string | undefined = ctx?.pageContext?.user?.loginName;
    if (loginName) {
      const { spfi } = await import("@pnp/sp");
      const { SPFx } = await import("@pnp/sp/behaviors/spfx");
      const sp = spfi().using(SPFx(ctx));
      const ensured: any = await sp.web.ensureUser(loginName);
      const ensuredId = ensured?.data?.Id;
      if (ensuredId) return ensuredId;
    }
  } catch (Error) { alert("Error") }

  return undefined;
}

/**  OData string literals */
const escapeODataString = (s: string) => (s ?? "").replace(/'/g, "''");

/**
 */
const buildStatusFilter = (value?: NonPoStatus | NonPoStatus[]) => {
  if (!value || (Array.isArray(value) && value.length === 0)) return "";
  const arr = Array.isArray(value) ? value : [value];
  const parts = arr.map((s) => `Status eq '${escapeODataString(String(s))}'`);
  return parts.length ? `(${parts.join(" or ")})` : "";
};

const InitiatorBaseTable: React.FC<InitiatorBaseTableProps> = ({
  listTitle = "NonPO",
  onAddNon,
  onViewRow,
  title = "NON PO",
  statusEquals,
  disableAuthorFilter = false,
  ...props
}) => {
  const navigate = useNavigate();

  const [query, setQuery] = useState("");
  const [rows, setRows] = useState<NonPoRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  // Pagination
  const PAGE_SIZE = 10;
  const [currentPage, setCurrentPage] = useState(1);

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

  useEffect(() => {
    let abort = false;

    (async () => {
      try {
        setLoading(true);
        setError("");

        const spCrudOps = await SPCRUDOPS(props.currentSPContext);

        let userId: number | undefined = undefined;
        if (!disableAuthorFilter) {
          userId = await resolveCurrentUserId(props.currentSPContext);
          if (!userId) {
            throw new Error(
              "User context not found. Hosted workbench (/_layouts/15/workbench.aspx)."
            );
          }
        }

        const select =
          "Id,Title,RequestDate,VendorName,InvoiceNumber,Totalamount,Amount,Status,VendorCode/Title";
        const expand = "VendorCode";

        // --- Build OData filter ---
        const clauses: string[] = [];
        if (!disableAuthorFilter && userId) {
          clauses.push(`AuthorId eq ${userId}`);
        }
        const statusClause = buildStatusFilter(statusEquals);
        if (statusClause) clauses.push(statusClause);

        const filter = clauses.length ? clauses.map((c) => `(${c})`).join(" and ") : "";

        const orderBy = { column: "Id", isAscending: false };

        const items =
          (await spCrudOps.getRootData(
            listTitle,
            select,
            expand,
            filter,
            orderBy,
            props
          )) ?? [];

        if (abort) return;

        const mapped: NonPoRow[] = items.map((it: any) => ({
          _spId: it?.Id,
          requestNo: it?.Title || "",
          requestDate: fmtDate(it?.RequestDate),
          vendorCode: it?.VendorCode?.Title || "",
          vendorName: it?.VendorName || "",
          invoiceNo: it?.InvoiceNumber || "",
          amount:
            it?.Totalamount !== undefined && it?.Totalamount !== null
              ? safeNum(it?.Totalamount)
              : safeNum(it?.Amount),
          status: it?.Status || "Pending for Approval",
        }));

        // Final client-side clamp for requested statuses (handles casing/whitespace)
        const normalize = (s: any) => String(s ?? "").trim().toLowerCase();
        const normalizedTargets = Array.isArray(statusEquals)
          ? statusEquals.map(normalize)
          : statusEquals
            ? [normalize(statusEquals)]
            : [];

        const finalRows =
          normalizedTargets.length > 0
            ? mapped.filter((r) => normalizedTargets.includes(normalize(r.status)))
            : mapped;

        setRows(finalRows);
      } catch (e: any) {
        console.error("NonPO fetch error:", e);
        setError(e?.message || "Failed to load data");
      } finally {
        if (!abort) setLoading(false);
      }
    })();

    return () => {
      abort = true;
    };
  }, [listTitle, statusEquals, disableAuthorFilter, props.currentSPContext]);

  const filtered = useMemo(() => {
    if (!query.trim()) return rows;
    const q = query.toLowerCase();
    return rows.filter((r) =>
      [
        r.requestNo,
        r.requestDate,
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

  useEffect(() => {
    setCurrentPage(1);
  }, [query, rows]);

  const totalItems = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  const page = Math.min(currentPage, totalPages);
  const startIdx = (page - 1) * PAGE_SIZE;
  const endIdx = Math.min(startIdx + PAGE_SIZE, totalItems);
  const paged = filtered.slice(startIdx, endIdx);

  const handleAddNon = () => {
    if (onAddNon) onAddNon();
    else navigate("/NonPoRequest?mode=new");
  };

  const handleView = (row: NonPoRow) => {
    if (onViewRow) return onViewRow(row);
    if (!row?._spId) return;
    navigate(`/NonPoRequest?id=${row._spId}&mode=view`, {
      state: { from: "/InitiatorLanding", fullscreen: true },
    });
  };

  const handleEdit = (row: NonPoRow) => {
    if (!row?._spId) return;
    navigate(`/NonPoRequest?id=${row._spId}&mode=edit`, {
      state: { from: "/InitiatorLanding", fullscreen: true },
    });
  };

  const goPrev = () => setCurrentPage((p) => Math.max(1, p - 1));
  const goNext = () => setCurrentPage((p) => Math.min(totalPages, p + 1));

  return (
    <div className='MainUplodForm' style={{ margin: "0px" }}>
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
                  <th className="px-4 py-2">Vendor Code</th>
                  <th className="px-4 py-2">Vendor Name </th>
                  <th className="px-4 py-2">Invoice No.</th>
                  <th className="px-4 py-2">Amount</th>
                  <th className="px-4 py-2">Status</th>
                  
                </tr>
              </thead>
            </table>
            <tbody>
            {!loading && totalItems === 0 ? (
              <tr>
                <td colSpan={8} className="empty">
                  No records found
                </td>
              </tr>
            ) : (
              paged.map((row) => (
                <tr key={row._spId ?? `${row.requestNo}-${row.invoiceNo}`}>
                  <td>
                    <div style={{ display: "flex", gap: 8, border: "0px solid #f8f3f3" }}>
                      <a
                        style={{ border: "0px solid #f8f3f3", background: "white" }}
                        onClick={() => handleView(row)}
                        title="View"
                        aria-label={`View ${row.requestNo}`}
                      >
                        <i className="fas fa-eye NAMC-action-icon edit-icon"></i>
                      </a>

                    </div>
                  </td>
                  <td>{row.requestNo}</td>
                  <td>{row.requestDate}</td>
                  <td>{row.vendorCode}</td>
                  <td>{row.vendorName}</td>
                  <td>{row.invoiceNo}</td>
                  <td className="num">{row.amount.toLocaleString("en-IN")}</td>
                  <td>{row.status}</td>
                  
                </tr>
              ))
            )}
          </tbody>
          </div>
        </div>
      </main>

      {/* Pagination Footer */}
      <div className="myApproval-paginationRow" role="navigation" aria-label="Pagination">
        <div className="myApproval-pageInfo">
          {totalItems > 0 ? (
            <>
              Showing <b>{startIdx + 1}</b>–<b>{endIdx}</b> of <b>{totalItems}</b>
            </>
          ) : (
            <>Showing 0 of 0</>
          )}
        </div>
        <div className="myApproval-pagerButtons">
          <button className="btnPager" onClick={goPrev} disabled={page <= 1} aria-label="Previous page">
            ← Prev
          </button>
          <span className="myApproval-pageCurrent">Page {page} / {totalPages}</span>
          <button className="btnPager" onClick={goNext} disabled={page >= totalPages} aria-label="Next page">
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
.myApproval-titleRow{padding:4px 2px 8px;border-bottom:1px solid #e5e7eb;margin-bottom:8px;display:flex;align-items:center;gap:10px}
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
  padding:9px;
}
.myApproval-grid tbody td{padding:10px;border-bottom:1px solid #e5e7eb}
.num{text-align:right}
.empty{text-align:center;padding:20px}
.btnPrimary{padding:8px 14px;border:0;border-radius:8px;background:#111827;color:#fff;font-weight:700;cursor:pointer}
.myApproval-paginationRow{display:flex;align-items:center;gap:12px;justify-content:space-between;margin:10px}
.myApproval-pageInfo{color:#374151}
.myApproval-pagerButtons{display:flex;align-items:center;gap:10px}
.btnPager{padding:6px 10px;border:1px solid #e5e7eb;border-radius:8px;background:#fff;color:#111827;cursor:pointer}
.btnPager:disabled{opacity:0.5;cursor:not-allowed}
.myApproval-pageCurrent{color:#4b5563}
`;

export default InitiatorBaseTable;