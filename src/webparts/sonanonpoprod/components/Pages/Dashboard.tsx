// // Dashboard.tsx — shows Vendor Name, Service Type, Start Date, Renewal Date, Status
// import * as React from "react";
// import { useState, useEffect, useMemo } from "react";
// import SPCRUDOPS from "../../service/DAL/spcrudops";
// import { useHistory } from "react-router-dom";
// // import { ICustomerComWebPartProps } from "../ICustomerComWebPartProps";
// import "../Pages/CSS/NewRequest.scss";
// import { ISonaNonPoWebPartProps } from '../../SonaNonPoWebPart';

// // ⚠️ Notes:
// // - We select VendorNameslt & ServiceTypeslt from AMCDetail.
// // - We fetch AMCRenewal rows for the current page's AMCDetail IDs, then compute
// //   latest Start & Renewal dates per AMCRequestedID.
// // - Renew button shows ONLY when Status=Approved AND latest EndDate is within 30 days.
// // - Renew button color:
// //    <= 7 days  : Red
// //    <= 14 days : Yellow
// //    <= 30 days : Blue
// //
// // ✅ New in this version:
// // - Row highlight uses the same tone as Renew button (only when button is visible).
// // - Highlighted rows bubble to the top: red > yellow > blue > others, then days left asc, then ID desc.
// // - CSS classes + !important ensure highlight wins over any white backgrounds.
// // ✅ Fixes in this version:
// // - Expand/select Author so `Author/EMail` filter works reliably.
// // - Defensive guards when lists are empty.

// export const Dashboard: React.FC<ISonaNonPoWebPartProps> = (props: ISonaNonPoWebPartProps) => {
//   const [searchTerm, setSearchTerm] = useState("");
//   const [listData, setListData] = useState<any[]>([]);
//   // const [filteredData, setFilteredData] = useState<any[]>([]);
//   const [filteredData, setFilteredData] = useState(listData ?? []);
//   // const [currentPage, setCurrentPage] = useState(1);

//   const [loading, setLoading] = useState(true);

//   // Map of AMCDetail.ID -> latest renewal snapshot { startISO, renewalISO, endISO }
//   const [latestRenewalByMaster, setLatestRenewalByMaster] = useState<
//     Record<
//       number,
//       {
//         startISO?: string;
//         renewalISO?: string;
//         endISO?: string;
//       }
//     >
//   >({});

//   // Pagination
//   const itemsPerPage = 10;
//   const [currentPage, setCurrentPage] = useState(1);
//   const totalPages = Math.max(1, Math.ceil((filteredData?.length ?? 0) / itemsPerPage));
//   const history = useHistory();

//   // Treat "Approved" (case-insensitive, trimmed) as eligible for Renew
//   const isApproved = (status?: string) => (status ?? "").trim().toLowerCase() === "approved";

//   // Format date to DD/MM/YYYY safely
//   const formatDateToDDMMYYYY = (iso?: string) => {
//     if (!iso) return "";
//     const s = iso.includes("T") ? iso.split("T")[0] : iso;
//     const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
//     if (!m) return "";
//     const [, y, mm, dd] = m;
//     return `${dd}/${mm}/${y}`;
//   };

//   // Parse ISO yyyy-mm-dd into a LOCAL date (avoid timezone shifting issues)
//   const parseISOToLocalDate = (iso?: string) => {
//     if (!iso) return null;
//     const s = iso.includes("T") ? iso.split("T")[0] : iso;
//     const parts = s.split("-");
//     if (parts.length !== 3) return null;
//     const y = Number(parts[0]);
//     const m = Number(parts[1]);
//     const d = Number(parts[2]);
//     if (!y || !m || !d) return null;
//     return new Date(y, m - 1, d); // local midnight
//   };

//   // Days left until end date (endDate - today)
//   const getDaysLeft = (endISO?: string) => {
//     const end = parseISOToLocalDate(endISO);
//     if (!end) return null;

//     const today = new Date();
//     today.setHours(0, 0, 0, 0);
//     end.setHours(0, 0, 0, 0);

//     const diffMs = end.getTime() - today.getTime();
//     return Math.ceil(diffMs / (1000 * 60 * 60 * 24));
//   };

//   // Renew eligibility + color bucket (ignores status; we check status separately)
//   const getRenewMeta = (endISO?: string) => {
//     const daysLeft = getDaysLeft(endISO);

//     // Show only if end date exists and is within 0..30 days
//     if (daysLeft === null || daysLeft < 0 || daysLeft > 30) {
//       return { show: false, daysLeft: daysLeft ?? undefined, tone: "" as "" | "red" | "yellow" | "blue" };
//     }

//     if (daysLeft <= 7) return { show: true, daysLeft, tone: "red" as const };
//     if (daysLeft <= 14) return { show: true, daysLeft, tone: "yellow" as const };
//     return { show: true, daysLeft, tone: "blue" as const };
//   };

//   const getRenewColor = (tone: "" | "red" | "yellow" | "blue") => {
//     if (tone === "red") return "#dc2626"; // red
//     if (tone === "yellow") return "#f1f120"; // yellow
//     if (tone === "blue") return "#1d4ed8"; // blue
//     return "#111827";
//   };

//   // Tone priority for sorting
//   const getTonePriority = (tone: "" | "red" | "yellow" | "blue") => {
//     if (tone === "red") return 3;
//     if (tone === "yellow") return 2;
//     if (tone === "blue") return 1;
//     return 0;
//   };

//   // CSS class for row tone (we'll style these in SCSS with !important)
//   const getRowToneClass = (tone: "" | "red" | "yellow" | "blue") => {
//     if (tone === "red") return "row-tone-red";
//     if (tone === "yellow") return "row-tone-yellow";
//     if (tone === "blue") return "row-tone-blue";
//     return "";
//   };

//   // --- Utility: safe normalize to string ---
//   const norm = (v: any) => (v ?? "").toString().toLowerCase().trim();

//   const GetListData = async () => {
//     setLoading(true);
//     const spCrudOps = await SPCRUDOPS();
//     const userEmail = props.currentSPContext.pageContext.user.email;

//     // Step 1: Fetch parent list (AMCDetail) for current user.
//     // ✅ FIX: Include Author in select + expand so Author/EMail filter works.
//     const parentItems = await spCrudOps.getRootData(
//       "AMCDetail",
//       "ID,RequestedBy,Status,VendorNameslt,ServiceTypeslt,Author/EMail,Author/Title",
//       "Author",
//       `Author/EMail eq '${userEmail}'`,
//       { column: "ID", isAscending: false },
//       props
//     );

//     const safeParents = parentItems ?? [];
//     setListData(safeParents);
//     // setFilteredData(safeParents);

//     // Step 3: For Start & Renewal, fetch latest renewal per master ID.
//     await loadLatestRenewalsForPage(safeParents.map((x: any) => x.ID), spCrudOps);

//     setLoading(false);
//   };

//   // Helper: chunk IDs and retrieve renewals; then compute latest per parent
//   const loadLatestRenewalsForPage = async (masterIds: number[], spCrudOps: any) => {
//     if (!masterIds || masterIds.length === 0) {
//       setLatestRenewalByMaster({});
//       return;
//     }

//     // Build batched OR filters
//     const chunkSize = 20;
//     const chunks: number[][] = [];
//     for (let i = 0; i < masterIds.length; i += chunkSize) {
//       chunks.push(masterIds.slice(i, i + chunkSize));
//     }

//     const allRows: any[] = [];
//     for (const ids of chunks) {
//       const filter = ids.map((id) => `AMCRequestedIDId eq ${id}`).join(" or ");
//       const rows = await spCrudOps.getRootData(
//         "AMCRenewal",
//         "ID,AMCRequestedID/Id,AMCStartDate,AMCEndDate,AMCRenewalDate",
//         "AMCRequestedID",
//         filter,
//         { column: "ID", isAscending: false },
//         props
//       );
//       allRows.push(...(rows ?? []));
//     }

//     // Reduce rows per AMCRequestedID to "latest" by AMCRenewalDate (fallback to AMCEndDate)
//     const latestByParent: Record<number, { startISO?: string; renewalISO?: string; endISO?: string }> = {};

//     for (const r of allRows) {
//       const pid: number = r?.AMCRequestedID?.Id;
//       if (typeof pid !== "number") continue;

//       const sISO: string | undefined = r?.AMCStartDate ? String(r.AMCStartDate).split("T")[0] : undefined;
//       const eISO: string | undefined = r?.AMCEndDate ? String(r.AMCEndDate).split("T")[0] : undefined;
//       const rnISO: string | undefined = r?.AMCRenewalDate ? String(r.AMCRenewalDate).split("T")[0] : undefined;

//       const current = latestByParent[pid];
//       const dateToCompare = rnISO ?? eISO ?? "";
//       const existingCompare = current?.renewalISO ?? current?.endISO ?? "";
//       const newer =
//         existingCompare === "" || (dateToCompare !== "" && new Date(dateToCompare) > new Date(existingCompare));

//       if (!current || newer) {
//         latestByParent[pid] = {
//           startISO: sISO,
//           renewalISO: rnISO,
//           endISO: eISO,
//         };
//       }
//     }

//     setLatestRenewalByMaster(latestByParent);
//   };

//   // Keep dashboard in sync if another page updates status
//   useEffect(() => {
//     const handler = (e: any) => {
//       const { parentId, status } = e.detail || {};
//       if (!parentId) return;

//       setListData((prev) => prev.map((it: any) => (it.ID === parentId ? { ...it, Status: status } : it)));
//       setFilteredData((prev) => prev.map((it: any) => (it.ID === parentId ? { ...it, Status: status } : it)));
//     };

//     window.addEventListener("amc:status-updated", handler as any);
//     return () => window.removeEventListener("amc:status-updated", handler as any);
//   }, []);

//   // Initial load
//   useEffect(() => {
//     GetListData();
//     // eslint-disable-next-line react-hooks/exhaustive-deps
//   }, []);

//   // 🔍 UNIVERSAL SEARCH (across ID, Vendor, Service, RequestedBy, Status, Start/Renewal dates)
//   useEffect(() => {
//     const data = Array.isArray(listData) ? listData : [];

//     const q = norm(searchTerm);
//     if (!q) {
//       setFilteredData(data);
//       setCurrentPage(1);
//       return;
//     }

//     const filtered = data.filter((item) => {
//       if (!item || typeof item !== "object") return false;

//       // Basic fields from AMCDetail
//       const idStr = (item.ID ?? "").toString();
//       const vendor = item.VendorNameslt ?? "";
//       const service = item.ServiceTypeslt ?? "";
//       const requestedBy = item.RequestedBy ?? "";
//       const status = item.Status ?? "";

//       // Latest renewal snapshot dates (formatted as DD/MM/YYYY for matching what user sees)
//       const latest = latestRenewalByMaster[item.ID] || {};
//       const startDDMM = formatDateToDDMMYYYY(latest.startISO);
//       const renewalDDMM = formatDateToDDMMYYYY(latest.renewalISO);

//       const haystacks = [
//         idStr,
//         vendor,
//         service,
//         requestedBy,
//         status,
//         startDDMM,
//         renewalDDMM,
//       ].map(norm);

//       return haystacks.some((h) => h.includes(q));
//     });

//     setFilteredData(filtered);
//     setCurrentPage(1);
//   }, [searchTerm, listData, latestRenewalByMaster]); // ✅ depends on latest renewals too

//   // Pagination
//   const handlePageChange = (page: number) => {
//     if (page >= 1 && page <= totalPages) setCurrentPage(page);
//   };

//   // ⭐ Enrich + sort before paginate (so highlighted rows show on top)
//   const enriched = useMemo(() => {
//     return (filteredData ?? []).map((item) => {
//       const latest = latestRenewalByMaster[item.ID] || {};
//       const renewMeta = getRenewMeta(latest.endISO);

//       // Row highlight should match button visibility => only when Approved & within 30 days
//       const showButton = isApproved(item.Status) && renewMeta.show;
//       const effectiveTone = showButton ? renewMeta.tone : ("" as "" | "red" | "yellow" | "blue");
//       const tonePriority = getTonePriority(effectiveTone);

//       return { item, latest, renewMeta, effectiveTone, showButton, tonePriority };
//     });
//   }, [filteredData, latestRenewalByMaster]);

//   const enrichedSorted = useMemo(() => {
//     return [...enriched].sort((a, b) => {
//       // 1) Tone priority: red > yellow > blue > none
//       if (b.tonePriority !== a.tonePriority) return b.tonePriority - a.tonePriority;

//       // 2) Within same tone, fewer days left first
//       const aDL = a.showButton ? (a.renewMeta.daysLeft ?? Number.POSITIVE_INFINITY) : Number.POSITIVE_INFINITY;
//       const bDL = b.showButton ? (b.renewMeta.daysLeft ?? Number.POSITIVE_INFINITY) : Number.POSITIVE_INFINITY;
//       if (aDL !== bDL) return aDL - bDL;

//       // 3) Fallback by ID desc
//       return b.item.ID - a.item.ID;
//     });
//   }, [enriched]);

//   const paginatedData = useMemo(
//     () => enrichedSorted.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage),
//     [enrichedSorted, currentPage]
//   );

//   return (
//     <div className="min-h-screen bg-gray-100">
//       {/* Header */}
//       <div className="header header-bar">
//         <div className="left-banner">
//           <div className="logo-text">
//             <h2>All My AMC</h2>
//           </div>
//         </div>
//       </div>

//       <main className="Main-Dash">
//         {loading ? (
//           <div className="loading-overlay">
//             <div className="loading-content">
//               <svg className="loading-spinner" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
//                 <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
//                 <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
//               </svg>
//               <p className="text-white text-lg">Please wait, loading data...</p>
//             </div>
//           </div>
//         ) : (
//           <>
//             {/* Search Bar + New AMC Button */}
//             <div className="Dashboard-Search flex items-center gap-2 mb-4">
//               {/* Search Input */}
//               <input
//                 type="text"
//                 style={{ marginTop: "2px" }}
//                 placeholder="Search..."
//                 value={searchTerm}
//                 onChange={(e) => setSearchTerm(e.target.value)}
//                 className="w-64 text-sm border-gray-300 rounded-full dashboard-sha focus:outline-none focus:ring-2 focus:ring-red-500"
//               />

//               {/* New AMC Button */}
//               <button
//                 onClick={() => history.push("/newAMC")}
//                 className="btn btn-sm btn-dark ml-auto"
//                 style={{ marginLeft: "600px" }}
//               >
//                 New AMC
//               </button>
//             </div>

//             {/* Table */}
//             <div className="overflow-x-auto -ml-4">
//               <div className="table-vert-scroll">
//                 <div className="table-grid shadow-md">
//                   <table className="min-w-full bg-white rounded-2xl shadow-md">
//                     <thead style={{ backgroundColor: "#3c3e45" }} className="text-white">
//                       <tr>
//                         <th className="px-4 py-2">ID</th>
//                         <th className="px-4 py-2">Vendor Name</th>
//                         <th className="px-4 py-2">Service Type</th>
//                         <th className="px-4 py-2">Start Date</th>
//                         <th className="px-4 py-2">Renewal Date</th>
//                         <th className="px-4 py-2">Status</th>
//                         <th className="px-4 py-2">Action</th>
//                       </tr>
//                     </thead>

//                     <tbody>
//                       {paginatedData.map((row, index) => {
//                         const { item, latest, renewMeta, effectiveTone, showButton } = row;
//                         const startDDMM = formatDateToDDMMYYYY(latest.startISO);
//                         const renewalDDMM = formatDateToDDMMYYYY(latest.renewalISO);

//                         const renewColor = getRenewColor(renewMeta.tone); // button color
//                         const toneClass = getRowToneClass(effectiveTone);

//                         return (
//                           <tr
//                             key={item.ID ?? index}
//                             className={`border-t ${toneClass}`} // ✅ class on TR (CSS below applies to all TD)
//                           >
//                             <td className={`px-4 py-2`}>{item.ID}</td>
//                             <td className={`px-4 py-2`}>{item.VendorNameslt ?? "-"}</td>
//                             <td className={`px-4 py-2`}>{item.ServiceTypeslt ?? "-"}</td>
//                             <td className={`px-4 py-2`}>{startDDMM || "-"}</td>
//                             <td className={`px-4 py-2`}>{renewalDDMM || "-"}</td>
//                             <td className={`px-4 py-2`}>{item.Status ?? "-"}</td>

//                             <td className={`px-2 py-1 text-center`}>
//                               {["draft", "rfi"].includes(item.Status?.trim?.().toLowerCase?.() || "") ? (
//                                 <button
//                                   type="button"
//                                   style={{
//                                     marginRight: "95px",
//                                     border: "none",
//                                     outline: "none",
//                                     background: "none",
//                                   }}
//                                   onClick={() => history.push(`/editDraft/${item.ID}`)}
//                                   title={
//                                     item.Status?.toLowerCase() === "rfi"
//                                       ? "Edit & Resubmit (Pending)"
//                                       : "Edit Draft"
//                                   }
//                                 >
//                                   <i className="fas fa-edit NAMC-action-icon edit-icon"></i>
//                                 </button>
//                               ) : (
//                                 <div className="d-flex gap-1">
//                                   {/* VIEW BUTTON */}
//                                   <button
//                                     type="button"
//                                     style={{
//                                       border: "none",
//                                       outline: "none",
//                                       background: "none",
//                                     }}
//                                     onClick={() => history.push(`/viewAMC/${item.ID}`)}
//                                     title="View"
//                                   >
//                                     <i className="fas fa-eye NAMC-action-icon edit-icon"></i>
//                                   </button>

//                                   {/* RENEW BUTTON: visible only when Approved AND within 0..30 days */}
//                                   {showButton && (
//                                     <button
//                                       type="button"
//                                       className="NAMC-action-icon renew-icon"
//                                       style={{
//                                         padding: "2px 6px",
//                                         fontSize: "15px",
//                                         lineHeight: "12px",
//                                         height: "22px",
//                                         border: "none",
//                                         outline: "none",
//                                         background: "none",
//                                         color: renewColor, // matches tone
//                                         fontWeight: 700,
//                                       }}
//                                       onClick={() => history.push(`/renewAMC/${item.ID}`)}
//                                       title={`Renew (Ends in ${renewMeta.daysLeft} day(s))`}
//                                     >
//                                       <i className="fas fa-sync"></i> Renew
//                                     </button>
//                                   )}
//                                 </div>
//                               )}
//                             </td>
//                           </tr>
//                         );
//                       })}
//                     </tbody>
//                   </table>
//                 </div>

//                 {/* Pagination */}
//                 <div className="flex justify-center mt-6 overflow-x-auto">
//                   <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow">
//                     {/* Previous */}
//                     <button
//                       onClick={() => handlePageChange(currentPage - 1)}
//                       disabled={currentPage === 1}
//                       style={{
//                         backgroundColor: "#3c3e45",
//                         color: "White",
//                         opacity: currentPage === 1 ? 0.5 : 1,
//                       }}
//                       className="px-3 py-1 border rounded"
//                     >
//                       Previous
//                     </button>

//                     {/* First Page Shortcut */}
//                     {currentPage > 3 && (
//                       <>
//                         <button
//                           onClick={() => handlePageChange(1)}
//                           style={{ backgroundColor: "#3c3e45", color: "White" }}
//                           className="px-3 py-1 border rounded"
//                         >
//                           1
//                         </button>
//                         <span className="px-2">...</span>
//                       </>
//                     )}

//                     {/* Main Page Numbers */}
//                     {Array.from({ length: totalPages }, (_, i) => i + 1)
//                       .filter((page) => Math.abs(page - currentPage) <= 2)
//                       .map((page) => (
//                         <button
//                           key={page}
//                           onClick={() => handlePageChange(page)}
//                           style={{
//                             backgroundColor: "#3c3e45",
//                             color: "White",
//                             fontWeight: currentPage === page ? "bold" : "normal",
//                           }}
//                           className="px-3 py-1 border rounded"
//                         >
//                           {page}
//                         </button>
//                       ))}

//                     {/* Last Page Shortcut */}
//                     {currentPage < totalPages - 2 && (
//                       <>
//                         <span className="px-2">...</span>
//                         <button
//                           onClick={() => handlePageChange(totalPages)}
//                           style={{ backgroundColor: "#3c3e45", color: "White" }}
//                           className="px-3 py-1 border rounded"
//                         >
//                           {totalPages}
//                         </button>
//                       </>
//                     )}

//                     {/* Next */}
//                     <button
//                       onClick={() => handlePageChange(currentPage + 1)}
//                       disabled={currentPage === totalPages}
//                       style={{
//                         backgroundColor: "#3c3e45",
//                         color: "White",
//                         opacity: currentPage === totalPages ? 0.5 : 1,
//                       }}
//                       className="px-3 py-1 border rounded"
//                     >
//                       Next
//                     </button>
//                   </div>
//                 </div>
//               </div>
//             </div>
//           </>
//         )}
//       </main>
//     </div>
//   );
// };

// export default Dashboard;