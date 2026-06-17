// import * as React from "react";
// import { useEffect, useMemo, useRef, useState } from "react";
// import { useParams, useHistory } from "react-router-dom";
// import { Formik, Form, Field, FormikProps } from "formik";
// import SPCRUDOPS from "../../service/DAL/spcrudops";
// import type { ISonanonpoprodProps } from "../ISonanonpoprodProps";
// import { sp } from "@pnp/sp/presets/all";
// import { Web } from "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "../Pages/CSS/NewRequest.scss";
// import "bootstrap/dist/css/bootstrap.min.css";
// import "bootstrap/dist/js/bootstrap.bundle.min.js";
// import { ISonaNonPoWebPartProps } from "../../SonaNonPoWebPart";

// /* -------- Types -------- */
// interface AMCDetailForm {
//   requestedBy: string;
//   requestorEmail: string;
//   requestDate: string; // yyyy-mm-dd

//   //  NewRequest/EditDraft style
//   requestorLocation: string;
//   Description: string;

//   locationsText: string;

//   amcType: string;
//   contractType: string;
//   serviceType: string;
//   vendorName: string;
//   Name: string;
//   ApproverRemark?: string;
// }

// interface ExistingDetailRow {
//   id: number;
//   amcStartDate: string;
//   amcEndDate: string;
//   amcRenewalDate: string;
//   estimatedAMC: string;
//   quantity: string;
//   responsibilityName?: string;
//   description: string;
//   rStatus?: string;
//   rName?: string;

//   // NEW: for grid — last Initiator remark found in AMCRenewal.ApproverRemark
//   initialRemark?: string;
// }

// interface IAttachmentItem {
//   Name: string;
//   ServerRelativeUrl: string;
//   TimeCreated?: string;
//   Length?: number;
// }

// interface IRouteParams {
//   id: string;
// }

// interface IRemark {
//   id: number;
//   approver?: string;
//   remark?: string;
//   created?: string; // yyyy-mm-dd or yyyy-mm-dd HH:mm
//   sourceList?: string;
//   status?: string;
//   role?: string;    // "Initiator" | "Approver"
// }

// /**
//  * Attachments library
//  */
// const AMC_SITE_ABS = "https://sonacomstargroup.sharepoint.com/sites/AMC_UAT";
// const ATTACH_LIB_NAME = "AMCAttachments";

// export const ViewNonPo: React.FC<ISonaNonPoWebPartProps> = (props) => {
//   const { id } = useParams<IRouteParams>();
//   const amcId = Number(id);
//   const history = useHistory();
//   const formikRef = useRef<FormikProps<AMCDetailForm>>(null);

//   const [loading, setLoading] = useState<boolean>(true);

//   /* ---- Master ---- */
//   const [masterInitial, setMasterInitial] = useState<AMCDetailForm>({
//     requestedBy: "",
//     requestorEmail: "",
//     requestDate: "",
//     requestorLocation: "",
//     Description: "",
//     locationsText: "",
//     amcType: "",
//     contractType: "",
//     serviceType: "",
//     vendorName: "",
//     Name: "",
//     ApproverRemark: "",
//   });

//   /* ---- Existing details ---- */
//   const [existingDetails, setExistingDetails] = useState<ExistingDetailRow[]>([]);

//   /* ---- Attachments per renewal row ---- */
//   const [attachmentsByRenewalId, setAttachmentsByRenewalId] = useState<Record<number, IAttachmentItem[]>>({});
//   const [attLoading, setAttLoading] = useState<boolean>(false);

//   //  Modal state for Action (eye)
//   const [showDetailsModal, setShowDetailsModal] = useState<boolean>(false);
//   const [selectedRow, setSelectedRow] = useState<ExistingDetailRow | null>(null);

//   // Approver remarks (Parent + Child)
//   const [remarksLoading, setRemarksLoading] = useState<boolean>(false);
//   const [parentRemarks, setParentRemarks] = useState<IRemark[]>([]);
//   const [childRemarks, setChildRemarks] = useState<IRemark[]>([]);

//   // Setup sp context
//   useEffect(() => {
//     sp.setup({ spfxContext: props.context as any });
//   }, [props.context]);

//   // AMC site web (attachments library )
//   const amcWeb = useMemo(() => Web(AMC_SITE_ABS), []);

//   //=============Comment history helper ==========================

//   const nowIST = (): string => {
//     const f = new Intl.DateTimeFormat("en-GB", {
//       timeZone: "Asia/Kolkata",
//       year: "numeric",
//       month: "2-digit",
//       day: "2-digit",
//       hour: "2-digit",
//       minute: "2-digit",
//       hour12: false,
//     });
//     const parts = Object.fromEntries(f.formatToParts(new Date()).map(p => [p.type, p.value]));
//     return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}`;
//   };

//   // multiline / pipe etc. ko clean kar de
//   const sanitizeRemark = (s: string) =>
//     (s ?? "")
//       .replace(/\r?\n/g, " ")
//       .replace(/\s+/g, " ")
//       .trim();

//   // Robust HTML stripping (raw <div>... and encoded &lt;div&gt;...)
//   const stripHtml = (s: string) =>
//     (s ?? "")
//       .replace(/<[^>]+>/g, "")
//       .replace(/&lt;[^&gt;]+&gt;/g, "")
//       .trim();

//   /**
//    * Parse remark log with role/status awareness
//    *  A) [YYYY-MM-DD HH:mm] Name (email): comment
//    *     → role inferred by sourceList (AMCDetail = Initiator, AMCRenewal = Approver), status=""
//    *  B) YYYY-MM-DD HH:mm | TOKEN | Name: comment
//    *     - if TOKEN is Initiator/Approver ⇒ role=TOKEN, status=""
//    *     - else TOKEN is workflow status (RFI/Approved/Reject/...), role inferred by source
//    */
//   const parseRemarkLog = (raw: string, sourceList: string): IRemark[] => {
//     if (!raw?.trim()) return [];

//     const log = stripHtml(raw);
//     const lines = log.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

//     const rxA = /^\[(\d{4}-\d{2}-\d{2})(?:\s+(\d{2}:\d{2}))?\]\s*([^:(]+?)(?:\s+\(([^)]+)\))?:\s*(.*)$/;
//     const rxB = /^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})\s*\|\s*([^|]+)\s*\|\s*([^:]+):\s*(.*)$/;

//     const defaultRoleBySource = (src: string) => (src === "AMCDetail" ? "Initiator" : "Approver");
//     const normalizeStatus = (s: string) => {
//       const t = (s || "").trim().toLowerCase();
//       if (!t) return "";
//       if (["approved", "approve"].includes(t)) return "Approved";
//       if (["rejected", "reject"].includes(t)) return "Reject";
//       if (["rfi", "info", "information"].includes(t)) return "RFI";
//       return (s || "").trim();
//     };

//     return lines.map((line, idx) => {
//       // Try B first (token format)
//       let m = line.match(rxB);
//       if (m) {
//         const [, ymd, hm, token, name, comment] = m;
//         const tokenClean = (token ?? "").trim();
//         const lc = tokenClean.toLowerCase();
//         const isRoleToken = lc === "initiator" || lc === "approver";
//         const role = isRoleToken ? tokenClean : defaultRoleBySource(sourceList);
//         const status = isRoleToken ? "" : normalizeStatus(tokenClean);
//         return {
//           id: idx + 1,
//           approver: (name ?? "").trim(),
//           remark: (comment ?? "").trim(),
//           created: `${ymd} ${hm}`,
//           sourceList,
//           role,
//           status,
//         } as IRemark;
//       }

//       // Try A (bracket format)
//       m = line.match(rxA);
//       if (m) {
//         const [, ymd, hm, name, email, comment] = m;
//         return {
//           id: idx + 1,
//           approver: email ? `${name.trim()} (${email.trim()})` : name.trim(),
//           remark: (comment ?? "").trim(),
//           created: hm ? `${ymd} ${hm}` : ymd,
//           sourceList,
//           role: defaultRoleBySource(sourceList),
//           status: "",
//         } as IRemark;
//       }

//       // Fallback
//       return { id: idx + 1, approver: undefined, remark: line, created: undefined, sourceList, role: "", status: "" } as IRemark;
//     });
//   };

//   /** ---------------------------------------------------------
//    * Append an entry to AMCRenewal.ApproverRemark (append-only)
//    * Format: [YYYY-MM-DD HH:mm] Full Name (email): Comment
//    * --------------------------------------------------------*/
//   const appendRenewalRemark = async (
//     renewalId: number,
//     newComment: string,
//     props: ISonaNonPoWebPartProps
//   ) => {
//     if (!renewalId || !newComment?.trim()) return;

//     // 1) Current user identity
//     const me = await sp.web.currentUser.select("Title", "Email")();
//     const who = `${me?.Title ?? "Unknown"}${me?.Email ? ` (${me.Email})` : ""}`;

//     // 2) New entry (IST + sanitized text)
//     const entry = `[${nowIST()}] ${who}: ${sanitizeRemark(newComment)}`;

//     // 3) Read current ApproverRemark
//     const spCrudOps = await SPCRUDOPS();
//     const rows = await spCrudOps.getRootData(
//       "AMCRenewal",
//       "ID,ApproverRemark",
//       "",
//       `ID eq ${renewalId}`,
//       { column: "ID", isAscending: true },
//       props
//     );
//     const current = (rows?.[0]?.ApproverRemark ?? "").toString().trim();

//     // 4) Merge & update
//     const finalLog = current ? `${current}\n${entry}` : entry;

//     // Use your DAL if it exposes updateItem; else PnPjs fallback
//     if ((spCrudOps as any).updateItem) {
//       await (spCrudOps as any).updateItem("AMCRenewal", renewalId, { ApproverRemark: finalLog }, props);
//     } else {
//       await sp.web.lists.getByTitle("AMCRenewal").items.getById(renewalId).update({ ApproverRemark: finalLog });
//     }
//   };

//   /** ---------------------------------------------------------
//    * Append an entry to AMCDetail.ApproverRemark (append-only)
//    * Same format as above
//    * --------------------------------------------------------*/
//   const appendDetailRemark = async (
//     amcDetailId: number,
//     newComment: string,
//     props: ISonaNonPoWebPartProps
//   ) => {
//     if (!amcDetailId || !newComment?.trim()) return;

//     const me = await sp.web.currentUser.select("Title", "Email")();
//     const who = `${me?.Title ?? "Unknown"}${me?.Email ? ` (${me.Email})` : ""}`;
//     const entry = `[${nowIST()}] ${who}: ${sanitizeRemark(newComment)}`;

//     const spCrudOps = await SPCRUDOPS();
//     const rows = await spCrudOps.getRootData(
//       "AMCDetail",
//       "ID,ApproverRemark",
//       "",
//       `ID eq ${amcDetailId}`,
//       { column: "ID", isAscending: true },
//       props
//     );
//     const current = (rows?.[0]?.ApproverRemark ?? "").toString().trim();
//     const finalLog = current ? `${current}\n${entry}` : entry;

//     if ((spCrudOps as any).updateItem) {
//       await (spCrudOps as any).updateItem("AMCDetail", amcDetailId, { ApproverRemark: finalLog }, props);
//     } else {
//       await sp.web.lists.getByTitle("AMCDetail").items.getById(amcDetailId).update({ ApproverRemark: finalLog });
//     }
//   };

//   // ---------- Helpers ----------
//   const toISODateOnly = (v: any): string => {
//     if (!v) return "";
//     if (v instanceof Date) return v.toISOString().split("T")[0];
//     const s = String(v).trim();
//     if (!s) return "";
//     if (s.includes("T") && /^\d{4}-\d{2}-\d{2}T/.test(s)) return s.split("T")[0];
//     if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
//     const d = new Date(s);
//     return isNaN(d.getTime()) ? "" : d.toISOString().split("T")[0];
//   };

//   const formatDateToDDMMYYYY = (dateInput: string | Date) => {
//     if (!dateInput) return "";
//     let iso = "";
//     if (dateInput instanceof Date) {
//       if (isNaN(dateInput.getTime())) return "";
//       iso = dateInput.toISOString().split("T")[0];
//     } else {
//       const str = dateInput.trim();
//       if (!str) return "";
//       iso = str.includes("T") ? str.split("T")[0] : str;
//       if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
//         const dt = new Date(str);
//         if (isNaN(dt.getTime())) return "";
//         iso = dt.toISOString().split("T")[0];
//       }
//     }
//     const [y, m, d] = iso.split("-");
//     return `${d}/${m}/${y}`;
//   };

//   /** truncate file name */
//   const shortName = (name: string, max = 35) => {
//     if (!name) return "";
//     return name.length > max ? `${name.slice(0, max)}...` : name;
//   };

//   // ---------- Attachments helpers (new structure: Renewal-{id} folder) ----------
//   async function resolveAttachLibraryRoot(): Promise<string> {
//     try {
//       const root = await amcWeb.lists.getByTitle(ATTACH_LIB_NAME).rootFolder.select("ServerRelativeUrl")();
//       return root.ServerRelativeUrl;
//     } catch {}

//     const knownPath = "/sites/AMC_UAT/AMCAttachments";
//     try {
//       const f = await amcWeb.getFolderByServerRelativePath(knownPath).select("ServerRelativeUrl")();
//       return f.ServerRelativeUrl;
//     } catch {}

//     const docLibs: any[] = await (amcWeb.lists
//       .filter("BaseTemplate eq 101")
//       .select("Title,RootFolder/ServerRelativeUrl")
//       .expand("RootFolder") as any)();

//     const match = docLibs.find(
//       (l: any) =>
//         (l.Title && l.Title.toLowerCase() === ATTACH_LIB_NAME.toLowerCase()) ||
//         (l.RootFolder?.ServerRelativeUrl &&
//           l.RootFolder.ServerRelativeUrl.toLowerCase().includes(ATTACH_LIB_NAME.toLowerCase()))
//     );

//     if (match?.RootFolder?.ServerRelativeUrl) return match.RootFolder.ServerRelativeUrl;
//     throw new Error("AMCAttachments library not found in AMC site.");
//   }

//   async function tryReadFolderFiles(folderServerRelativeUrl: string): Promise<IAttachmentItem[]> {
//     try {
//       const folder = amcWeb.getFolderByServerRelativePath(folderServerRelativeUrl);
//       const files = await folder.files.select("Name,ServerRelativeUrl,TimeCreated,Length")();
//       return (files ?? []) as any[];
//     } catch {
//       return [];
//     }
//   }

//   async function loadAttachmentsForRenewalRows(amcDetailId: number, renewalIds: number[]) {
//     if (!renewalIds.length) {
//       setAttachmentsByRenewalId({});
//       return;
//     }

//     setAttLoading(true);
//     try {
//       const libRoot = await resolveAttachLibraryRoot();
//       const masterFolder = `${libRoot}/AMCDetail-${amcDetailId}`;

//       const results = await Promise.allSettled(
//         renewalIds.map(async (rid) => {
//           const folderUrl = `${masterFolder}/Renewal-${rid}`;
//           const files = await tryReadFolderFiles(folderUrl);
//           return { rid, files };
//         })
//       );

//       const merged: Record<number, IAttachmentItem[]> = {};
//       results.forEach((r) => {
//         if (r.status === "fulfilled") merged[r.value.rid] = r.value.files;
//       });

//       setAttachmentsByRenewalId(merged);
//     } catch (e) {
//       console.error("[ViewAMC] loadAttachmentsForRenewalRows error:", e);
//       setAttachmentsByRenewalId({});
//     } finally {
//       setAttLoading(false);
//     }
//   }

//   // ---------- Main load ----------
//   useEffect(() => {
//     if (!Number.isNaN(amcId) && amcId > 0) loadData();
//     // eslint-disable-next-line react-hooks/exhaustive-deps
//   }, [amcId]);

//   const loadData = async () => {
//     try {
//       setLoading(true);
//       const spCrudOps = await SPCRUDOPS();

//       /* ===== MASTER (AMCDetail) ===== */
//       const master = await spCrudOps.getRootData(
//         "AMCDetail",
//         "ID,RequestedBy,RequestorEmail,RequestDate,Created,Description,RequestorLocation,Locationslt,AMCTypeslt,ContractTypeslt,ServiceTypeslt,VendorNameslt,Name,ApproverRemark",
//         "",
//         `ID eq ${amcId}`,
//         { column: "ID", isAscending: true },
//         props
//       );

//       if (!master || master.length === 0) {
//         alert("No record found for this ID");
//         setExistingDetails([]);
//         setAttachmentsByRenewalId({});
//         return;
//       }

//       const m = master[0];
//       const normalizedRequestDate = toISODateOnly(m.RequestDate) || toISODateOnly(m.Created);

//       const locationsText = String(m.Locationslt ?? ""); //  master union locations (for table)

//       setMasterInitial({
//         requestedBy: m.RequestedBy ?? "",
//         requestorEmail: m.RequestorEmail ?? "",
//         requestDate: normalizedRequestDate,
//         requestorLocation: m.RequestorLocation ?? "",
//         Description: m.Description ?? "",
//         locationsText: locationsText,

//         amcType: m.AMCTypeslt ?? "",
//         contractType: m.ContractTypeslt ?? "",
//         serviceType: m.ServiceTypeslt ?? "",
//         vendorName: m.VendorNameslt ?? "",
//         Name: m.Name ?? "",
//         ApproverRemark: m.ApproverRemark ?? "",
//       });

//       /* ===== DETAILS (AMCRenewal) EXISTING ===== */
//       let details = await spCrudOps.getRootData(
//         "AMCRenewal",
//         "ID,AMCStartDate,AMCEndDate,AMCRenewalDate,EstimatedAMCCost,Quantity,Description,RStatus,Responsibility/Title,RName,ApproverRemark",
//         "Responsibility",
//         `AMCRequestedIDId eq ${amcId}`,
//         { column: "ID", isAscending: true },
//         props
//       );

//       if (!details || details.length === 0) {
//         details = await spCrudOps.getRootData(
//           "AMCRenewal",
//           "ID,AMCStartDate,AMCEndDate,AMCRenewalDate,EstimatedAMCCost,Quantity,Description,RStatus,Responsibility/Title,AMCRequestedID/Id,RName,ApproverRemark",
//           "Responsibility,AMCRequestedID",
//           `AMCRequestedID/Id eq ${amcId}`,
//           { column: "ID", isAscending: true },
//           props
//         );
//       }

//       const mappedExisting: ExistingDetailRow[] = (details ?? []).map((d: any) => {
//         const parsed = parseRemarkLog((d?.ApproverRemark ?? "").toString(), "AMCRenewal");
//         const lastInitiator = [...parsed].reverse().find(r => (r.role || "").toLowerCase() === "initiator");
//         return {
//           id: d.ID,
//           amcStartDate: d.AMCStartDate?.split("T")?.[0] ?? "",
//           amcEndDate: d.AMCEndDate?.split("T")?.[0] ?? "",
//           amcRenewalDate: d.AMCRenewalDate?.split("T")?.[0] ?? "",
//           estimatedAMC: d.EstimatedAMCCost ?? "",
//           quantity: d.Quantity ?? "",
//           responsibilityName: d.Responsibility?.Title ?? "",
//           description: d.Description ?? "",
//           rStatus: d.RStatus ?? "",
//           rName: d.RName ?? "",
//           initialRemark: lastInitiator?.remark || ""
//         };
//       });

//       setExistingDetails(mappedExisting);

//       // ✅ attachments per renewal row folder
//       const renewalIds = mappedExisting.map(r => r.id);
//       await loadAttachmentsForRenewalRows(amcId, renewalIds);
//     } catch (e) {
//       console.error("[ViewAMC] Load error:", e);
//       alert("Failed to load record");
//     } finally {
//       setLoading(false);
//     }
//   };

//   // ---------- Approver Remarks Loader (Parent + Child) ----------
//   const loadApproverRemarks = async (row: ExistingDetailRow) => {
//     setRemarksLoading(true);
//     setParentRemarks([]);
//     setChildRemarks([]);
//     try {
//       // ===== CHILD (AMCRenewal): text first, fallback to raw
//       try {
//         const childAsText = await amcWeb.lists
//           .getByTitle("AMCRenewal")
//           .items.getById(row.id)
//           .fieldValuesAsText
//           .select("ApproverRemark")();
//         let cText = (childAsText?.ApproverRemark ?? "").toString();
//         if (!cText.trim()) {
//           const cItem = await amcWeb.lists.getByTitle("AMCRenewal").items.getById(row.id).select("ID,ApproverRemark")();
//           cText = (cItem?.ApproverRemark ?? "").toString();
//         }
//         setChildRemarks(parseRemarkLog(cText, "AMCRenewal"));
//       } catch {
//         setChildRemarks([]);
//       }

//       // ===== PARENT (AMCDetail): text first, fallback to raw
//       try {
//         const parentAsText = await amcWeb.lists
//           .getByTitle("AMCDetail")
//           .items.getById(amcId)
//           .fieldValuesAsText
//           .select("ApproverRemark")();
//         let pText = (parentAsText?.ApproverRemark ?? "").toString();
//         if (!pText.trim()) {
//           const pItem = await amcWeb.lists.getByTitle("AMCDetail").items.getById(amcId).select("ID,ApproverRemark")();
//           pText = (pItem?.ApproverRemark ?? "").toString();
//         }
//         setParentRemarks(parseRemarkLog(pText, "AMCDetail"));
//       } catch {
//         setParentRemarks([]);
//       }
//     } catch (e) {
//       console.error("[ViewAMC] Remarks load error:", e);
//       setChildRemarks([]);
//       setParentRemarks([]);
//     } finally {
//       setRemarksLoading(false);
//     }
//   };

//   // Merge & sort remarks (latest first)
//   const mergedRemarks = React.useMemo(() => {
//     const toTs = (created?: string) => {
//       const raw = (created || "").trim();
//       if (!raw) return 0;
//       const iso = /^\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2})?$/.test(raw)
//         ? raw.replace(" ", "T") + (raw.includes(":") ? ":00" : "T00:00:00")
//         : raw;
//       const d = new Date(iso);
//       return isNaN(d.getTime()) ? 0 : d.getTime();
//     };
//     const all = [...(parentRemarks || []), ...(childRemarks || [])];
//     return all.sort((a, b) => toTs(b.created) - toTs(a.created));
//   }, [parentRemarks, childRemarks]);

//   /** Render remarks with role-aware columns like EditDraft */
//   const renderRemarksTable = (rows: IRemark[]) => {
//     if (!rows?.length) return <div className="text-muted small">No remarks</div>;
//     const toDDMMYYYY = (s?: string) => (s ? formatDateToDDMMYYYY(s) : "-");
//     const roleLabel = (r: IRemark) => {
//       const raw = (r.role || "").trim();
//       if (!raw) return r.sourceList === "AMCDetail" ? "Initiator" : "Approver";
//       return raw;
//     };
//     const showStatus = (s?: string) => (s && s.trim() ? s.trim() : "-");
//     return (
//       <div className="table-responsive">
//         <table className="table table-sm table-bordered mb-0">
//           <thead className="table-light">
//             <tr>
//               <th style={{ width: 110 }}>Date</th>
//               <th style={{ width: 100 }}>Status</th>
//               <th style={{ width: 110 }}>Role</th>
//               <th style={{ width: 240 }}>By</th>
//               <th style={{ width: 260 }}>Approver Remark</th>
//               <th>Initiator Remark</th>
//             </tr>
//           </thead>
//           <tbody>
//             {rows.map((r) => {
//               const role = roleLabel(r);
//               const isApprover = role.toLowerCase() === "approver";
//               return (
//                 <tr key={`${r.sourceList ?? 'src'}-${r.id}`}>
//                   <td>{toDDMMYYYY(r.created)}</td>
//                   <td>{showStatus(r.status)}</td>
//                   <td>{role}</td>
//                   <td>{r.approver || "-"}</td>
//                   <td style={{ whiteSpace: "pre-wrap" }}>{isApprover ? (r.remark || "-") : "-"}</td>
//                   <td style={{ whiteSpace: "pre-wrap" }}>{!isApprover ? (r.remark || "-") : "-"}</td>
//                 </tr>
//               );
//             })}
//           </tbody>
//         </table>
//       </div>
//     );
//   };

//   // ---------- Attachments render per row ----------
//   const renderRowAttachments = (renewalId: number) => {
//     const files = attachmentsByRenewalId[renewalId] ?? [];
//     if (attLoading) return <span className="text-muted">Loading…</span>;
//     if (!files.length) return <span className="text-muted">No files</span>;

//     return (
//       <ul className="list-unstyled mb-0">
//         {files.map((f, i) => (
//           <li key={`${f.ServerRelativeUrl}-${i}`}>
//             <a
//               href={`${f.ServerRelativeUrl}?web=1`}
//               target="_blank"
//               rel="noreferrer"
//               title={f.Name}
//               style={{
//                 display: "inline-block",
//                 maxWidth: "240px",
//                 whiteSpace: "nowrap",
//                 overflow: "hidden",
//                 textOverflow: "ellipsis",
//                 verticalAlign: "middle"
//               }}
//             >
//               {shortName(f.Name)}
//             </a>
//           </li>
//         ))}
//       </ul>
//     );
//   };

//   // ---------- Action (eye) ----------
//   const openRowDetails = async (row: ExistingDetailRow) => {
//     setSelectedRow(row);
//     setShowDetailsModal(true);
//     await loadApproverRemarks(row);
//   };

//   if (loading) return <div>Loading...</div>;

//   return (
//     <Formik innerRef={formikRef} initialValues={masterInitial} enableReinitialize onSubmit={() => { }}>
//       <Form>
//         {/* ===== Header ===== */}
//         <div className="header">
//           <div className="left-banner">
//             <div className="logo-text">
//               <h2 style={{ color: "black", textAlign: "center" }}>View AMC</h2>
//             </div>
//           </div>
//         </div>

//         <div className="container p-3">
//           {/* ===== 1) AMC (Master, read-only) ===== */}
//           <div className="NAMC-section">
//             <div className="NAMC-section-header d-flex align-items-center justify-content-between">
//               <h4 className="mb-0">AMC</h4>
//               <button type="button" className="btn btn-dark btn-sm" onClick={() => history.push("/myAMC")}>
//                 Back
//               </button>
//             </div>

//             <div className="NAMC-section-body">
//               <div className="row g-3 mb-2">
//                 <div className="col-md-4">
//                   <label>Requested By</label>
//                   <Field name="requestedBy" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Email</label>
//                   <Field name="requestorEmail" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Request Date</label>
//                   <input
//                     type="text"
//                     value={formatDateToDDMMYYYY(masterInitial.requestDate)}
//                     disabled
//                     className="form-control"
//                   />
//                 </div>

//                 {/* ✅ Requestor Location */}
//                 <div className="col-md-4">
//                   <label>Requestor Location</label>
//                   <Field name="requestorLocation" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>AMC Type</label>
//                   <Field name="amcType" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Contract Type</label>
//                   <Field name="contractType" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Service Type</label>
//                   <Field name="serviceType" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Vendor</label>
//                   <Field name="vendorName" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Name</label>
//                   <Field name="Name" disabled className="form-control" />
//                 </div>

//                 <div className="col-md-4">
//                   <label>Initiator Remark</label>
//                   <Field
//                     as="textarea"
//                     rows={2}
//                     name="ApproverRemark"
//                     disabled
//                     className="form-control"
//                   />
//                 </div>

//                 {/* AMC Description */}
//                 <div className="col-md-4">
//                   <label>AMC Description</label>
//                   <Field as="textarea" rows={3} name="Description" disabled className="form-control" />
//                 </div>

//                 {/*  Locations field removed from master section */}
//                 {/*  AMC Period removed already */}
//               </div>
//             </div>
//           </div>

//           {/* ===== 2) Existing AMC Renewal Details ===== */}
//           <div className="NAMC-table-wrapper mt-4">
//             <div className="NAMC-table-header">
//               <h5>Existing AMC Renewal Details</h5>
//             </div>

//             <div className="NAMC-table-scroll">
//               <table className="table NAMC-table">
//                 <thead>
//                   <tr>
//                     <th>Start</th>
//                     <th>End</th>
//                     <th>Renewal</th>
//                     <th>Estimated AMC</th>
//                     <th>Qty</th>
//                     <th>Approver</th>
//                     {/* <th>Location</th>
//                     <th>Description</th> */}
//                     <th>RName</th>
//                     <th>Initiator Remark</th>
//                     <th>RStatus</th>
//                     <th>Attachments</th>
//                     {/* New Action column */}
//                     <th style={{ width: 60, textAlign: "center" }}>Action</th>
//                   </tr>
//                 </thead>

//                 <tbody>
//                   {existingDetails.length === 0 ? (
//                     <tr>
//                       <td colSpan={11} className="text-center text-muted">
//                         No renewal rows found
//                       </td>
//                     </tr>
//                   ) : (
//                     existingDetails.map((item) => (
//                       <tr key={item.id}>
//                         <td>{formatDateToDDMMYYYY(item.amcStartDate)}</td>
//                         <td>{formatDateToDDMMYYYY(item.amcEndDate)}</td>
//                         <td>{formatDateToDDMMYYYY(item.amcRenewalDate)}</td>
//                         <td>{item.estimatedAMC}</td>
//                         <td>{item.quantity}</td>
//                         <td>{item.responsibilityName}</td>

//                         {/*  Location could be shown as masterInitial.locationsText */}
//                         {/* <td>{masterInitial.locationsText}</td>
//                         <td>{item.description}</td> */}
//                         <td>{item.rName || "-"}</td>
//                         <td>{(item.initialRemark || "").trim() || "-"}</td>
//                         <td>{item.rStatus || "-"}</td>

//                         {/* ✅ Attachments per renewal row */}
//                         <td>{renderRowAttachments(item.id)}</td>

//                         {/* ✅ Action with eye icon */}
//                         <td style={{ textAlign: "center" }}>
//                           <button
//                             type="button"
//                             className="btn btn-link p-0"
//                             title="View details"
//                             onClick={() => openRowDetails(item)}
//                             aria-label="View details"
//                           >
//                             {/* Inline SVG Eye Icon */}
//                             <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="currentColor" viewBox="0 0 16 16">
//                               <path d="M16 8s-3-5.5-8-5.5S0 8 0 8s3 5.5 8 5.5S16 8 16 8" />
//                               <path d="M8 5.5a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5" />
//                             </svg>
//                           </button>
//                         </td>
//                       </tr>
//                     ))
//                   )}
//                 </tbody>
//               </table>
//             </div>
//           </div>

//           {/* Footer */}
//           <div className="mt-3">
//             <button type="button" className="btn btn-dark" onClick={() => history.push("/myAMC")}>
//               Close
//             </button>
//           </div>
//         </div>

//         {/* ================= Modal: Details + Approver Remarks ================= */}
//         <div
//           className={`modal fade ${showDetailsModal ? "show d-block" : ""}`}
//           tabIndex={-1}
//           role="dialog"
//           aria-modal="true"
//           style={{ background: showDetailsModal ? "rgba(0,0,0,0.5)" : "transparent" }}
//         >
//           <div className="modal-dialog modal-lg modal-dialog-scrollable">
//             <div className="modal-content">
//               <div className="modal-header">
//                 <h5 className="modal-title">Renewal Row Details</h5>
//                 <button
//                   type="button"
//                   className="btn-close"
//                   aria-label="Close"
//                   onClick={() => setShowDetailsModal(false)}
//                 />
//               </div>

//               <div className="modal-body">
//                 {!selectedRow ? (
//                   <div className="text-muted">No row selected</div>
//                 ) : (
//                   <>
//                     {/* --- Row summary mirroring table --- */}
//                     <div className="row g-3">
//                       <div className="col-md-4">
//                         <label className="form-label">Start</label>
//                         <input className="form-control" disabled value={formatDateToDDMMYYYY(selectedRow.amcStartDate)} />
//                       </div>
//                       <div className="col-md-4">
//                         <label className="form-label">End</label>
//                         <input className="form-control" disabled value={formatDateToDDMMYYYY(selectedRow.amcEndDate)} />
//                       </div>
//                       <div className="col-md-4">
//                         <label className="form-label">Renewal</label>
//                         <input className="form-control" disabled value={formatDateToDDMMYYYY(selectedRow.amcRenewalDate)} />
//                       </div>

//                       <div className="col-md-4">
//                         <label className="form-label">Estimated AMC</label>
//                         <input className="form-control" disabled value={selectedRow.estimatedAMC} />
//                       </div>
//                       <div className="col-md-4">
//                         <label className="form-label">Qty</label>
//                         <input className="form-control" disabled value={selectedRow.quantity} />
//                       </div>
//                       <div className="col-md-4">
//                         <label className="form-label">Approver</label>
//                         <input className="form-control" disabled value={selectedRow.responsibilityName || ""} />
//                       </div>

//                       <div className="col-md-12">
//                         <label className="form-label">Location</label>
//                         <input className="form-control" disabled value={masterInitial.locationsText} />
//                       </div>

//                       <div className="col-md-4">
//                         <label className="form-label">RName</label>
//                         <textarea className="form-control" disabled rows={2} value={selectedRow?.rName || ""} />
//                       </div>

//                       <div className="col-md-12">
//                         <label className="form-label">Description</label>
//                         <textarea className="form-control" disabled rows={2} value={selectedRow.description || ""} />
//                       </div>

//                       <div className="col-md-4">
//                         <label className="form-label">RStatus</label>
//                         <input className="form-control" disabled value={selectedRow.rStatus || "-"} />
//                       </div>

//                       <div className="col-md-12">
//                         <label className="form-label">Attachments</label>
//                         <div className="border rounded p-2">
//                           {renderRowAttachments(selectedRow.id)}
//                         </div>
//                       </div>
//                     </div>

//                     {/* --- Combined Comment History --- */}
//                     <hr className="my-3" />
//                     <div className="d-flex align-items-center justify-content-between">
//                       <h6 className="mb-0">Comment History</h6>
//                       {remarksLoading && <span className="text-muted small">Loading…</span>}
//                     </div>
//                     <div className="card mt-2">
//                       <div className="card-body p-2" style={{ maxHeight: 260, overflowY: "auto" }}>
//                         {renderRemarksTable(mergedRemarks)}
//                       </div>
//                     </div>
//                   </>
//                 )}
//               </div>

//               <div className="modal-footer">
//                 <button type="button" className="btn btn-secondary" onClick={() => setShowDetailsModal(false)}>
//                   Close
//                 </button>
//               </div>
//             </div>
//           </div>
//         </div>
//         {/* ================= /Modal ================= */}
//       </Form>
//     </Formik>
//   );
// };

// export default ViewNonPo;
