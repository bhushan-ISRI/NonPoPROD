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

import Left from "../../components/Pages/Image/LeftArrow.png";
import Right from "../../components/Pages/Image/RightArrow.png";
import Edit from "../../components/Pages/Image/Pencil.png"
import View from "../../components/Pages/Image/Eye.png";

import { Link } from "react-router-dom";
import { useNavigate } from 'react-router-dom';

interface IRouteParams extends Record<string, string | undefined> {
    id?: string;
}

export const InitiatorApproved = (props: ISonanonpoprodProps) => {
    const navigate = useNavigate();
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [searchTerm, setSearchTerm] = useState("");
    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);
    const [ApprovedData, setApprovedData] = React.useState<any[]>([]);
    const [ItemID, setItemID] = useState<any>();
    const [requestNo, setrequestNo] = useState<any>();// useState<string | undefined>(undefined);                     

    useEffect(() => {
        loadApproveItems();
    }, []);
    const formatDate = (dateString: string) => {
        if (!dateString) return "";
        const date = new Date(dateString);

        return date
            .toLocaleDateString("en-GB", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
            })
            .replace(/\//g, "-"); // dd-MM-yyyy
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
    const handlePageChange = (page: number) => {
        if (page >= 1 && page <= totalPages) {
            setCurrentPage(page);
        }
    };



    const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

    const paginatedData = sortedData.slice(
        (currentPage - 1) * itemsPerPage,
        currentPage * itemsPerPage
    );
    useEffect(() => {
        if (!searchTerm) {
            setFilteredData(ApprovedData);
            setCurrentPage(1);
            return;
        }

        const lowerSearch = searchTerm.toLowerCase();


        const filtered = ApprovedData.filter((item) =>
            Object.values({
                Id: item.Id,
                Title: item.Title,
                RequestDate: formatDate(item.RequestDate),
                EmployeeName: item.EmployeeName,
                VendorName: item.VendorName,
                VendorCode: item.VendorCode?.VendorCode,
                InvoiceNumber: item.InvoiceNumber,
                Amount: item.Amount,
                Status: item.Status
            })
                .join(" ")
                .toLowerCase()
                .includes(lowerSearch)
        );

        setFilteredData(filtered);
        setCurrentPage(1);
    }, [searchTerm, ApprovedData]);

    const loadApproveItems = async () => {
        const spCrudOps = await SPCRUDOPS(props);
        const sp = spfi().using(SPFx(props.currentSPContext));
        //  const currentUser = await sp.web.currentUser();
        const currentUser = await sp.web.currentUser();
        const parentItems = await spCrudOps.getData(
            "NonPO",
            "ID,Title,RequestDate,EmployeeName,ApprovalMatrix,WorkflowHistory,Email,ContactNo,Division,Department,Location,RM,HOD,EmployeeStatus,ApprovalMatrix,VendorName,InvoiceNumber,InvoiceDate,Basicamount,GST,OtherCharges,Totalamount,WorkflowHistory,Status,PaymentRemarks,ApprovalRemarks,VendorCode/VendorCode,VendorCode/Id,VendorCode/Title,NatureofExpense/Title,CurrentApprover/ID,CurrentApprover/EMail,CurrentApprover/Title",
            "VendorCode,NatureofExpense,CurrentApprover",
            `Status eq 'Approved' and Email eq '${props.userEmail}'`,
            { column: "ID", isAscending: true }

        );
        console.log(parentItems);

        console.log(parentItems)
        setApprovedData(parentItems)
        setFilteredData(parentItems);

    }

    // ✅ UI must return something
    return (
        <div className='MainUplodForm' style={{ margin: "0px" }}>
            <div className='row'>
                <div className='col-md-12'>
                    <div className='Main-Boxpoup'>
                        <div className="bordered">
                            <img src={logo} />
                            <h1>My Approved Requests </h1>
                        </div>
                        <div className='col-md-12 px-2 d-flex justify-content-between align-items-center flex-wrap' style={{ margin: "5px" }}>
                            <div>
                                <input type="text" placeholder="Search..."
                                    className="form-control" style={{ width: "250px" }}
                                    value={searchTerm}
                                    onChange={(e) => setSearchTerm(e.target.value)}
                                />
                            </div>
                        </div>
                        <div className="mx-2" style={{ overflowX: "auto" }}>
                            <table className="custom-table" style={{ width: '100%', borderCollapse: 'collapse' }}>
                                <thead>
                                    <tr>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>View</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>NonPo Request No</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Request Date</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Requestor Name</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Vendor Code</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Vendor Name</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Invoice No</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Amount</th>
                                        <th style={{ padding: '8px', textAlign: 'left' }}>Status</th>


                                    </tr>
                                </thead>
                                <tbody>
                                    {paginatedData.map((item, index) => (
                                        <tr key={index} className="border-t">
                                            <td className="px-4 py-2">
                                                <Link to={`/ViewRequest/${item.Id}`}>
                                                    <img src={View} width={15} />
                                                </Link>
                                            </td>
                                            <td className="px-4 py-2">{item.Title}</td>
                                            <td className="px-4 py-2">{item.RequestDate}</td>
                                            <td className="px-4 py-2">{item.EmployeeName}</td>
                                            <td className="px-4 py-2">{item.VendorName}</td>
                                            <td className="px-4 py-2">{item.VendorCode.VendorCode}</td>
                                            <td className="px-4 py-2">{item.InvoiceNumber}</td>
                                            <td className="px-4 py-2">{item.Amount}</td>
                                            <td className="px-4 py-2">{item.Status}</td>

                                        </tr>
                                    ))}
                                </tbody>

                            </table>
                        </div>
                        {/* Pagination */}
                        <div className="flex justify-center mt-6 overflow-x-auto">
                            <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow" style={{ textAlign: "end" }}>
                                {/* Previous Button */}
                                <button
                                    onClick={() => handlePageChange(currentPage - 1)}
                                    disabled={currentPage === 1}
                                    style={{
                                        backgroundColor: "#fff",
                                        border: "1px solid #000 !important",
                                        marginRight: "5px",
                                        opacity: currentPage === 1 ? 0.5 : 1,
                                    }}
                                    className="px-3 py-1 border rounded"
                                >
                                    <img src={Left} alt="" width={15} />
                                </button>
                                {/* Main Page Numbers */}
                                {Array.from({ length: totalPages }, (_, i) => i + 1)
                                    .filter((page) => Math.abs(page - currentPage) <= 2)
                                    .map((page) => (
                                        <button
                                            key={page}
                                            onClick={() => handlePageChange(page)}
                                            style={{
                                                backgroundColor: currentPage === page ? "#3c3e45" : "#fff",
                                                color: currentPage === page ? "#fff" : "#000",
                                                fontWeight: currentPage === page ? "bold" : "normal",
                                                margin: currentPage === page ? "5px" : "5px",
                                            }}
                                            className="px-3 py-1 border rounded"
                                        >
                                            {page}
                                        </button>
                                    ))}

                                {/* Next Button */}
                                <button
                                    onClick={() => handlePageChange(currentPage + 1)}
                                    disabled={currentPage === totalPages}
                                    style={{
                                        backgroundColor: "#fff",
                                        border: "1px solid #000 !important",
                                        marginLeft: "5px",
                                        opacity: currentPage === totalPages ? 0.5 : 1,
                                    }}
                                    className="px-3 py-1 border rounded"
                                >
                                    <img src={Right} alt="" width={15} />
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

    );
};