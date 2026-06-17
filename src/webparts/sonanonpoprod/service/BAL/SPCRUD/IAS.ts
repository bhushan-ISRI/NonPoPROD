import { IIAS } from "../../INTERFACE/IIAS";
import { ISonanonpoprodProps } from "../../../components/ISonanonpoprodProps";
import SPCRUDOPS from "../../DAL/spcrudops";
// import { ISonaNonPoWebPartProps } from "../../../SonaNonPoWebPart";
// import type { ISonanonpoprodProps } from '../../../components/ISonanonpoprodProps';
export interface IIITRequestsOps {
    getIIASData(props: IIAS): Promise<IIAS>;
    getIASDatafilter(props: IIAS): Promise<IIAS>;
}

export default async function IASRequestsOps(props: ISonanonpoprodProps) {

const spCrudOps = await SPCRUDOPS(props.currentSPContext);


    // const getIIASData = async (strFilter: string, sorting: any,props: IItProps): Promise<IIAS[]> => {
    const getIIASData = async (sorting: any, props: ISonanonpoprodProps, filter:string): Promise<IIAS[]> => {
        return await (await spCrudOps).getData("IAS_List"
            , "*,ApprovalNoteNo,SAPNo,MovementReason,Status,Stage,NA/ID,NA/Title,NA/EMail,DA/ID,DA/Title,DA/EMail,Details,Summary,WF,HeaderName,MovementType,Department,CostCenter,UniquePartImpact,GrossValue,NetValue,LastAction,CostCenterDescription,Author/Title,AttachmentFiles,NextApproverEmpID"
            , "NA,DA,Author,AttachmentFiles"
            , filter
            , sorting
            , props).then(results => {
                let brr: Array<IIAS> = new Array<IIAS>();
                results.map((item: {
                    ID: number;
                    ApprovalNoteNo?: number;
                    SAPNo:any; 
                    MovementReason:any;  
                    Status:any;  
                    Stage:any; 
                    NA:any; 
                    NATitle:any;
                    NAEmail:any;
                    DA:any;
                    Details:any;
                    Summary:any;
                    WF:any;
                    HeaderName:any;
                    MovementType:any;
                    Department:any;
                    CostCenter:any;
                    UniquePartImpact:any;
                    GrossValue:any;
                    NetValue:any;
                    LastAction:any;
                    CostCenterDescription:any; 
                    Author:any;
                    Title:any; 
                    Created:any;  
                    NextApproverEmpID:any; 
                    DelegateApproverEmpID:any;
                    InitiatorEmpId:any;
                    InventoryAttachment:any;
                }) => {
                    brr.push({
                    ID:item.ID,
                    ApprovalNoteNo:item?.ApprovalNoteNo?? null,
                    SAPNo:item?.SAPNo?? null,
                    MovementReason:item?.MovementReason?? null,
                    Status:item?.Status?? null,
                    Stage:item?.Stage?? null,
                    NAID:item?.NA?.ID?? null,  
                    NATitle:item?.NA?.Title?? null,
                    NAEmail:item?.NA?.EMail?? null,
                    DAID:item?.DA?.ID?? null,  
                    DATitle:item?.DA?.Title?? null,
                    DAEmail:item?.DA?.EMail?? null,
                    Details:item?.Details?? null,
                    Summary:item?.Summary?? null,
                    WF:item?.WF?? null,
                    HeaderName:item?.HeaderName?? null,
                    MovementType:item?.MovementType?? null,
                    Department:item?.Department?? null,
                    CostCenter:item?.CostCenter?? null,
                    UniquePartImpact:item?.UniquePartImpact?? null,
                    GrossValue:item?.GrossValue?? null,
                    NetValue:item?.NetValue?? null,
                    LastAction:item?.LastAction?? null,
                    CostCenterDescription:item?.CostCenterDescription?? null,
                    AttachmentFiles:null,
                    Title:item?.Title,
                    Author:item?.Author.Title??null,
                    Created:item?.Created??null,
                    NextApproverEmpID:item?.NextApproverEmpID??null,
                    InitiatorEmpId:item?.InitiatorEmpId??null,
                    DelegateApproverEmpID:item?.DelegateApproverEmpID??null,
                    InventoryAttachment:item?.InventoryAttachment??null
                    });
                });
                return brr;
                }
            );
    };

    const getIASDatafilter = async (ArtId: string | number, props: ISonanonpoprodProps): Promise<IIAS[]> => {
        return await (await spCrudOps).getData("IAS_List"
            , "*,ApprovalNoteNo,SAPNo,MovementReason,Status,Stage,NA/ID,NA/Title,NA/EMail,DA/ID,DA/Title,DA/EMail,Details,Summary,WF,HeaderName,MovementType,Department,CostCenter,UniquePartImpact,GrossValue,NetValue,LastAction,CostCenterDescription,Author/Title,AttachmentFiles,NextApproverEmpID"
            , "NA,DA,Author,AttachmentFiles"
            , "Id eq '" + ArtId + "'"
            // , sorting,
            , { column: 'Order0', isAscending: true },
            props).then(results => {
                let brr: Array<IIAS> = new Array<IIAS>();
                results.map((item: {
                    ID: number;
                    ApprovalNoteNo?: number;
                    SAPNo:any; 
                    MovementReason:any;  
                    Status:any;  
                    Stage:any; 
                    NA:any; 
                    NATitle:any;
                    NAEmail:any;
                    DA:any;
                    Details:any;
                    Summary:any;
                    WF:any;
                    HeaderName:any;
                    MovementType:any;
                    Department:any;
                    CostCenter:any;
                    UniquePartImpact:any;
                    GrossValue:any;
                    NetValue:any;
                    LastAction:any;
                    CostCenterDescription:any; 
                    AttachmentFiles:any; 
                    Title:any;  
                    Author:any;   
                    Created:any; 
                    NextApproverEmpID:any; 
                    DelegateApproverEmpID:any;
                    InitiatorEmpId:any;  
                    InventoryAttachment:any;                                    
                }) => {
                    brr.push({
                    ID:item.ID,
                    ApprovalNoteNo:item?.ApprovalNoteNo?? null,
                    SAPNo:item?.SAPNo?? null,
                    MovementReason:item?.MovementReason?? null,
                    Status:item?.Status?? null,
                    Stage:item?.Stage?? null,
                    NAID:item?.NA?.ID?? null,  
                    NATitle:item?.NA?.Title?? null,
                    NAEmail:item?.NA?.EMail?? null,
                    DAID:item?.DA?.ID?? null,  
                    DATitle:item?.DA?.Title?? null,
                    DAEmail:item?.DA?.EMail?? null,
                    Details:item?.Details?? null,
                    Summary:item?.Summary?? null,
                    WF:item?.WF?? null,
                    HeaderName:item?.HeaderName?? null,
                    MovementType:item?.MovementType?? null,
                    Department:item?.Department?? null,
                    CostCenter:item?.CostCenter?? null,
                    UniquePartImpact:item?.UniquePartImpact?? null,
                    GrossValue:item?.GrossValue?? null,
                    NetValue:item?.NetValue?? null,
                    LastAction:item?.LastAction?? null,
                    CostCenterDescription:item?.CostCenterDescription?? null,
                    AttachmentFiles:item?.AttachmentFiles?? null,
                    Title:item?.Title,
                    Author:item?.Author.Title??null,
                    Created:item?.Created??null,
                    NextApproverEmpID:item?.NextApproverEmpID??null,
                    InitiatorEmpId:item?.InitiatorEmpId??null,
                    DelegateApproverEmpID:item?.DelegateApproverEmpID??null,
                    InventoryAttachment:item?.InventoryAttachment??null
                    });
                });
                return brr;
            }
            );
    };

    return {
        getIIASData, getIASDatafilter
    };
}