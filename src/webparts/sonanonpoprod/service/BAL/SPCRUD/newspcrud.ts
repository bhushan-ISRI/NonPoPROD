import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/attachments";
//import NEWSPCRUDOPS from
import SPCRUDOPS from "../../DAL/newspcrudops";
//import SPCRUDOPS from "../../service/DAL/spcrudops";
//import  NEWInewspCrudOps  from "../../DAL/newnewspCrudOps";
import { ISonanonpoprodProps } from "../../../components/ISonanonpoprodProps";
 
export interface ISPCRUD {
    [x: string]: any;
    getData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: ISonanonpoprodProps): Promise<any>;
    getRootData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: ISonanonpoprodProps): Promise<any>;
    insertData(listName: string, data: any, props: ISonanonpoprodProps): Promise<any>;
    updateData(listName: string, itemId: number, data: any, props: ISonanonpoprodProps): Promise<any>;
    deleteData(listName: string, itemId: number, props: ISonanonpoprodProps): Promise<any>;
    getListInfo(listName: string, props: ISonanonpoprodProps): Promise<any>;
    getListData(listName: string, columnsToRetrieve: string, props: ISonanonpoprodProps): Promise<any>;
    createFolder(listName: string, folderName: string, props: ISonanonpoprodProps):Promise<any>;
  
    addAttchmentInList(attFiles: File, listName: string, itemId: number, fileName: string, props: ISonanonpoprodProps): Promise<any>;
}

export default async function USESPCRUD(this: any): Promise<ISPCRUD> {
   // const newspCrudOps = await newspCrudOps();
    //const spCrudOps = await NEWISPCRUDOPS();
    //const newspCrudOps=await  SPCRUDOPS(ISonanonpoprodProps)
    const props: ISonanonpoprodProps = {
        currentSPContext: this.context,
        context: undefined,
        isDarkTheme: false,
        environmentMessage: "",
        hasTeamsContext: false,
        userDisplayName: "",
        userEmail:"",

    }; // <-- type here is fine
const newspCrudOps = SPCRUDOPS(props);
  //  const newspCrudOps =SPCRUDOPS(props: ISonanonpoprodProps); //
    return {
        getData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: ISonanonpoprodProps) => {
            return await newspCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby);
        },
        getRootData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: ISonanonpoprodProps) => {
            return await newspCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby);
        },
        insertData: async (listName: string, data: any, props: ISonanonpoprodProps) => {
            return await newspCrudOps.insertData(listName, data);
        },
        updateData: async (listName: string, itemId: number, data: any, props: ISonanonpoprodProps) => {
            return await newspCrudOps.updateData(listName, itemId, data);
        },
        deleteData: async (listName: string, itemId: number, props: ISonanonpoprodProps) => {
            return await newspCrudOps.deleteData(listName, itemId);
        },
        getListInfo: async (listName: string, props: ISonanonpoprodProps) => {
            return await newspCrudOps.getListInfo(listName);
        },
        getListData: async (listName: string, columnsToRetrieve: string, props: ISonanonpoprodProps) => {
            return await newspCrudOps.getListData(listName, columnsToRetrieve);
        },
      
        createFolder: async (listName: string, folderName: string, props: ISonanonpoprodProps) => {
            return await newspCrudOps.createFolder(listName, folderName);
        },
        // uploadFile: async (folderServerRelativeUrl: string, file: File, props: ISonanonpoprodProps) => {
        //     return await newspCrudOps.uploadFile(folderServerRelativeUrl, file, props);
        // },
        // deleteFile: async (fileServerRelativeUrl: string, props: ISonanonpoprodProps) => {
        //     return await newspCrudOps.deleteFile(fileServerRelativeUrl, props);
        // },
       
       
        addAttchmentInList: async (attFiles: File, listName: string, itemId: number, fileName: string, props: ISonanonpoprodProps) => {
            return await newspCrudOps.addAttchmentInList(attFiles, listName, itemId, fileName);
        }
    };
}