// SPCRUDOPS.ts
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/attachments";
import { ISonanonpoprodProps } from "../../components/ISonanonpoprodProps";

export interface ISPCRUDOPS {
    getData(listName: string, columnsToRetrieve?: string, columnsToExpand?: string, filters?: string, orderby?: { column: string, isAscending: boolean }): Promise<any>;
    getRootData(listName: string, columnsToRetrieve?: string, columnsToExpand?: string, filters?: string, orderby?: { column: string, isAscending: boolean }): Promise<any>;
    insertData(listName: string, data: any): Promise<any>;
    updateData(listName: string, itemId: number, data: any): Promise<any>;
    deleteData(listName: string, itemId: number): Promise<any>;
    getListInfo(listName: string): Promise<any>;
    getListData(listName: string, columnsToRetrieve?: string): Promise<any>;
    createFolder(listName: string, folderName: string): Promise<any>;
   // uploadFile(folderServerRelativeUrl: string, file: File): Promise<any>;
    //deleteFile(fileServerRelativeUrl: string): Promise<any>;
    addAttchmentInList(data: File, listName: string, itemId: number, fileName: string): Promise<any>;
}

class SPCRUDOPSImpl implements ISPCRUDOPS {
    private sp;

    constructor(props: ISonanonpoprodProps) {
        if (!props?.currentSPContext) throw new Error("SharePoint context is not available");
        this.sp = spfi().using(SPFx(props.currentSPContext));
    }

    async getData(listName: string, columnsToRetrieve?: string, columnsToExpand?: string, filters?: string, orderby?: { column: string, isAscending: boolean }): Promise<any> {
        let items = this.sp.web.lists.getByTitle(listName).items;
        if (columnsToRetrieve) items = items.select(columnsToRetrieve);
        if (columnsToExpand) items = items.expand(columnsToExpand);
        if (filters) items = items.filter(filters);
        if (orderby) items = items.orderBy(orderby.column, orderby.isAscending);
        return await items(); // v4 replacement for getAll()
    }

    async getRootData(listName: string, columnsToRetrieve?: string, columnsToExpand?: string, filters?: string, orderby?: { column: string, isAscending: boolean }): Promise<any> {
        return this.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby);
    }

    async insertData(listName: string, data: any): Promise<any> {
        return await this.sp.web.lists.getByTitle(listName).items.add(data);
    }

    async updateData(listName: string, itemId: number, data: any): Promise<any> {
        return await this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
    }

    async deleteData(listName: string, itemId: number): Promise<any> {
        return await this.sp.web.lists.getByTitle(listName).items.getById(itemId).recycle();
    }

    async getListInfo(listName: string): Promise<any> {
        return await this.sp.web.lists.getByTitle(listName);
    }

    async getListData(listName: string, columnsToRetrieve?: string): Promise<any> {
        let items = this.sp.web.lists.getByTitle(listName).items;
        if (columnsToRetrieve) items = items.select(columnsToRetrieve);
        return await items();
    }

    async createFolder(listName: string, folderName: string): Promise<any> {
        return await this.sp.web.lists.getByTitle(listName).rootFolder.folders.addUsingPath(folderName);
    }

    // async uploadFile(folderServerRelativeUrl: string, file: File): Promise<any> {
    //     return await this.sp.web.getFolderByServerRelativeUrl(folderServerRelativeUrl).files.add(file.name, file, true);
    // }

    // async deleteFile(fileServerRelativeUrl: string): Promise<any> {
    //     return await this.sp.web.getFileByServerRelativeUrl(fileServerRelativeUrl).recycle();
    // }

    async addAttchmentInList(data: File, listName: string, itemId: number, fileName: string): Promise<any> {
        return await this.sp.web.lists.getByTitle(listName).items.getById(itemId).attachmentFiles.add(fileName, data);
    }
}

// ✅ Factory function
export default function SPCRUDOPS(props: ISonanonpoprodProps): ISPCRUDOPS {
    return new SPCRUDOPSImpl(props);
}