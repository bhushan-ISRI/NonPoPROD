// src/service/DAL/spcrudops.ts

import type { ISonanonpoprodProps } from "../../components/ISonanonpoprodProps";

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx"; 
// import { SPFx } from "@pnp/sp/presets/all"; 



// PnP feature imports
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

export interface ISPCRUDOPS {
  getData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    props: ISonanonpoprodProps
  ): Promise<any>;

  getRootData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    props: ISonanonpoprodProps
  ): Promise<any>;

  insertData(listName: string, data: any, props: ISonanonpoprodProps): Promise<any>;

  updateData(
    listName: string,
    itemId: number,
    data: any,
    props: ISonanonpoprodProps
  ): Promise<any>;

  deleteData(
    listName: string,
    itemId: number,
    props: ISonanonpoprodProps
  ): Promise<any>;

  getListInfo(listName: string, props: ISonanonpoprodProps): Promise<any>;

  getListData(
    listName: string,
    columnsToRetrieve: string,
    props: ISonanonpoprodProps
  ): Promise<any>;

  getTopData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    top: number,
    props: ISonanonpoprodProps
  ): Promise<any>;

  addAttchmentInList(
    attFiles: File,
    listName: string,
    itemId: number,
    fileName: string,
    props: ISonanonpoprodProps
  ): Promise<any>;
}

// Singleton SPFI with SPFx behavior
let _sp: SPFI | null = null;

const getSP = (ctx: any): SPFI => {
  if (!_sp) {
    _sp = spfi().using(SPFx(ctx)); //  registers fetch behavior
  }
  return _sp;
};

class SPCRUDOPSImpl implements ISPCRUDOPS {
  private applyQuery(
    items: any,
    columnsToRetrieve?: string,
    columnsToExpand?: string,
    filters?: string,
    orderby?: { column: string; isAscending: boolean }
  ) {
    if (columnsToRetrieve) items = items.select(columnsToRetrieve);
    if (columnsToExpand) items = items.expand(columnsToExpand);
    if (filters) items = items.filter(filters);
    if (orderby) items = items.orderBy(orderby.column, orderby.isAscending);
    return items;
  }

  async getData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    props: ISonanonpoprodProps
  ): Promise<any> {
    const sp = getSP(props.currentSPContext);
    let items = sp.web.lists.getByTitle(listName).items;
    items = this.applyQuery(items, columnsToRetrieve, columnsToExpand, filters, orderby);
    return await items();
  }

  async getRootData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    props: ISonanonpoprodProps
  ): Promise<any> {
    // Currently same as getData (root web). Change to tenant root if you need site collection root.
    const sp = getSP(props.currentSPContext);
    let items = sp.web.lists.getByTitle(listName).items;
    items = this.applyQuery(items, columnsToRetrieve, columnsToExpand, filters, orderby);
    return await items();
  }

  async insertData(listName: string, data: any, props: ISonanonpoprodProps): Promise<any> {
    const sp = getSP(props.currentSPContext);
    return await sp.web.lists.getByTitle(listName).items.add(data);
  }

  async updateData(
    listName: string,
    itemId: number,
    data: any,
    props: ISonanonpoprodProps
  ): Promise<any> {
    const sp = getSP(props.currentSPContext);
    return await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
  }

  async deleteData(
    listName: string,
    itemId: number,
    props: ISonanonpoprodProps
  ): Promise<any> {
    const sp = getSP(props.currentSPContext);
    return await sp.web.lists.getByTitle(listName).items.getById(itemId).recycle();
  }

  async getListInfo(listName: string, props: ISonanonpoprodProps): Promise<any> {
    const sp = getSP(props.currentSPContext);
    return await sp.web.lists.getByTitle(listName)();
  }

  async getListData(
    listName: string,
    columnsToRetrieve: string,
    props: ISonanonpoprodProps
  ): Promise<any> {
    const sp = getSP(props.currentSPContext);
    let items = sp.web.lists.getByTitle(listName).items;
    if (columnsToRetrieve) items = items.select(columnsToRetrieve);
    return await items();
  }

  async getTopData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    top: number,
    props: ISonanonpoprodProps
  ): Promise<any> {
    const sp = getSP(props.currentSPContext);
    let items = sp.web.lists.getByTitle(listName).items;
    items = this.applyQuery(items, columnsToRetrieve, columnsToExpand, filters, orderby);
    if (top) items = items.top(top);
    return await items();
  }

  async addAttchmentInList(
    attFiles: File,
    listName: string,
    itemId: number,
    fileName: string,
    props: ISonanonpoprodProps
  ): Promise<any> {
    const sp = getSP(props.currentSPContext);
    return await sp.web.lists
      .getByTitle(listName)
      .items.getById(itemId)
      .attachmentFiles.add(fileName, attFiles);
  }
}

export default function SPCRUDOPS(ctx: any): Promise<ISPCRUDOPS> {
  // Ensure SP is initialized once per app lifecycle
  getSP(ctx);
  return Promise.resolve(new SPCRUDOPSImpl());
}