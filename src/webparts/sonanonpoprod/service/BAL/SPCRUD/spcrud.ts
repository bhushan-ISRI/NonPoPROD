// src/service/SPCRUD.ts

import type { ISonanonpoprodProps } from "../../../components/ISonanonpoprodProps";
import SPCRUDOPS from "../../DAL/spcrudops";

export interface ISPCRUD {
  getData(
    listName: string,
    columnsToRetrieve: string,
    columnsToExpand: string,
    filters: string,
    orderby: { column: string; isAscending: boolean },
    top: number,
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

export default async function USESPCRUD(
  props: ISonanonpoprodProps
): Promise<ISPCRUD> {
  const spCrudOps = await SPCRUDOPS(props.currentSPContext);

  return {
    getData: async (
      listName,
      columnsToRetrieve,
      columnsToExpand,
      filters,
      orderby,
      top,
      props
    ) => {
      return await spCrudOps.getTopData(
        listName,
        columnsToRetrieve,
        columnsToExpand,
        filters,
        orderby,
        top,
        props
      );
    },

    getRootData: async (
      listName,
      columnsToRetrieve,
      columnsToExpand,
      filters,
      orderby,
      props
    ) => {
      return await spCrudOps.getRootData(
        listName,
        columnsToRetrieve,
        columnsToExpand,
        filters,
        orderby,
        props
      );
    },

    insertData: async (listName, data, props) => {
      return await spCrudOps.insertData(listName, data, props);
    },

    updateData: async (listName, itemId, data, props) => {
      return await spCrudOps.updateData(listName, itemId, data, props);
    },

    deleteData: async (listName, itemId, props) => {
      return await spCrudOps.deleteData(listName, itemId, props);
    },

    getListInfo: async (listName, props) => {
      return await spCrudOps.getListInfo(listName, props);
    },

    getListData: async (listName, columnsToRetrieve, props) => {
      return await spCrudOps.getListData(listName, columnsToRetrieve, props);
    },

    getTopData: async (
      listName,
      columnsToRetrieve,
      columnsToExpand,
      filters,
      orderby,
      top,
      props
    ) => {
      return await spCrudOps.getTopData(
        listName,
        columnsToRetrieve,
        columnsToExpand,
        filters,
        orderby,
        top,
        props
      );
    },

    addAttchmentInList: async (
      attFiles,
      listName,
      itemId,
      fileName,
      props
    ) => {
      return await spCrudOps.addAttchmentInList(
        attFiles,
        listName,
        itemId,
        fileName,
        props
      );
    },
  };
}