// Utilities.ts (PnPjs v4 compatible)

import { IDeviationDevProps } from "../../INTERFACE/IDeviationDevProps";

import { IEmailProperties } from "@pnp/sp/sputilities"; // <- email utility types
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";             // <- SPFx binding
import "@pnp/sp/sputilities";                           // <- attach utility APIs (sendEmail)

export interface IMonth {
  Id: number;
  Title: string;
  ShortMonth: string;
  NarrowMonth: string;
}

export interface TypedHash<T> {
  [key: string]: T;
}

export interface IUtilities {
  filterData(data: any[], filterValue: string, filterColumns: string[]): Promise<any[]>;
  MonthColl(): Array<IMonth>;
  sendEmail(
    to: string[],
    cc: string[],
    bcc: string[],
    subject: string,
    body: string,
    AdditionalHeaders: TypedHash<string>,
    props: IDeviationDevProps,
    from?: string
  ): Promise<void>;
}

// --- create & reuse a single bound SP instance per runtime ---
let _sp: SPFI | null = null;
function getSP(context: any): SPFI {
  if (!_sp) {
    _sp = spfi().using(SPFx(context)); // bind to SPFx web part context once
  }
  return _sp;
}

export default function Utilities(): Promise<IUtilities> {
  return new Promise((resolve) => {
    const utilities: IUtilities = {
      filterData: async (data: any[], filterValue: string, filterColumns: string[]): Promise<any[]> => {
        if (!filterValue) return data;

        const searchValue = filterValue.toLowerCase();
        return data.filter((item) =>
          filterColumns.some((column) => {
            const value = item[column];
            if (value === null || value === undefined) return false;
            return value.toString().toLowerCase().includes(searchValue);
          })
        );
      },

      MonthColl: () => {
        const months: Array<IMonth> = [];
        const now = new Date();

        for (let m = 0; m <= 11; m++) {
          const dt = new Date(now.getFullYear(), m, 1);
          months.push({
            Id: m + 1,
            Title: dt.toLocaleString("en-us", { month: "long" }),
            ShortMonth: dt.toLocaleString("en-us", { month: "short" }),
            NarrowMonth: dt.toLocaleString("en-us", { month: "narrow" }),
          });
        }
        return months;
      },

      sendEmail: async (
        to: string[],
        cc: string[],
        bcc: string[],
        subject: string,
        body: string,
        additionalHeaders: TypedHash<string>,
        props: IDeviationDevProps,
        from?: string
      ): Promise<void> => {
        // Get a context-bound SP instance
        const sp = getSP(props.currentSPContext);

        const emailProps: IEmailProperties = {
          To: to,
          CC: cc,
          BCC: bcc,
          Subject: subject,
          Body: body,
        };

        if (from && from.trim()) {
          emailProps.From = from;
        }
        if (additionalHeaders) {
          emailProps.AdditionalHeaders = additionalHeaders;
        }

        await sp.utility.sendEmail(emailProps);
      },
    };

    resolve(utilities);
  });
}