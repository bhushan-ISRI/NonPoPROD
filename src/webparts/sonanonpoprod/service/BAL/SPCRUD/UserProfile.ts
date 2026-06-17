// UserProfile.ts (PnPjs v4-compatible)

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all"; // SPFx binding helper
import "@pnp/sp/profiles";                  // attach profiles APIs
import { IUserProfile, IUserProfileProperty } from "../../INTERFACE/IUserProfile";
// import { ISonaNonPoWebPartProps } from "../../../SonaNonPoWebPart";
import { ISonanonpoprodProps } from "../../../components/ISonanonpoprodProps";
// --- simple singleton so we bind context once and reuse ---
let _sp: SPFI | null = null;
function getSP(context: any): SPFI {
  if (!_sp) {
    _sp = spfi().using(SPFx(context)); // bind SPFx web part context
  }
  return _sp;
}

const UserProfileOps = () => {
  const getLoggUserProfile = async (props: ISonanonpoprodProps): Promise<IUserProfile> => {
    // get a bound SP instance
    const sp = getSP(props.currentSPContext);

    // v4: profiles available after importing "@pnp/sp/profiles"
    const myProps: any = await sp.profiles.myProperties();

    const propsArr: IUserProfileProperty[] =
      (myProps?.UserProfileProperties ?? []).map((p: any) => ({
        Key: p?.Key,
        Value: p?.Value,
      }));

    const getVal = (key: string) => propsArr.find((p) => p.Key === key)?.Value ?? "";

    const loc =
      getVal("SPS-Location") ||
      getVal("Office") ||
      getVal("SPS-OfficeLocation") ||
      getVal("OfficeNumber") ||
      getVal("PhysicalDeliveryOfficeName") ||
      "";

    return {
      accountName: myProps?.AccountName,
      displayName: myProps?.DisplayName,
      email: myProps?.Email,
      Location: loc,
      UserProfileProperties: propsArr,
      get: (key: string) => getVal(key),
    };
  };

  return { getLoggUserProfile };
};

export default UserProfileOps;