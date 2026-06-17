import * as React from "react";
import { useLocation } from "react-router-dom";
import "../Pages/CSS/Sidebar.scss";
import { ISonanonpoprodProps } from "../ISonanonpoprodProps";
import "@fortawesome/fontawesome-free/css/all.min.css";
import logo from "../../assets/SonaPNGLogo.png";

// ✅ PnPJS v3 Imports
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

type TabItem = {
  id: number;
  title: string;
  seq: number;
  url: string;
};

const Sidebar = (props: ISonanonpoprodProps) => {
  const location = useLocation();

  const [tabs, setTabs] = React.useState<TabItem[]>([]);
  const [selectedTabUrl, setSelectedTabUrl] = React.useState<string>("");

  const username = props.userDisplayName;

  // ✅ UseRef for SP instance (best practice)
  const spRef = React.useRef<SPFI | null>(null);

  /** ---------------------------------------------------------
   * Initialize PnPJS
   * --------------------------------------------------------- */
  React.useEffect(() => {
    spRef.current = spfi().using(SPFx(props.context));
  }, [props.context]);

  /** ---------------------------------------------------------
   * Load tabs with access control
   * --------------------------------------------------------- */
  const loadTabsWithAccess = async () => {
    try {
      if (!spRef.current) return;

      // ✅ Get current user's SharePoint groups
      const userGroups = await spRef.current.web.currentUser.groups();

      const userGroupIds = userGroups.map((g: { Id: number }) => g.Id);

      // ✅ Get all tabs from SharePoint list
      const items: any[] = await spRef.current.web.lists
        .getByTitle("NONPOTabbing")
        .items.select(
          "Id",
          "Title",
          "SeqNo",
          "PageUrl",
          "TabingViewGroup/Id",
          "TabingViewGroup/Title"
        )
        .expand("TabingViewGroup")();

      // ✅ Filter tabs based on group access
      const allowedTabs: TabItem[] = items
        .filter((tab) => {
          const groups: { Id: number }[] = tab.TabingViewGroup || [];

          // ✅ Public tab (no group assigned)
          if (groups.length === 0) return true;

          // ✅ User belongs to group
          return groups.some((g) => userGroupIds.includes(g.Id));
        })
        .map((tab) => ({
          id: tab.Id,
          title: tab.Title,
          seq: tab.SeqNo || 999,
          url: tab.PageUrl ? tab.PageUrl.replace(/\s+/g, "") : "",
        }))
        .sort((a, b) => a.seq - b.seq);

      setTabs(allowedTabs);

      if (allowedTabs.length > 0) {
        setSelectedTabUrl(allowedTabs[0].url);
      }
    } catch (err) {
      console.error("Load Tabs Error:", err);
    }
  };

  /** ---------------------------------------------------------
   * Get active class
   * --------------------------------------------------------- */
  const getActiveClass = (tabUrl: string) => {
    const currentRoute = window.location.hash.replace("#", "") || "/InitiatorLanding";
    const tabRoute = tabUrl.split("#")[1] || "/InitiatorLanding";

    return currentRoute === tabRoute ? "active" : "";
  };

  /** ---------------------------------------------------------
   * Initial Load
   * --------------------------------------------------------- */
  React.useEffect(() => {
    loadTabsWithAccess();
  }, []);

  return (
    <div className="sidebar">
      {/* Header */}
      <div className="sidehead">
        <div className="logo">
          <img src={logo} width="25px" height="25px" alt="logo" />
        </div>
        <div className="sidehead-right">SONA COMSTAR</div>
      </div>

      {/* User */}
      <div className="sidehead-user">
        <i
          className="fas fa-user"
          style={{ marginLeft: "20px", marginRight: "15px" }}
        ></i>
        {username}
      </div>

      {/* Navigation */}
      <ul className="nav">
        {tabs.map((tab) => (
          <li className="nav-item" key={tab.id}>
            <a
              href={tab.url}
              className={`nav-link ${getActiveClass(tab.url)}`}
              style={{ cursor: "pointer" }}
            >
              {tab.title}
            </a>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default Sidebar;