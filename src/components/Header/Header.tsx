import * as React from "react";
import { IGlobalNavItem } from "../../models/GlobalNavItem";
import { IBrandingConfig } from "../../models/BrandingConfig";
import { getSP } from "../../services/spFactory";
import { GlobalNavService } from "../../services/GlobalNavService";
import { BrandingConfigService } from "../../services/BrandingConfigService";
import HeaderView from "./HeaderView";
import styles from "./Header.module.scss";

// ─── Default Colors / Values ─────────────────────────────────────────

const DEFAULT_ROW1_BG = "#49742A";
const DEFAULT_MIDDLE_ROW_BG = "#FFFFFF";
const DEFAULT_SEARCH_PLACEHOLDER = "Search this site";

// ─── Debug Fallback ──────────────────────────────────────────────────

/**
 * Hardcoded items used ONLY when the SharePoint list fetch fails.
 * Includes children so dropdown can be tested during local debug.
 */
const DEBUG_FALLBACK_ITEMS: IGlobalNavItem[] = [
  {
    id: 901, title: "Applications", url: "/sites/IntranetConfig/SitePages/Home.aspx",
    order: 1, openInNewTab: false,
    children: [
      { id: 911, title: "Benefits Communications", url: "/sites/IntranetConfig/SitePages/Home.aspx", order: 1, openInNewTab: false, parentId: 901 },
      { id: 912, title: "Customer Benefits", url: "/sites/IntranetConfig/SitePages/Home.aspx", order: 2, openInNewTab: false, parentId: 901 },
      { id: 913, title: "Enterprise Planning", url: "/sites/IntranetConfig/SitePages/Home.aspx", order: 3, openInNewTab: false, parentId: 901 },
    ],
  },
  {
    id: 902, title: "Contracts", url: "/sites/IntranetConfig/SitePages/Home.aspx",
    order: 2, openInNewTab: false,
  },
  {
    id: 903, title: "Travel", url: "/sites/IntranetConfig/SitePages/Home.aspx",
    order: 3, openInNewTab: false,
  },
];

// ─── Props ───────────────────────────────────────────────────────────

export interface IHeaderProps {
  context: {
    pageContext: {
      web: { absoluteUrl: string };
      legacyPageContext: { formDigestTimeoutSeconds: number; formDigestValue: string };
    };
  };
  listTitle: string;
}

// ─── State ───────────────────────────────────────────────────────────

interface IHeaderState {
  items: IGlobalNavItem[];
  row1Bg: string;
  middleRowBg: string;
  logoUrl: string | undefined;
  searchPlaceholder: string;
  siteUrl: string;
  loading: boolean;
  error: boolean;
}

// ─── Container Component ─────────────────────────────────────────────

const Header: React.FC<IHeaderProps> = ({ context, listTitle }) => {
  const [state, setState] = React.useState<IHeaderState>({
    items: [],
    row1Bg: DEFAULT_ROW1_BG,
    middleRowBg: DEFAULT_MIDDLE_ROW_BG,
    logoUrl: undefined,
    searchPlaceholder: DEFAULT_SEARCH_PLACEHOLDER,
    siteUrl: context.pageContext.web.absoluteUrl,
    loading: true,
    error: false,
  });

  React.useEffect(() => {
    let cancelled = false;

    const fetchData = async (): Promise<void> => {
      try {
        const sp = getSP(context);
        const webAbsoluteUrl = context.pageContext.web.absoluteUrl;

        // Fetch nav hierarchy and branding in parallel
        const navService = new GlobalNavService(sp, webAbsoluteUrl);
        const brandingService = new BrandingConfigService(sp);

        const [items, brandingConfig] = await Promise.all([
          navService.getNavHierarchy(listTitle),
          brandingService.getConfig(),
        ]);

        console.log("[Header] Fetched " + items.length + " top-level nav items from SharePoint");

        const branding: IBrandingConfig = brandingConfig || {};
        const row1Bg = branding.headerRow1Bg || DEFAULT_ROW1_BG;
        const middleRowBg = branding.middleRowBg || DEFAULT_MIDDLE_ROW_BG;
        const logoUrl = branding.logoUrl || undefined;
        const searchPlaceholder = branding.searchPlaceholder || DEFAULT_SEARCH_PLACEHOLDER;

        if (!cancelled) {
          setState({
            items,
            row1Bg,
            middleRowBg,
            logoUrl,
            searchPlaceholder,
            siteUrl: webAbsoluteUrl,
            loading: false,
            error: false,
          });
        }
      } catch (err) {
        console.error("[Header] Navigation load failed, using debug fallback:", err);
        if (!cancelled) {
          setState({
            items: DEBUG_FALLBACK_ITEMS,
            row1Bg: DEFAULT_ROW1_BG,
            middleRowBg: DEFAULT_MIDDLE_ROW_BG,
            logoUrl: undefined,
            searchPlaceholder: DEFAULT_SEARCH_PLACEHOLDER,
            siteUrl: context.pageContext.web.absoluteUrl,
            loading: false,
            error: false,
          });
        }
      }
    };

    fetchData()
      .catch(() => { /* handled above */ });

    return (): void => {
      cancelled = true;
    };
  }, [context, listTitle]);

  // ─── Render States ───────────────────────────────────────────────

  if (state.loading) {
    return (
      <header className={styles.globalHeader} style={{ backgroundColor: state.row1Bg }}>
        <div className={styles.headerInner}>
          <span className={styles.statusMessage}>Loading navigation…</span>
        </div>
      </header>
    );
  }

  if (state.error) {
    return (
      <header className={styles.globalHeader} style={{ backgroundColor: state.row1Bg }}>
        <div className={styles.headerInner}>
          <span className={styles.statusMessage}>Couldn&apos;t load navigation.</span>
        </div>
      </header>
    );
  }

  if (state.items.length === 0) {
    return (
      <header className={styles.globalHeader} style={{ backgroundColor: state.row1Bg }}>
        <div className={styles.headerInner}>
          <span className={styles.statusMessage}>No navigation links configured.</span>
        </div>
      </header>
    );
  }

  return (
    <HeaderView
      items={state.items}
      row1Bg={state.row1Bg}
      middleRowBg={state.middleRowBg}
      logoUrl={state.logoUrl}
      searchPlaceholder={state.searchPlaceholder}
      siteUrl={state.siteUrl}
    />
  );
};

export default Header;
