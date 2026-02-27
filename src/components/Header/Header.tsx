import * as React from "react";
import { IGlobalNavItem } from "../../models/GlobalNavItem";
import { IBrandingConfig } from "../../models/BrandingConfig";
import { getSP } from "../../services/spFactory";
import { GlobalNavService } from "../../services/GlobalNavService";
import { BrandingConfigService } from "../../services/BrandingConfigService";
import HeaderView from "./HeaderView";
import DivisionNav from "../DivisionNav/DivisionNav";
import styles from "./Header.module.scss";

// ─── Default Colors / Values ─────────────────────────────────────────

const DEFAULT_ROW1_BG = "#49742A";
const DEFAULT_MIDDLE_ROW_BG = "#FFFFFF";
const DEFAULT_ROW2_BG = "#99AE7A";
const DEFAULT_SEARCH_PLACEHOLDER = "Search this site";
const DIVISION_NAV_LIST = "DivisionNav";

// ─── Debug Fallback ──────────────────────────────────────────────────

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
      web: { absoluteUrl: string; title: string };
      legacyPageContext: { formDigestTimeoutSeconds: number; formDigestValue: string };
    };
  };
  listTitle: string;
}

// ─── State ───────────────────────────────────────────────────────────

interface IHeaderState {
  items: IGlobalNavItem[];
  divisionItems: IGlobalNavItem[];
  row1Bg: string;
  middleRowBg: string;
  row2Bg: string;
  hoverColor: string | undefined;
  logoUrl: string | undefined;
  searchPlaceholder: string;
  siteUrl: string;
  siteTitle: string;
  loading: boolean;
  error: boolean;
}

// ─── Helper ──────────────────────────────────────────────────────────

/** Normalizes a logo URL — makes relative paths absolute */
function normalizeLogoUrl(url: string | undefined, webAbsoluteUrl: string): string | undefined {
  if (!url) return undefined;
  if (url.indexOf("http://") === 0 || url.indexOf("https://") === 0) return url;
  if (url.indexOf("/") === 0) return url;
  return webAbsoluteUrl + "/" + url;
}

// ─── Container Component ─────────────────────────────────────────────

const Header: React.FC<IHeaderProps> = ({ context, listTitle }) => {
  const [state, setState] = React.useState<IHeaderState>({
    items: [],
    divisionItems: [],
    row1Bg: DEFAULT_ROW1_BG,
    middleRowBg: DEFAULT_MIDDLE_ROW_BG,
    row2Bg: DEFAULT_ROW2_BG,
    hoverColor: undefined,
    logoUrl: undefined,
    searchPlaceholder: DEFAULT_SEARCH_PLACEHOLDER,
    siteUrl: context.pageContext.web.absoluteUrl,
    siteTitle: context.pageContext.web.title || "Intranet",
    loading: true,
    error: false,
  });

  React.useEffect(() => {
    let cancelled = false;

    const fetchData = async (): Promise<void> => {
      try {
        const sp = getSP(context);
        const webAbsoluteUrl = context.pageContext.web.absoluteUrl;

        const navService = new GlobalNavService(sp, webAbsoluteUrl);
        const brandingService = new BrandingConfigService(sp);

        // Fetch GlobalNav, DivisionNav, and branding in parallel
        // DivisionNav is wrapped in its own catch so it can't break the main fetch
        const [items, brandingConfig, divisionItems] = await Promise.all([
          navService.getNavHierarchy(listTitle),
          brandingService.getConfig(),
          navService.getNavHierarchy(DIVISION_NAV_LIST).catch((divErr) => {
            console.warn("[Header] DivisionNav fetch failed (non-blocking):", divErr);
            return [] as IGlobalNavItem[];
          }),
        ]);

        console.log("[Header] Fetched " + items.length + " GlobalNav + " + divisionItems.length + " DivisionNav items");

        const branding: IBrandingConfig = brandingConfig || {};
        const row1Bg = branding.headerRow1Bg || DEFAULT_ROW1_BG;
        const middleRowBg = branding.middleRowBg || DEFAULT_MIDDLE_ROW_BG;
        const row2Bg = branding.headerRow2Bg || DEFAULT_ROW2_BG;
        const hoverColor = branding.hoverColor || undefined;
        const logoUrl = normalizeLogoUrl(branding.logoUrl, webAbsoluteUrl);
        const searchPlaceholder = branding.searchPlaceholder || DEFAULT_SEARCH_PLACEHOLDER;

        if (!cancelled) {
          setState({
            items,
            divisionItems,
            row1Bg,
            middleRowBg,
            row2Bg,
            hoverColor,
            logoUrl,
            searchPlaceholder,
            siteUrl: webAbsoluteUrl,
            siteTitle: context.pageContext.web.title || "Intranet",
            loading: false,
            error: false,
          });
        }
      } catch (err) {
        console.error("[Header] Navigation load failed, using debug fallback:", err);
        if (!cancelled) {
          setState({
            items: DEBUG_FALLBACK_ITEMS,
            divisionItems: [],
            row1Bg: DEFAULT_ROW1_BG,
            middleRowBg: DEFAULT_MIDDLE_ROW_BG,
            row2Bg: DEFAULT_ROW2_BG,
            hoverColor: undefined,
            logoUrl: undefined,
            searchPlaceholder: DEFAULT_SEARCH_PLACEHOLDER,
            siteUrl: context.pageContext.web.absoluteUrl,
            siteTitle: context.pageContext.web.title || "Intranet",
            loading: false,
            error: false,
          });
        }
      }
    };

    fetchData().catch(() => { /* handled above */ });

    return (): void => { cancelled = true; };
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
    <div>
      {/* Row 1 + Row 2 (existing header — safe) */}
      <HeaderView
        items={state.items}
        row1Bg={state.row1Bg}
        middleRowBg={state.middleRowBg}
        logoUrl={state.logoUrl}
        searchPlaceholder={state.searchPlaceholder}
        siteUrl={state.siteUrl}
        siteTitle={state.siteTitle}
      />

      {/* Row 3 — DivisionNav (separate container — isolated) */}
      {state.divisionItems.length > 0 && (
        <DivisionNav
          items={state.divisionItems}
          row2Bg={state.row2Bg}
          hoverColor={state.hoverColor}
        />
      )}
    </div>
  );
};

export default Header;
