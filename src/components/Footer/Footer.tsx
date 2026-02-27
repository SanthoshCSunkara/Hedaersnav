import * as React from "react";
import { IFooterLinkItem } from "../../models/FooterLinkItem";
import { IBrandingConfig } from "../../models/BrandingConfig";
import { getSP } from "../../services/spFactory";
import { FooterLinksService } from "../../services/FooterLinksService";
import { BrandingConfigService } from "../../services/BrandingConfigService";
import FooterView from "./FooterView";
import styles from "./Footer.module.scss";

// ─── Default Colors / Values ─────────────────────────────────────────

const DEFAULT_FOOTER_BG = "#114461";
const DEFAULT_FOOTER_TEXT_COLOR = "#ffffff";
const FOOTER_LINKS_LIST = "FooterLinks";

// ─── Props ───────────────────────────────────────────────────────────

export interface IFooterProps {
  context: {
    pageContext: {
      web: { absoluteUrl: string; title: string };
      legacyPageContext: { formDigestTimeoutSeconds: number; formDigestValue: string };
    };
  };
}

// ─── State ───────────────────────────────────────────────────────────

interface IFooterState {
  links: IFooterLinkItem[];
  footerBg: string;
  footerTextColor: string;
  copyrightText: string | undefined;
  siteHostname: string;
  loading: boolean;
}

// ─── Helper ──────────────────────────────────────────────────────────

/**
 * Extracts hostname from a URL string. Falls back to empty string.
 */
function extractHostname(absoluteUrl: string): string {
  try {
    return new URL(absoluteUrl).hostname;
  } catch {
    return "";
  }
}

// ─── Container Component ─────────────────────────────────────────────

const Footer: React.FC<IFooterProps> = ({ context }) => {
  const [state, setState] = React.useState<IFooterState>({
    links: [],
    footerBg: DEFAULT_FOOTER_BG,
    footerTextColor: DEFAULT_FOOTER_TEXT_COLOR,
    copyrightText: undefined,
    siteHostname: extractHostname(context.pageContext.web.absoluteUrl),
    loading: true,
  });

  React.useEffect(() => {
    let cancelled = false;

    const fetchData = async (): Promise<void> => {
      let links: IFooterLinkItem[] = [];
      let branding: IBrandingConfig = {};

      // Fetch branding config — non-blocking (footer still renders with defaults)
      try {
        const sp = getSP();
        const brandingService = new BrandingConfigService(sp);
        const config = await brandingService.getConfig();
        if (config) {
          branding = config;
        }
      } catch (brandingErr) {
        console.warn("[Footer] BrandingConfig fetch failed (non-blocking):", brandingErr);
      }

      // Fetch footer links — non-blocking (footer still renders copyright only)
      try {
        const sp = getSP();
        const webAbsoluteUrl = context.pageContext.web.absoluteUrl;
        const linksService = new FooterLinksService(sp, webAbsoluteUrl);
        links = await linksService.getFooterLinks(FOOTER_LINKS_LIST);
        console.log("[Footer] Fetched " + links.length + " footer links");
      } catch (linksErr) {
        console.warn("[Footer] FooterLinks fetch failed (non-blocking):", linksErr);
      }

      if (!cancelled) {
        setState({
          links,
          footerBg: branding.footerBg || DEFAULT_FOOTER_BG,
          footerTextColor: branding.footerTextColor || DEFAULT_FOOTER_TEXT_COLOR,
          copyrightText: branding.copyrightText || undefined,
          siteHostname: extractHostname(context.pageContext.web.absoluteUrl),
          loading: false,
        });
      }
    };

    fetchData().catch(() => {
      /* errors handled inside fetchData */
    });

    return (): void => {
      cancelled = true;
    };
  }, [context]);

  // ─── Render ────────────────────────────────────────────────────

  if (state.loading) {
    return (
      <footer
        className={styles.footer}
        style={{ backgroundColor: state.footerBg, color: state.footerTextColor }}
      >
        <div className={styles.footerInner}>
          <span className={styles.statusMessage}>Loading footer…</span>
        </div>
      </footer>
    );
  }

  return (
    <FooterView
      links={state.links}
      footerBg={state.footerBg}
      footerTextColor={state.footerTextColor}
      copyrightText={state.copyrightText}
      siteHostname={state.siteHostname}
    />
  );
};

export default Footer;
