import * as React from "react";
import { IFooterLinkItem } from "../../models/FooterLinkItem";
import styles from "./Footer.module.scss";

// ─── Props ───────────────────────────────────────────────────────────

export interface IFooterViewProps {
  links: IFooterLinkItem[];
  footerBg: string;
  footerTextColor: string;
  copyrightText: string | undefined;
  siteHostname: string;
}

// ─── Helpers ─────────────────────────────────────────────────────────

/**
 * Returns true if the URL points to a different hostname (external link).
 */
function isExternalUrl(url: string, siteHostname: string): boolean {
  try {
    if (url.indexOf("http://") === 0 || url.indexOf("https://") === 0) {
      const linkHost = new URL(url).hostname.toLowerCase();
      return linkHost !== siteHostname.toLowerCase();
    }
  } catch {
    // Malformed URL — treat as internal
  }
  return false;
}

// ─── Component ───────────────────────────────────────────────────────

const FooterView: React.FC<IFooterViewProps> = ({
  links,
  footerBg,
  footerTextColor,
  copyrightText,
  siteHostname,
}) => {
  const footerStyle: React.CSSProperties = {
    backgroundColor: footerBg,
    color: footerTextColor,
  };

  return (
    <footer className={styles.footer} style={footerStyle}>
      <div className={styles.footerInner}>
        {/* ── Left: Copyright Text ── */}
        {copyrightText && (
          <span className={styles.copyrightText}>{copyrightText}</span>
        )}

        {/* ── Right: Footer Links ── */}
        {links.length > 0 && (
          <nav className={styles.footerNav} aria-label="Footer links">
            {links.map((item, index) => {
              const hasUrl = !!item.url;
              const external = hasUrl && isExternalUrl(item.url as string, siteHostname);

              // Render divider before each link except the first
              const divider =
                index > 0 ? (
                  <span
                    key={"div-" + item.id}
                    className={styles.footerDivider}
                    aria-hidden="true"
                  />
                ) : undefined;

              if (!hasUrl) {
                return (
                  <React.Fragment key={item.id}>
                    {divider}
                    <span
                      className={styles.footerLinkDisabled}
                      aria-disabled="true"
                      title={item.title}
                    >
                      {item.title}
                    </span>
                  </React.Fragment>
                );
              }

              return (
                <React.Fragment key={item.id}>
                  {divider}
                  <a
                    href={item.url}
                    className={styles.footerLink}
                    target={external ? "_blank" : undefined}
                    rel={external ? "noopener noreferrer" : undefined}
                    title={item.title}
                  >
                    {item.title}
                  </a>
                </React.Fragment>
              );
            })}
          </nav>
        )}

        {/* If no copyright and no links — show nothing meaningful */}
        {!copyrightText && links.length === 0 && (
          <span className={styles.statusMessage}>Footer</span>
        )}
      </div>
    </footer>
  );
};

export default FooterView;
