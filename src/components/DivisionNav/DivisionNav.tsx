import * as React from "react";
import { IGlobalNavItem } from "../../models/GlobalNavItem";
import styles from "./DivisionNav.module.scss";

// ─── Props ───────────────────────────────────────────────────────────

export interface IDivisionNavProps {
  items: IGlobalNavItem[];
  row2Bg: string;
  hoverColor: string | undefined;
}

// ─── Component ───────────────────────────────────────────────────────

const DivisionNav: React.FC<IDivisionNavProps> = ({ items, row2Bg }) => {
  if (items.length === 0) {
    return null;
  }

  return (
    <nav
      className={styles.divisionNav}
      style={{ backgroundColor: row2Bg }}
      aria-label="Division navigation"
    >
      <div className={styles.divisionNavInner}>
        {items.map((item) => {
          const hasUrl = !!item.url;
          const isExternal = item.openInNewTab && hasUrl;

          if (!hasUrl) {
            return (
              <span
                key={item.id}
                className={styles.divNavLinkDisabled}
                aria-disabled="true"
              >
                {item.title}
              </span>
            );
          }

          return (
            <a
              key={item.id}
              href={item.url}
              className={styles.divNavLink}
              target={isExternal ? "_blank" : undefined}
              rel={isExternal ? "noopener noreferrer" : undefined}
              title={item.title}
            >
              {item.title}
            </a>
          );
        })}
      </div>
    </nav>
  );
};

export default DivisionNav;
