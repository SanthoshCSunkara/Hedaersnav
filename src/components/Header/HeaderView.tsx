import * as React from "react";
import { IGlobalNavItem } from "../../models/GlobalNavItem";
import styles from "./Header.module.scss";

// ─── Props ───────────────────────────────────────────────────────────

export interface IHeaderViewProps {
  items: IGlobalNavItem[];
  row1Bg: string;
  middleRowBg: string;
  logoUrl: string | undefined;
  searchPlaceholder: string;
  siteUrl: string;
  siteTitle: string;
}

// ─── Helpers ─────────────────────────────────────────────────────────

const MOBILE_PANEL_ID = "globalHeaderMobilePanel";

function hasChildren(item: IGlobalNavItem): boolean {
  return !!item.children && item.children.length > 0;
}

function dropdownId(parentId: number): string {
  return "globalNav-dd-" + parentId;
}

// ─── Component ───────────────────────────────────────────────────────

const HeaderView: React.FC<IHeaderViewProps> = ({
  items,
  row1Bg,
  middleRowBg,
  logoUrl,
  searchPlaceholder,
  siteUrl,
  siteTitle,
}) => {
  // Track which desktop dropdown is open (by parent id, or -1 for none)
  const [openDropdown, setOpenDropdown] = React.useState<number>(-1);
  const [mobileOpen, setMobileOpen] = React.useState<boolean>(false);
  // Track which mobile accordion sections are expanded
  const [mobileExpanded, setMobileExpanded] = React.useState<Set<number>>(new Set());
  // Search query state
  const [searchQuery, setSearchQuery] = React.useState<string>("");

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const headerRef = React.useRef<HTMLDivElement>(null as any);

  // ── Close dropdown on outside click or Escape ─────────────────

  React.useEffect(() => {
    const handleClick = (e: MouseEvent): void => {
      if (headerRef.current && !headerRef.current.contains(e.target as Node)) {
        setOpenDropdown(-1);
      }
    };
    const handleKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape") {
        setOpenDropdown(-1);
      }
    };
    document.addEventListener("click", handleClick);
    document.addEventListener("keydown", handleKey);
    return (): void => {
      document.removeEventListener("click", handleClick);
      document.removeEventListener("keydown", handleKey);
    };
  }, []);

  // ── Desktop dropdown toggle ───────────────────────────────────

  const toggleDropdown = React.useCallback((id: number): void => {
    setOpenDropdown((prev) => (prev === id ? -1 : id));
  }, []);

  const handleMouseEnter = React.useCallback((id: number): void => {
    setOpenDropdown(id);
  }, []);

  const handleMouseLeave = React.useCallback((): void => {
    setOpenDropdown(-1);
  }, []);

  // ── Mobile toggle ─────────────────────────────────────────────

  const toggleMobile = React.useCallback((): void => {
    setMobileOpen((prev) => !prev);
  }, []);

  const toggleMobileSection = React.useCallback((id: number): void => {
    setMobileExpanded((prev) => {
      const next = new Set<number>();
      prev.forEach((v) => next.add(v));
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  }, []);

  // ── Search handler ────────────────────────────────────────────

  const handleSearch = React.useCallback(
    (e: React.FormEvent): void => {
      e.preventDefault();
      const trimmed = searchQuery.trim();
      if (trimmed.length === 0) return;
      // Navigate to SharePoint search results page
      const searchUrl = siteUrl + "/_layouts/15/search.aspx/siteall?q=" + encodeURIComponent(trimmed);
      window.location.href = searchUrl;
    },
    [searchQuery, siteUrl]
  );

  // ── Render a single link ──────────────────────────────────────

  const renderLink = (
    item: IGlobalNavItem,
    linkClass: string,
    disabledClass: string
  ): JSX.Element => {
    const hasUrl = !!item.url;
    const isExternal = item.openInNewTab && hasUrl;

    if (!hasUrl) {
      return (
        <span
          key={item.id}
          className={disabledClass}
          title={item.title}
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
        className={linkClass}
        target={isExternal ? "_blank" : undefined}
        rel={isExternal ? "noopener noreferrer" : undefined}
        title={item.title}
      >
        {item.title}
      </a>
    );
  };

  // ── Render desktop parent item (with or without dropdown) ─────

  const renderDesktopItem = (item: IGlobalNavItem): JSX.Element => {
    if (!hasChildren(item)) {
      return renderLink(item, styles.navLink, styles.navLinkDisabled);
    }

    const isOpen = openDropdown === item.id;
    const ddId = dropdownId(item.id);

    return (
      <div
        key={item.id}
        className={styles.navItemWithDropdown}
        onMouseEnter={(): void => handleMouseEnter(item.id)}
        onMouseLeave={handleMouseLeave}
      >
        <button
          className={styles.navLink + " " + styles.navTrigger}
          onClick={(): void => toggleDropdown(item.id)}
          aria-expanded={isOpen}
          aria-controls={ddId}
          aria-haspopup="true"
          type="button"
        >
          {item.title}
          <span className={styles.caretIcon} aria-hidden="true">&#9662;</span>
        </button>
        <div
          id={ddId}
          className={isOpen ? styles.dropdownPanel : styles.dropdownPanel + " " + styles.dropdownHidden}
          role="menu"
          aria-label={item.title + " submenu"}
        >
          {(item.children || []).map((child) => {
            const childHasUrl = !!child.url;
            const childExternal = child.openInNewTab && childHasUrl;

            if (!childHasUrl) {
              return (
                <span
                  key={child.id}
                  className={styles.dropdownItemDisabled}
                  role="menuitem"
                  aria-disabled="true"
                >
                  {child.title}
                </span>
              );
            }
            return (
              <a
                key={child.id}
                href={child.url}
                className={styles.dropdownItem}
                role="menuitem"
                target={childExternal ? "_blank" : undefined}
                rel={childExternal ? "noopener noreferrer" : undefined}
              >
                {child.title}
              </a>
            );
          })}
        </div>
      </div>
    );
  };

  // ── Render mobile item (with or without accordion) ────────────

  const renderMobileItem = (item: IGlobalNavItem): JSX.Element => {
    if (!hasChildren(item)) {
      return renderLink(item, styles.mobileLink, styles.mobileLinkDisabled);
    }

    const isExpanded = mobileExpanded.has(item.id);
    const ddId = "mobileNav-dd-" + item.id;

    return (
      <div key={item.id} className={styles.mobileSection}>
        <button
          className={styles.mobileSectionTrigger}
          onClick={(): void => toggleMobileSection(item.id)}
          aria-expanded={isExpanded}
          aria-controls={ddId}
          type="button"
        >
          {item.title}
          <span className={styles.caretIcon} aria-hidden="true">
            {isExpanded ? "▴" : "▾"}
          </span>
        </button>
        {isExpanded && (
          <div id={ddId} className={styles.mobileSectionPanel} role="menu">
            {(item.children || []).map((child) =>
              renderLink(child, styles.mobileSubLink, styles.mobileLinkDisabled)
            )}
          </div>
        )}
      </div>
    );
  };

  // ── Main Render ───────────────────────────────────────────────

  return (
    <div ref={headerRef}>
      {/* ── Row 1: Global Navigation ── */}
      <header className={styles.globalHeader} style={{ backgroundColor: row1Bg }}>
        <div className={styles.headerInner}>
          {/* Desktop horizontal navigation */}
          <nav className={styles.desktopNav} aria-label="Global navigation">
            {items.map((item) => renderDesktopItem(item))}
          </nav>

          {/* Mobile hamburger button */}
          <button
            className={styles.menuButton}
            onClick={toggleMobile}
            aria-expanded={mobileOpen}
            aria-controls={MOBILE_PANEL_ID}
            type="button"
          >
            {mobileOpen ? "✕" : "☰"} Menu
          </button>
        </div>

        {/* Mobile vertical panel */}
        <nav
          id={MOBILE_PANEL_ID}
          className={
            mobileOpen ? styles.mobilePanel : styles.mobilePanel + " " + styles.mobilePanelHidden
          }
          aria-label="Global navigation (mobile)"
        >
          {items.map((item) => renderMobileItem(item))}
        </nav>
      </header>

      {/* ── Row 2: Logo + Search Bar ── */}
      <div className={styles.middleRow} style={{ backgroundColor: middleRowBg }}>
        <div className={styles.middleRowInner}>
          {/* Logo + Site Title */}
          <a href={siteUrl} className={styles.logoLink} title="Home">
            {logoUrl ? (
              <img
                src={logoUrl}
                alt={siteTitle}
                className={styles.logoImage}
              />
            ) : (
              <span className={styles.logoFallback}>&#9733;</span>
            )}
            <span className={styles.siteTitle}>{siteTitle}</span>
          </a>

          {/* Search Bar */}
          <form className={styles.searchForm} onSubmit={handleSearch} role="search">
            <input
              type="text"
              className={styles.searchInput}
              placeholder={searchPlaceholder}
              value={searchQuery}
              onChange={(e): void => setSearchQuery(e.target.value)}
              aria-label={searchPlaceholder}
            />
            <button type="submit" className={styles.searchButton} aria-label="Search">
              &#128269;
            </button>
          </form>
        </div>
      </div>
    </div>
  );
};

export default HeaderView;
