import { SPFI } from "@pnp/sp";
import { IGlobalNavItem } from "../models/GlobalNavItem";

// ─── Field‑Parsing Helpers ───────────────────────────────────────────

/**
 * SharePoint Hyperlink/URL fields may return:
 *  - a plain string
 *  - an object with { Url: string; Description?: string }
 *  - undefined
 */
function parseUrlField(value: string | { Url?: string; Description?: string } | undefined): string | undefined {
  if (!value) return undefined;
  if (typeof value === "string") return value;
  if (typeof value === "object" && typeof value.Url === "string") return value.Url;
  return undefined;
}

/**
 * Lookup fields may be returned in expanded or flat form.
 * Expanded: item[fieldName] = { Id: 5 }
 * Flat:     item[fieldName + "Id"] = 5
 */
function getLookupId(item: Record<string, unknown>, fieldName: string): number | undefined {
  const expanded = item[fieldName];
  if (expanded && typeof expanded === "object" && "Id" in (expanded as Record<string, unknown>)) {
    const id = (expanded as Record<string, unknown>).Id;
    return typeof id === "number" ? id : undefined;
  }
  const flat = item[fieldName + "Id"];
  if (typeof flat === "number") return flat;
  return undefined;
}

/**
 * Normalizes a URL:
 * - http(s):// → as-is (absolute)
 * - Starts with / → tenant-absolute, use as-is
 * - Otherwise → prefix with webAbsoluteUrl + "/"
 */
function normalizeUrl(url: string | undefined, webAbsoluteUrl: string): string | undefined {
  if (!url) return undefined;
  if (url.indexOf("http://") === 0 || url.indexOf("https://") === 0) return url;
  if (url.indexOf("/") === 0) return url;
  return webAbsoluteUrl + "/" + url;
}

// ─── Raw SharePoint Item Shape ───────────────────────────────────────

interface IRawNavItem {
  Id: number;
  Title: string | undefined;
  Url: string | { Url?: string; Description?: string } | undefined;
  Order0: number | undefined;
  Enabled: boolean | undefined;
  ParentId: { Id: number } | undefined;
  ParentIdId: number | undefined;
  MenuType: string | undefined;
  MegaColumn: number | undefined;
  AudienceGroups: string | undefined;
  OpenInNewTab: boolean | undefined;
  IconName: string | undefined;
}

// ─── Service ─────────────────────────────────────────────────────────

const DEFAULT_ORDER = 9999;

const SELECT_FIELDS = [
  "Id",
  "Title",
  "Url",
  "Order0",
  "Enabled",
  "ParentId/Id",
  "MenuType",
  "MegaColumn",
  "AudienceGroups",
  "OpenInNewTab",
  "IconName",
];

const EXPAND_FIELDS = ["ParentId"];

export class GlobalNavService {
  private readonly _sp: SPFI;
  private readonly _webAbsoluteUrl: string;

  constructor(sp: SPFI, webAbsoluteUrl: string) {
    this._sp = sp;
    this._webAbsoluteUrl = webAbsoluteUrl;
  }

  /**
   * Maps a raw SharePoint item to our domain model.
   */
  private _mapItem(raw: IRawNavItem): IGlobalNavItem | undefined {
    if (!raw.Title || raw.Title.trim().length === 0) {
      return undefined;
    }

    const parentId = getLookupId(raw as unknown as Record<string, unknown>, "ParentId");
    const parsedUrl = parseUrlField(raw.Url);
    const normalizedUrl = normalizeUrl(parsedUrl, this._webAbsoluteUrl);

    if (!normalizedUrl) {
      console.warn("[GlobalNavService] Item \"" + raw.Title + "\" (Id: " + raw.Id + ") has no valid URL.");
    }

    return {
      id: raw.Id,
      title: raw.Title,
      url: normalizedUrl,
      order: typeof raw.Order0 === "number" ? raw.Order0 : DEFAULT_ORDER,
      openInNewTab: raw.OpenInNewTab === true,
      parentId: parentId,
      menuType: raw.MenuType || undefined,
      megaColumn: typeof raw.MegaColumn === "number" ? raw.MegaColumn : undefined,
      audienceGroups: raw.AudienceGroups || undefined,
      iconName: raw.IconName || undefined,
    };
  }

  /**
   * Fetches ALL enabled navigation items and builds a parent→children hierarchy.
   *
   * Returns top-level items (no ParentId) sorted by Order0.
   * Each top-level item's `.children` array contains its child items, also sorted.
   */
  public async getNavHierarchy(listTitle: string): Promise<IGlobalNavItem[]> {
    const rawItems: IRawNavItem[] = await this._sp.web.lists
      .getByTitle(listTitle)
      .items.select(...SELECT_FIELDS)
      .expand(...EXPAND_FIELDS)
      .filter("Enabled eq 1")
      .orderBy("Order0", true)
      .top(500)();

    // Map all raw items
    const allItems: IGlobalNavItem[] = [];
    for (const raw of rawItems) {
      const mapped = this._mapItem(raw);
      if (mapped) {
        allItems.push(mapped);
      }
    }

    // Separate parents and children
    const parents: IGlobalNavItem[] = [];
    const childMap: Map<number, IGlobalNavItem[]> = new Map();

    for (const item of allItems) {
      if (item.parentId === undefined) {
        item.children = [];
        parents.push(item);
      } else {
        const existing = childMap.get(item.parentId);
        if (existing) {
          existing.push(item);
        } else {
          childMap.set(item.parentId, [item]);
        }
      }
    }

    // Attach children to parents (already sorted by Order0 from query)
    for (const parent of parents) {
      const kids = childMap.get(parent.id);
      if (kids && kids.length > 0) {
        parent.children = kids;
      }
    }

    return parents;
  }

  /**
   * Fetches top-level (no parent) navigation items from the given list.
   * @deprecated Use getNavHierarchy() instead for dropdown support.
   */
  public async getTopLevelLinks(listTitle: string): Promise<IGlobalNavItem[]> {
    const hierarchy = await this.getNavHierarchy(listTitle);
    // Strip children for backward compat
    for (const item of hierarchy) {
      item.children = undefined;
    }
    return hierarchy;
  }
}
