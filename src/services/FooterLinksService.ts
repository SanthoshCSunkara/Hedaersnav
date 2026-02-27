import { SPFI } from "@pnp/sp";
import { IFooterLinkItem } from "../models/FooterLinkItem";

// ─── Field‑Parsing Helpers ───────────────────────────────────────────

/**
 * SharePoint Hyperlink/URL fields may return:
 *  - a plain string
 *  - an object with { Url: string; Description?: string }
 *  - undefined
 */
function parseUrlField(
  value: string | { Url?: string; Description?: string } | undefined
): string | undefined {
  if (!value) return undefined;
  if (typeof value === "string") return value;
  if (typeof value === "object" && typeof value.Url === "string") return value.Url;
  return undefined;
}

/**
 * Normalizes a URL:
 * - http(s):// → as-is (absolute)
 * - Starts with / → tenant-absolute, use as-is
 * - Otherwise → prefix with webAbsoluteUrl + "/"
 */
function normalizeUrl(
  url: string | undefined,
  webAbsoluteUrl: string
): string | undefined {
  if (!url) return undefined;
  if (url.indexOf("http://") === 0 || url.indexOf("https://") === 0) return url;
  if (url.indexOf("/") === 0) return url;
  return webAbsoluteUrl + "/" + url;
}

// ─── Raw SharePoint Item Shape ───────────────────────────────────────

interface IRawFooterItem {
  Id: number;
  Title: string | undefined;
  Url: string | { Url?: string; Description?: string } | undefined;
  Order0: number | undefined;
  Enabled: boolean | undefined;
}

// ─── Service ─────────────────────────────────────────────────────────

const DEFAULT_ORDER = 9999;

const SELECT_FIELDS = ["Id", "Title", "Url", "Order0", "Enabled"];

export class FooterLinksService {
  private readonly _sp: SPFI;
  private readonly _webAbsoluteUrl: string;

  constructor(sp: SPFI, webAbsoluteUrl: string) {
    this._sp = sp;
    this._webAbsoluteUrl = webAbsoluteUrl;
  }

  /**
   * Maps a raw SharePoint item to the domain model.
   * Returns undefined if the item should be skipped (blank Title).
   */
  private _mapItem(raw: IRawFooterItem): IFooterLinkItem | undefined {
    if (!raw.Title || raw.Title.trim().length === 0) {
      return undefined;
    }

    const parsedUrl = parseUrlField(raw.Url);
    const normalizedUrl = normalizeUrl(parsedUrl, this._webAbsoluteUrl);

    if (!normalizedUrl) {
      console.warn(
        "[FooterLinksService] Item \"" +
          raw.Title +
          "\" (Id: " +
          raw.Id +
          ") has no valid URL — will render disabled."
      );
    }

    return {
      id: raw.Id,
      title: raw.Title,
      url: normalizedUrl,
      order: typeof raw.Order0 === "number" ? raw.Order0 : DEFAULT_ORDER,
    };
  }

  /**
   * Fetches all enabled footer links from the given list, sorted by Order0 ascending.
   *
   * @param listTitle - The SharePoint list title (default: "FooterLinks").
   */
  public async getFooterLinks(
    listTitle: string = "FooterLinks"
  ): Promise<IFooterLinkItem[]> {
    const rawItems: IRawFooterItem[] = await this._sp.web.lists
      .getByTitle(listTitle)
      .items.select(...SELECT_FIELDS)
      .filter("Enabled eq 1")
      .orderBy("Order0", true)
      .top(200)();

    const items: IFooterLinkItem[] = [];
    for (const raw of rawItems) {
      const mapped = this._mapItem(raw);
      if (mapped) {
        items.push(mapped);
      }
    }

    return items;
  }
}
