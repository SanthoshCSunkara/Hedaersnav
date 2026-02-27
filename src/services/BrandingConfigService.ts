import { SPFI } from "@pnp/sp";
import { IBrandingConfig } from "../models/BrandingConfig";

// ─── Helpers ─────────────────────────────────────────────────────────

function parseUrlField(
  value: string | { Url?: string; Description?: string } | undefined
): string | undefined {
  if (!value) return undefined;
  if (typeof value === "string") return value;
  if (typeof value === "object" && typeof value.Url === "string") return value.Url;
  return undefined;
}

// ─── Service ─────────────────────────────────────────────────────────

const LIST_TITLE = "BrandingConfig";

export class BrandingConfigService {
  private readonly _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Fetches the active branding configuration row (Enabled eq 1).
   * If multiple rows are enabled, takes the first by Modified descending.
   * Returns undefined if no enabled row exists or on error.
   *
   * Tries to fetch all branding fields first; if that fails (e.g. columns
   * don't exist yet), falls back to fetching only HeaderRow1Bg.
   */
  public async getConfig(): Promise<IBrandingConfig | undefined> {
    // Try full fetch with all fields
    try {
      const rows = await this._sp.web.lists
        .getByTitle(LIST_TITLE)
        .items.select("HeaderRow1Bg", "MiddleRowBg", "LogoUrl", "SearchPlaceholder")
        .filter("Enabled eq 1")
        .orderBy("Modified", false)
        .top(1)();

      if (rows.length === 0) {
        console.warn("[BrandingConfigService] No enabled BrandingConfig row found.");
        return undefined;
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const raw: any = rows[0];
      return {
        headerRow1Bg: raw.HeaderRow1Bg || undefined,
        middleRowBg: raw.MiddleRowBg || undefined,
        logoUrl: parseUrlField(raw.LogoUrl),
        searchPlaceholder: raw.SearchPlaceholder || undefined,
      };
    } catch (fullErr) {
      console.warn("[BrandingConfigService] Full fetch failed, trying minimal fields:", fullErr);
    }

    // Fallback: fetch only HeaderRow1Bg (the original field)
    try {
      const rows = await this._sp.web.lists
        .getByTitle(LIST_TITLE)
        .items.select("HeaderRow1Bg")
        .filter("Enabled eq 1")
        .orderBy("Modified", false)
        .top(1)();

      if (rows.length === 0) {
        return undefined;
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const raw: any = rows[0];
      return {
        headerRow1Bg: raw.HeaderRow1Bg || undefined,
      };
    } catch (minErr) {
      console.warn("[BrandingConfigService] Minimal fetch also failed:", minErr);
      return undefined;
    }
  }
}
