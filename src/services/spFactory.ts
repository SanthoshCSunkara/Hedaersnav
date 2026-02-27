import { SPFI, spfi, SPFx, ISPFXContext } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

/** Cached SPFI instance — avoids re-initialization on every call. */
let _sp: SPFI | undefined;

/**
 * Returns a configured SPFI instance bound to the current SPFx context.
 * Uses PnPjs v4 SPFx behavior for auth and request digest.
 *
 * @param context - The SPFx component context (ApplicationCustomizerContext or similar).
 */
export function getSP(context?: ISPFXContext): SPFI {
  if (context) {
    _sp = spfi().using(SPFx(context));
  }
  if (!_sp) {
    throw new Error("PnPjs SPFI has not been initialized. Call getSP(context) first.");
  }
  return _sp;
}
