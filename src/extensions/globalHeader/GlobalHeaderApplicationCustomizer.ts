import * as React from "react";
import * as ReactDom from "react-dom";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";

import Header, { IHeaderProps } from "../../components/Header/Header";
import Footer, { IFooterProps } from "../../components/Footer/Footer";
import { getSP } from "../../services/spFactory";

export interface IGlobalHeaderApplicationCustomizerProperties {
  listTitle: string;
}

/** CSS injected once to hide the default SharePoint site navigation */
const HIDE_SITE_NAV_CSS = `
  /* Hide the default horizontal site navigation bar */
  [data-automationid="HorizontalNav"] {
    display: none !important;
  }
  /* Hide the default site header row (site logo + title + nav links) */
  [data-automationid="SiteHeader"] {
    display: none !important;
  }
`;

/** Unique DOM id used to prevent duplicate footer host elements */
const FOOTER_HOST_ID = "ppr-global-footer-host";

export default class GlobalHeaderApplicationCustomizer extends BaseApplicationCustomizer<IGlobalHeaderApplicationCustomizerProperties> {
  // ─── Header State ──────────────────────────────────────────────
  private _headerPlaceholder: PlaceholderContent | undefined;
  private _isHeaderRendered: boolean = false;
  private _headerRenderContainer: HTMLElement | undefined;

  // ─── Footer State ──────────────────────────────────────────────
  private _footerPlaceholder: PlaceholderContent | undefined;
  private _isFooterRendered: boolean = false;
  private _footerRenderContainer: HTMLElement | undefined;

  // ─── Shared ────────────────────────────────────────────────────
  private _styleElement: HTMLStyleElement | undefined;

  public onInit(): Promise<void> {
    console.log("[GlobalHeader] onInit reached");

    // Initialize PnPjs with the SPFx context (once)
    getSP(this.context as unknown as import("@pnp/sp").ISPFXContext);

    // Hide the default SharePoint site navigation
    this._injectHideNavCss();

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceholders
    );

    this._renderPlaceholders();

    return Promise.resolve();
  }

  /**
   * Injects a <style> tag to hide the default SharePoint horizontal nav.
   */
  private _injectHideNavCss(): void {
    if (this._styleElement) {
      return;
    }
    const style: HTMLStyleElement = document.createElement("style");
    style.type = "text/css";
    style.textContent = HIDE_SITE_NAV_CSS;
    document.head.appendChild(style);
    this._styleElement = style;
    console.log("[GlobalHeader] Injected CSS to hide default site navigation");
  }

  // ═══════════════════════════════════════════════════════════════════
  //  RENDER ORCHESTRATOR
  // ═══════════════════════════════════════════════════════════════════

  private _renderPlaceholders(): void {
    this._renderHeader();
    this._renderFooter();
  }

  // ═══════════════════════════════════════════════════════════════════
  //  HEADER
  // ═══════════════════════════════════════════════════════════════════

  private _renderHeader(): void {
    if (this._isHeaderRendered) {
      return;
    }

    console.log("[GlobalHeader] _renderHeader called");

    // Try Top placeholder
    if (!this._headerPlaceholder) {
      this._headerPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDisposeHeader.bind(this) }
      );
      if (this._headerPlaceholder) {
        console.log("[GlobalHeader] Top placeholder acquired");
      }
    }

    // Render into placeholder if available
    if (this._headerPlaceholder && this._headerPlaceholder.domElement) {
      console.log("[GlobalHeader] Rendering React Header into placeholder");
      this._mountHeader(this._headerPlaceholder.domElement);
      return;
    }

    // Fallback: try inserting after SuiteNavPlaceHolder
    const suiteNav: HTMLElement | undefined =
      document.getElementById("SuiteNavPlaceHolder") || undefined;
    if (suiteNav && suiteNav.parentElement) {
      console.log("[GlobalHeader] Rendering after #SuiteNavPlaceHolder (fallback)");
      const wrapper: HTMLDivElement = document.createElement("div");
      wrapper.id = "globalHeaderWrapper";
      suiteNav.parentElement.insertBefore(wrapper, suiteNav.nextSibling);
      this._mountHeader(wrapper);
      return;
    }

    // Last resort: prepend to body
    if (document.body) {
      console.log("[GlobalHeader] Rendering at body prepend (last resort)");
      const bodyWrapper: HTMLDivElement = document.createElement("div");
      bodyWrapper.id = "globalHeaderWrapper";
      document.body.insertBefore(bodyWrapper, document.body.firstChild);
      this._mountHeader(bodyWrapper);
      return;
    }

    console.warn("[GlobalHeader] No header anchor found. Will retry on changedEvent.");
  }

  /**
   * Mounts the React Header into the given container.
   */
  private _mountHeader(container: HTMLElement): void {
    const listTitle: string = this.properties.listTitle || "GlobalNav";

    const element: React.ReactElement<IHeaderProps> = React.createElement(
      Header,
      {
        context: this.context as unknown as IHeaderProps["context"],
        listTitle: listTitle,
      }
    );

    ReactDom.render(element, container);
    this._headerRenderContainer = container;
    this._isHeaderRendered = true;
    console.log("[GlobalHeader] React Header component rendered");
  }

  private _onDisposeHeader(): void {
    console.log("[GlobalHeader] Disposing header");

    if (this._headerRenderContainer) {
      ReactDom.unmountComponentAtNode(this._headerRenderContainer);

      // If we created a fallback wrapper (not a placeholder), remove it from DOM
      if (
        this._headerRenderContainer.id === "globalHeaderWrapper" &&
        this._headerRenderContainer.parentElement
      ) {
        this._headerRenderContainer.parentElement.removeChild(this._headerRenderContainer);
      }

      this._headerRenderContainer = undefined;
    }

    this._headerPlaceholder = undefined;
    this._isHeaderRendered = false;
  }

  // ═══════════════════════════════════════════════════════════════════
  //  FOOTER
  // ═══════════════════════════════════════════════════════════════════

  private _renderFooter(): void {
    if (this._isFooterRendered) {
      return;
    }

    console.log("[GlobalFooter] _renderFooter called");

    // Try Bottom placeholder
    if (!this._footerPlaceholder) {
      this._footerPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDisposeFooter.bind(this) }
      );
      if (this._footerPlaceholder) {
        console.log("[GlobalFooter] Bottom placeholder acquired");
      }
    }

    // Render into placeholder if available
    if (this._footerPlaceholder && this._footerPlaceholder.domElement) {
      console.log("[GlobalFooter] Rendering React Footer into Bottom placeholder");
      this._mountFooter(this._footerPlaceholder.domElement);
      return;
    }

    // Fallback: append to body (prevent duplicates via unique id)
    if (document.body) {
      const existing = document.getElementById(FOOTER_HOST_ID);
      if (existing) {
        console.log("[GlobalFooter] Reusing existing footer host element");
        this._mountFooter(existing);
        return;
      }

      console.log("[GlobalFooter] Appending footer host to body (fallback)");
      const footerHost: HTMLDivElement = document.createElement("div");
      footerHost.id = FOOTER_HOST_ID;
      document.body.appendChild(footerHost);
      this._mountFooter(footerHost);
      return;
    }

    console.warn("[GlobalFooter] No footer anchor found. Will retry on changedEvent.");
  }

  /**
   * Mounts the React Footer into the given container.
   */
  private _mountFooter(container: HTMLElement): void {
    const element: React.ReactElement<IFooterProps> = React.createElement(
      Footer,
      {
        context: this.context as unknown as IFooterProps["context"],
      }
    );

    ReactDom.render(element, container);
    this._footerRenderContainer = container;
    this._isFooterRendered = true;
    console.log("[GlobalFooter] React Footer component rendered");
  }

  private _onDisposeFooter(): void {
    console.log("[GlobalFooter] Disposing footer");

    if (this._footerRenderContainer) {
      ReactDom.unmountComponentAtNode(this._footerRenderContainer);

      // If we created a fallback host (not a placeholder), remove it from DOM
      if (
        this._footerRenderContainer.id === FOOTER_HOST_ID &&
        this._footerRenderContainer.parentElement
      ) {
        this._footerRenderContainer.parentElement.removeChild(this._footerRenderContainer);
      }

      this._footerRenderContainer = undefined;
    }

    this._footerPlaceholder = undefined;
    this._isFooterRendered = false;
  }

  // ═══════════════════════════════════════════════════════════════════
  //  GLOBAL DISPOSE
  // ═══════════════════════════════════════════════════════════════════

  protected onDispose(): void {
    console.log("[GlobalHeader] Disposing all");

    this._onDisposeHeader();
    this._onDisposeFooter();

    // Remove injected style
    if (this._styleElement && this._styleElement.parentNode) {
      this._styleElement.parentNode.removeChild(this._styleElement);
      this._styleElement = undefined;
    }

    super.onDispose();
  }
}
