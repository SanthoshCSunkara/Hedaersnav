import * as React from "react";
import * as ReactDom from "react-dom";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";

import Header, { IHeaderProps } from "../../components/Header/Header";
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
`;

export default class GlobalHeaderApplicationCustomizer extends BaseApplicationCustomizer<IGlobalHeaderApplicationCustomizerProperties> {
  private _placeholder: PlaceholderContent | undefined;
  private _isRendered: boolean = false;
  private _renderContainer: HTMLElement | undefined;
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

  private _renderPlaceholders(): void {
    if (this._isRendered) {
      return;
    }

    console.log("[GlobalHeader] _renderPlaceholders called");

    // Try Top placeholder
    if (!this._placeholder) {
      this._placeholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose.bind(this) }
      );
      if (this._placeholder) {
        console.log("[GlobalHeader] Top placeholder acquired");
      }
    }

    // Fallback: try Bottom placeholder
    if (!this._placeholder) {
      this._placeholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose.bind(this) }
      );
      if (this._placeholder) {
        console.log("[GlobalHeader] Bottom placeholder acquired (fallback)");
      }
    }

    // Render into placeholder if available
    if (this._placeholder && this._placeholder.domElement) {
      console.log("[GlobalHeader] Rendering React Header into placeholder");
      this._mountReactComponent(this._placeholder.domElement);
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
      this._mountReactComponent(wrapper);
      return;
    }

    // Last resort: prepend to body
    if (document.body) {
      console.log("[GlobalHeader] Rendering at body prepend (last resort)");
      const bodyWrapper: HTMLDivElement = document.createElement("div");
      bodyWrapper.id = "globalHeaderWrapper";
      document.body.insertBefore(bodyWrapper, document.body.firstChild);
      this._mountReactComponent(bodyWrapper);
      return;
    }

    console.warn("[GlobalHeader] No anchor found. Will retry on changedEvent.");
  }

  /**
   * Mounts the React Header into the given container and tracks it
   * so _onDispose can cleanly unmount.
   */
  private _mountReactComponent(container: HTMLElement): void {
    const listTitle: string = this.properties.listTitle || "GlobalNav";

    const element: React.ReactElement<IHeaderProps> = React.createElement(
      Header,
      {
        context: this.context as unknown as IHeaderProps["context"],
        listTitle: listTitle,
      }
    );

    ReactDom.render(element, container);
    this._renderContainer = container;
    this._isRendered = true;
    console.log("[GlobalHeader] React Header component rendered");
  }

  private _onDispose(): void {
    console.log("[GlobalHeader] Disposing");

    // Unmount React from the tracked container
    if (this._renderContainer) {
      ReactDom.unmountComponentAtNode(this._renderContainer);

      // If we created a fallback wrapper (not a placeholder), remove it from DOM
      if (this._renderContainer.id === "globalHeaderWrapper" && this._renderContainer.parentElement) {
        this._renderContainer.parentElement.removeChild(this._renderContainer);
      }

      this._renderContainer = undefined;
    }

    // Remove injected style
    if (this._styleElement && this._styleElement.parentNode) {
      this._styleElement.parentNode.removeChild(this._styleElement);
      this._styleElement = undefined;
    }

    this._placeholder = undefined;
    this._isRendered = false;
  }
}
