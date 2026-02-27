/**
 * Represents a single link item from the FooterLinks SharePoint list.
 */
export interface IFooterLinkItem {
  id: number;
  title: string;
  url?: string;
  order: number;
}
