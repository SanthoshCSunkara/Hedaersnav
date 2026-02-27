/**
 * Represents a navigation item from the GlobalNav SharePoint list.
 * Top-level items may contain nested children for dropdown menus.
 */
export interface IGlobalNavItem {
  id: number;
  title: string;
  url?: string;
  order: number;
  openInNewTab: boolean;
  parentId?: number;
  menuType?: string;
  megaColumn?: number;
  audienceGroups?: string;
  iconName?: string;
  colorTag?: string;
  /** Child items for dropdown — populated by service hierarchy builder. */
  children?: IGlobalNavItem[];
}
