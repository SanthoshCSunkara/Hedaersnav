/**
 * Represents branding settings from the BrandingConfig SharePoint list.
 * Extended to include all three header rows and theming properties.
 */
export interface IBrandingConfig {
  headerRow1Bg?: string;
  middleRowBg?: string;
  headerRow2Bg?: string;
  hoverColor?: string;
  logoUrl?: string;
  searchPlaceholder?: string;
  // Future stages:
  // footerBg?: string;
  // footerTextColor?: string;
  // copyrightText?: string;
  // cacheTtlMinutes?: number;
}
