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
  /** Footer background color (e.g. "#114461") */
  footerBg?: string;
  /** Footer text / link color (e.g. "#ffffff") */
  footerTextColor?: string;
  /** Copyright text rendered in footer left area */
  copyrightText?: string;
  // Future stages:
  // cacheTtlMinutes?: number;
}
