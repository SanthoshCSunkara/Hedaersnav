/**
 * Represents branding settings from the BrandingConfig SharePoint list.
 * Extended to include Row 2 (logo + search) properties.
 */
export interface IBrandingConfig {
  headerRow1Bg?: string;
  middleRowBg?: string;
  logoUrl?: string;
  searchPlaceholder?: string;
  // Future stages will add:
  // headerRow2Bg?: string;
  // hoverColor?: string;
  // footerBg?: string;
  // footerTextColor?: string;
  // copyrightText?: string;
  // cacheTtlMinutes?: number;
}
