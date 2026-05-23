export type CaptionKind = 'Figure' | 'Table';
export type CaptionPlacement = 'above' | 'below';

export interface CaptionTypography {
  font: string;
  size: number;
  alignment: 'left' | 'center' | 'right';
}

export interface CaptionSettings {
  toc: boolean;
  listFigures: boolean;
  listTables: boolean;
  figureCaptionPlacement: CaptionPlacement;
  tableCaptionPlacement: CaptionPlacement;
  typography?: CaptionTypography;
}

export interface DocumentAudit {
  totalFigures: number;
  figuresWithCaptions: number;
  figuresMissingCaptions: number;
  totalTables: number;
  tablesWithCaptions: number;
  tablesMissingCaptions: number;
  hasToc: boolean;
  hasTableOfFigures: boolean;
  hasTableOfTables: boolean;
}

export function fallbackCaption(kind: CaptionKind): string {
  return `${kind}: Add description`;
}

export function cleanCaptionText(caption: string, kind: CaptionKind): string {
  const trimmed = caption.trim();
  if (!trimmed) return 'Add description';

  const cleaned = trimmed
    .replace(new RegExp(`^(?:${kind}|${kind === 'Figure' ? 'Fig\\.' : 'Tbl\\.'})\\s*\\d*[\\s:.-]*`, 'i'), '')
    .trim();
  return sanitizeCaptionDescription(cleaned);
}

export function startsLikeCaption(text: string, kind: CaptionKind): boolean {
  const trimmed = text.trim();
  if (kind === 'Figure') {
    return /^(figure|fig\.)\s*(\d+|[ivxlcdm]+)?\s*[:.)-]?/i.test(trimmed);
  }
  if (/^table of (contents|figures|tables)\b/i.test(trimmed)) return false;
  return /^table\s*(\d+|[ivxlcdm]+)?\s*[:.)-]?/i.test(trimmed);
}

export function sanitizeCaptionDescription(value: string | undefined): string {
  const cleaned = (value ?? '').replace(/\s+/g, ' ').trim();
  if (!cleaned) return 'Add description';
  if (looksLikeLocalFilePath(cleaned) || looksLikeImageFileName(cleaned)) return 'Add description';
  return cleaned;
}

export function looksLikeLocalFilePath(value: string): boolean {
  const normalized = value.trim();
  return /^[a-z]:[\\/]/i.test(normalized) ||
    /^\\\\[^\\]+\\/i.test(normalized) ||
    /[\\/](users|desktop|downloads|documents|onedrive|pictures|lab\d*|part\d*)[\\/]/i.test(normalized) ||
    /\.(png|jpe?g|gif|bmp|webp|tiff?|emf|wmf)(\s|$)/i.test(normalized) && /[\\/]/.test(normalized);
}

export function looksLikeImageFileName(value: string): boolean {
  return /^[\w .()\-]+?\.(png|jpe?g|gif|bmp|webp|tiff?|emf|wmf)$/i.test(value.trim());
}
