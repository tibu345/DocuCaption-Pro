import type { CaptionPlacement, CaptionSettings, DocumentAudit } from './captionUtils';
import { sanitizeCaptionDescription } from './captionUtils';
import type { UserProfile } from './auth';
import { getAccessToken, mapServerProfile, type ServerProfile } from './accountApi';
import { exportDocx } from './docxExporter';
import { parseDocxFile } from './docxParser';
import { friendlyErrorMessage } from './errorMessages';
import { apiFetch } from './apiClient';

export interface ProcessedDoc extends CaptionSettings {
  title: string;
  content: DocElement[];
  originalBuffer?: ArrayBuffer;
}

export interface ParsedDocument {
  elements: DocElement[];
  audit: DocumentAudit;
}

export type DocElement =
  | { type: 'paragraph'; text: string; style?: string }
  | {
      type: 'table';
      rows: string[][];
      caption?: string;
      id: string;
      order: number;
      excluded?: boolean;
      hasExistingCaption?: boolean;
    }
  | {
      type: 'image';
      src: string;
      caption?: string;
      id: string;
      order: number;
      alt?: string;
      excluded?: boolean;
      hasExistingCaption?: boolean;
    }
  | { type: 'toc'; id: string }
  | { type: 'list-figures'; id: string }
  | { type: 'list-tables'; id: string };

export type { CaptionPlacement, DocumentAudit };

export async function parseDocument(file: File): Promise<ParsedDocument> {
  return parseDocxFile(file);
}

export async function generateCaptions(elements: DocElement[], user: UserProfile): Promise<{ elements: DocElement[]; creditsUsed: number; warning?: string; user?: UserProfile }> {
  const targetElements = elements.filter(
    (element): element is Extract<DocElement, { type: 'table' | 'image' }> => (element.type === 'table' || element.type === 'image') && !element.excluded,
  );
  if (targetElements.length === 0) return { elements, creditsUsed: 0 };

  const requestElements = await Promise.all(
    targetElements.map(async (element) => {
      if (element.type === 'table') {
        return { id: element.id, type: element.type, rows: element.rows.slice(0, 5) };
      }
      return {
        id: element.id,
        type: element.type,
        alt: element.alt ?? '',
        imageDataUrl: await compressImageDataUrl(element.src),
      };
    }),
  );

  const accessToken = await getAccessToken();
  const response = await apiFetch('/api/generate-captions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...(accessToken ? { Authorization: `Bearer ${accessToken}` } : {}),
    },
    body: JSON.stringify({
      elements: requestElements,
    }),
  });

  if (!response.ok) {
    const message = await readErrorMessage(response);
    throw new Error(message || 'Generated caption service is not configured. Manual captions and fallback captions are still available.');
  }

  const result = (await response.json()) as { captions: { id: string; caption: string }[]; creditsUsed: number; warning?: string; profile?: ServerProfile };
  return {
    creditsUsed: result.creditsUsed,
    warning: result.warning,
    user: result.profile ? mapServerProfile(result.profile) : user,
    elements: elements.map((element) => {
      if (element.type !== 'table' && element.type !== 'image') return element;
      const match = result.captions.find((caption) => caption.id === element.id);
      return match ? { ...element, caption: normalizeCaptionWhitespace(match.caption) } : element;
    }),
  };
}

export function generateLocalTestingCaptions(elements: DocElement[]): { elements: DocElement[]; creditsUsed: number; warning: string; user?: UserProfile } {
  const targetElements = elements.filter(
    (element): element is Extract<DocElement, { type: 'table' | 'image' }> => (element.type === 'table' || element.type === 'image') && !element.excluded,
  );
  const captions = targetElements.map((element) => ({
    id: element.id,
    caption: normalizeCaptionWhitespace(localCaptionFor(element)),
  }));

  return {
    creditsUsed: captions.length,
    warning: 'Draft captions were added locally. Sign in and run generated captions when the document service is available.',
    elements: elements.map((element) => {
      if (element.type !== 'table' && element.type !== 'image') return element;
      const match = captions.find((caption) => caption.id === element.id);
      return match ? { ...element, caption: match.caption } : element;
    }),
  };
}

function normalizeCaptionWhitespace(caption: string): string {
  return caption.replace(/\s+/g, ' ').trim();
}

function localCaptionFor(element: Extract<DocElement, { type: 'table' | 'image' }>): string {
  const label = `${element.type === 'image' ? 'Figure' : 'Table'} ${element.order + 1}`;
  if (element.type === 'image') {
    return `${label}: ${sanitizeCaptionDescription(element.alt)}`;
  }
  const firstUsefulRow = element.rows.find((row) => row.some((cell) => cell.trim()));
  const summary = firstUsefulRow
    ?.map((cell) => cell.replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .slice(0, 4)
    .join(', ');
  return `${label}: ${summary ? `Summary of ${summary}` : 'Add description'}`;
}

async function readErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get('content-type') ?? '';
  if (contentType.includes('application/json')) {
    const payload = (await response.json()) as { error?: string };
    return friendlyErrorMessage(payload.error ?? response.statusText);
  }
  return friendlyErrorMessage(await response.text());
}

async function compressImageDataUrl(dataUrl: string): Promise<string | undefined> {
  if (!dataUrl.startsWith('data:image/')) return undefined;
  if (dataUrl.length <= 350_000) return dataUrl;

  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      const maxSide = 768;
      const scale = Math.min(1, maxSide / Math.max(image.width, image.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(image.width * scale));
      canvas.height = Math.max(1, Math.round(image.height * scale));
      const context = canvas.getContext('2d');
      if (!context) {
        resolve(undefined);
        return;
      }
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL('image/jpeg', 0.72));
    };
    image.onerror = () => resolve(undefined);
    image.src = dataUrl;
  });
}

export async function exportToDocxStructural(processed: ProcessedDoc, originalBuffer: ArrayBuffer): Promise<void> {
  return exportDocx(processed, originalBuffer);
}
