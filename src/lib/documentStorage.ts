import { supabase } from './supabaseClient';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const DOCUMENT_BUCKET = 'docucaption-documents';

export interface StoredDocumentRef {
  bucket: string;
  path: string;
  deleteAfterProcessing: boolean;
}

export async function uploadDocumentForProcessing(file: File): Promise<StoredDocumentRef | undefined> {
  if (!supabase) return undefined;

  const { data } = await supabase.auth.getSession();
  const userId = data.session?.user.id;
  if (!userId) return undefined;

  const deleteAfterProcessing = import.meta.env.VITE_PERSIST_UPLOADED_DOCS !== 'true';
  const path = `${userId}/uploads/${Date.now()}-${safeStorageFileName(file.name)}`;
  const { error } = await supabase.storage.from(DOCUMENT_BUCKET).upload(path, file, {
    cacheControl: '60',
    contentType: file.type || DOCX_MIME,
    upsert: false,
  });
  if (error) throw new Error(error.message);

  return {
    bucket: DOCUMENT_BUCKET,
    path,
    deleteAfterProcessing,
  };
}

export async function deleteStoredDocument(ref: StoredDocumentRef | undefined): Promise<void> {
  if (!supabase || !ref?.deleteAfterProcessing) return;
  const { error } = await supabase.storage.from(ref.bucket).remove([ref.path]);
  if (error) throw new Error(error.message);
}

function safeStorageFileName(fileName: string): string {
  return fileName
    .replace(/\.docx$/i, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80) || 'document';
}
