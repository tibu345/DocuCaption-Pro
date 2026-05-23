import { type ChangeEvent, type Dispatch, type ReactNode, type RefObject, type SetStateAction, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowRight,
  CheckCircle2,
  Download,
  FileText,
  Image as ImageIcon,
  Loader2,
  Lock,
  ShieldCheck,
  Table as TableIcon,
  Trash2,
  Upload,
  User,
} from 'lucide-react';
import { exportToDocxStructural, generateCaptions, generateLocalTestingCaptions, parseDocument } from './lib/documentProcessor';
import type { DocElement, DocumentAudit, ProcessedDoc } from './lib/documentProcessor';
import {
  getRemainingCaptionCredits,
  getRemainingDocuments,
  getRemainingExports,
  isAuthenticated,
  loadUser,
  consumeLocalCaptionCredits,
  consumeLocalDocument,
  consumeLocalExport,
  trackProcessingFailure,
  type UserProfile,
} from './lib/auth';
import { recordDocumentProcessed, recordExport } from './lib/accountApi';
import { friendlyErrorMessage } from './lib/errorMessages';
import { canExportDocument, canProcessDocument, canUseGeneratedCaptions, estimateAuditCaptionCredits } from './lib/limits';
import { getPlan } from './lib/access';
import { continueWithGoogle, getCurrentAuthUser, onAuthStateChanged, signInWithEmail, signOut, signUpWithEmail } from './lib/supabaseAuth';

type Page = 'landing' | 'process' | 'audit' | 'export' | 'account' | 'privacy';
type AppSettings = Omit<ProcessedDoc, 'title' | 'content' | 'originalBuffer'>;

const defaultSettings: AppSettings = {
  toc: true,
  listFigures: true,
  listTables: true,
  figureCaptionPlacement: 'below',
  tableCaptionPlacement: 'above',
  typography: {
    font: 'Times New Roman',
    size: 11,
    alignment: 'center',
  },
};

export default function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [page, setPage] = useState<Page>('landing');
  const [user, setUser] = useState<UserProfile>(() => loadUser());
  const [file, setFile] = useState<File | null>(null);
  const [originalBuffer, setOriginalBuffer] = useState<ArrayBuffer | null>(null);
  const [elements, setElements] = useState<DocElement[]>([]);
  const [audit, setAudit] = useState<DocumentAudit | null>(null);
  const [settings, setSettings] = useState<AppSettings>(defaultSettings);
  const [authMode, setAuthMode] = useState<'signin' | 'signup'>('signup');
  const [pendingVerificationEmail, setPendingVerificationEmail] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isCaptioning, setIsCaptioning] = useState(false);
  const [isExportingDocx, setIsExportingDocx] = useState(false);
  const [showAuthPrompt, setShowAuthPrompt] = useState(false);
  const [message, setMessage] = useState<string | null>(null);

  const assets = useMemo(() => elements.filter((element) => element.type === 'image' || element.type === 'table'), [elements]);
  const activeAssets = assets.filter((element) => !element.excluded);
  const plan = getPlan(user.planId);
  const estimatedCaptionCredits = estimateAuditCaptionCredits(audit);
  const signedIn = isAuthenticated(user);

  useEffect(() => {
    getCurrentAuthUser()
      .then((currentUser) => {
        setUser(currentUser);
      })
      .catch((err: unknown) => setMessage(friendlyErrorMessage(err, 'Could not load authentication state.')));
    return onAuthStateChanged((currentUser) => {
      setUser(currentUser);
      if (isAuthenticated(currentUser)) setShowAuthPrompt(false);
    });
  }, []);

  useEffect(() => {
    if (signedIn) return;
    const timeout = window.setTimeout(() => setShowAuthPrompt(true), 2500);
    return () => window.clearTimeout(timeout);
  }, [signedIn]);

  async function handleFileUpload(event: ChangeEvent<HTMLInputElement>) {
    if (!signedIn) {
      setShowAuthPrompt(true);
      setMessage('Sign in to process Word documents.');
      if (event.target) event.target.value = '';
      return;
    }
    const selectedFile = event.target.files?.[0];
    if (!selectedFile) return;

    setMessage(null);
    if (!selectedFile.name.toLowerCase().endsWith('.docx')) {
      setMessage('Please choose a valid .docx file.');
      return;
    }
    if (getRemainingDocuments(user) <= 0) {
      setMessage('No document scans remaining this month.');
      return;
    }

    setFile(selectedFile);
    setIsProcessing(true);
    try {
      const buffer = await selectedFile.arrayBuffer();
      const parsed = await parseDocument(selectedFile);
      const assetCount = parsed.elements.filter((element) => element.type === 'image' || element.type === 'table').length;
      const limit = canProcessDocument(user, assetCount);
      if (!limit.allowed) {
        const updated = trackProcessingFailure(user);
        setUser(updated);
        setMessage(limit.message ?? 'This document exceeds the current processing limits.');
        return;
      }

      setOriginalBuffer(buffer);
      setElements(parsed.elements);
      setAudit(parsed.audit);
      setMessage(
        assetCount > 0
          ? `Document scanned. Found ${assetCount} figure/table item${assetCount === 1 ? '' : 's'} for review.`
          : 'Document scanned, but no figures or tables were detected. Try another .docx or check whether the visuals are linked images or embedded objects.',
      );
      try {
        setUser(await recordDocumentProcessed(selectedFile.name, assetCount));
      } catch (error) {
        setUser(consumeLocalDocument(user));
        console.warn(error);
      }
    } catch (err) {
      setUser(trackProcessingFailure(user));
      setMessage(friendlyErrorMessage(err, 'Failed to parse the document.'));
    } finally {
      setIsProcessing(false);
    }
  }

  async function handleAutoCaption() {
    if (!signedIn) {
      setShowAuthPrompt(true);
      setMessage('Sign in to generate captions.');
      return;
    }
    const creditsNeeded = activeAssets.length;
    const limit = canUseGeneratedCaptions(user, creditsNeeded);
    if (!limit.allowed) {
      setMessage(limit.message ?? 'Generated captions are not available for this account.');
      return;
    }

    setIsCaptioning(true);
    setMessage(null);
    try {
      const result = await generateCaptions(elements, user);
      setElements(result.elements);
      if (result.user) setUser(result.user);
      if (!result.user) setUser(consumeLocalCaptionCredits(user, result.creditsUsed));
      setMessage(result.warning ?? `Generated captions were staged for review. Edit any caption before exporting. ${result.creditsUsed} caption credits deducted.`);
    } catch (err) {
      const fallback = generateLocalTestingCaptions(elements);
      setElements(fallback.elements);
      setMessage(friendlyErrorMessage(err, 'Generated caption request failed. Draft captions were added for manual review.'));
    } finally {
      setIsCaptioning(false);
    }
  }

  function createProcessedDoc(): ProcessedDoc | undefined {
    if (!file) {
      setMessage('Import a .docx before exporting.');
      return undefined;
    }

    return {
      title: file.name.replace(/\.docx$/i, ''),
      content: elements,
      ...settings,
    };
  }

  async function handleDocxExport() {
    if (!signedIn) {
      setShowAuthPrompt(true);
      setMessage('Sign in to export corrected Word documents.');
      return;
    }
    if (!originalBuffer || !file) {
      setMessage('Import a .docx before exporting.');
      return;
    }

    const limit = canExportDocument(user);
    if (!limit.allowed) {
      setMessage(limit.message ?? 'No exports remaining.');
      return;
    }

    const processed = createProcessedDoc();
    if (!processed) return;

    setIsExportingDocx(true);
    setMessage(null);
    try {
      const updatedUser = await recordExport(file.name).catch((error) => {
        console.warn(error);
        setMessage('Export is being prepared.');
        return consumeLocalExport(user);
      });
      await exportToDocxStructural(processed, originalBuffer);
      setUser(updatedUser);
      setMessage('Export created as a new captioned Word copy. The original file was not overwritten.');
    } catch (err) {
      setUser(trackProcessingFailure(user));
      setMessage(friendlyErrorMessage(err, 'Export failed. The document structure may be unsupported.'));
    } finally {
      setIsExportingDocx(false);
    }
  }

  function handleSignIn() {
    setMessage(null);
    signInWithEmail(email, password)
      .then((updated) => {
        setUser(updated);
        setShowAuthPrompt(false);
        setPage('process');
        setMessage('Signed in successfully.');
      })
      .catch((err: unknown) => setMessage(friendlyErrorMessage(err, 'Could not sign in.')));
  }

  function handleSignUp() {
    setMessage(null);
    if (password !== confirmPassword) {
      setMessage('Passwords do not match.');
      return;
    }

    signUpWithEmail(email, password)
      .then((result) => {
        if (result.user) {
          setUser(result.user);
          setPendingVerificationEmail('');
          setShowAuthPrompt(false);
          setPage('process');
          setMessage('Account created and signed in.');
          return;
        }
        setPendingVerificationEmail(email.trim().toLowerCase());
        setAuthMode('signin');
        setMessage(null);
      })
      .catch((err: unknown) => setMessage(friendlyErrorMessage(err, 'Could not create account.')));
  }

  async function handleGoogleSignIn() {
    try {
      setMessage(null);
      await continueWithGoogle();
    } catch (err) {
      setMessage(friendlyErrorMessage(err, 'Could not start Google sign-in.'));
    }
  }

  async function handleSignOut() {
    setUser(await signOut());
    setPage('landing');
    setShowAuthPrompt(true);
    setMessage('Signed out.');
  }

  function updateCaption(id: string, caption: string) {
    setElements((current) =>
      current.map((element) => ((element.type === 'table' || element.type === 'image') && element.id === id ? { ...element, caption } : element)),
    );
  }

  function toggleExclusion(id: string) {
    setElements((current) =>
      current.map((element) => ((element.type === 'table' || element.type === 'image') && element.id === id ? { ...element, excluded: !element.excluded } : element)),
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200">
        <div className="mx-auto flex min-h-16 max-w-7xl flex-col gap-3 px-4 py-3 sm:px-6 md:flex-row md:items-center md:justify-between md:py-0">
          <button onClick={() => setPage('landing')} className="flex shrink-0 items-center gap-3">
            <span className="w-9 h-9 rounded-lg bg-blue-600 text-white flex items-center justify-center">
              <FileText className="w-5 h-5" />
            </span>
            <span className="text-lg font-bold">DocuCaption Pro</span>
          </button>
          <nav className="flex w-full items-center gap-1 overflow-x-auto pb-1 text-sm md:w-auto md:pb-0">
            {(['process', 'audit', 'export', 'account', 'privacy'] as Page[]).map((item) => (
              <button key={item} onClick={() => setPage(item)} className={`nav-pill ${page === item ? 'nav-pill-active' : ''}`}>
                {item}
              </button>
            ))}
          </nav>
        </div>
      </header>

      <main className="mx-auto max-w-7xl px-4 py-6 sm:px-6 sm:py-8">
        {message && <div className="mb-6 rounded-lg border border-blue-200 bg-blue-50 p-4 text-sm text-blue-900">{message}</div>}
        {page === 'landing' && <LandingPage signedIn={signedIn} captionCredits={getRemainingCaptionCredits(user)} onStart={() => signedIn ? setPage('process') : setShowAuthPrompt(true)} />}
        {page === 'process' && (
          <ProcessPage
            fileInputRef={fileInputRef}
            onUpload={handleFileUpload}
            isProcessing={isProcessing}
            isCaptioning={isCaptioning}
            assets={activeAssets}
            activeAssets={activeAssets}
            hasScannedDocument={Boolean(file)}
            canUseGeneratedCaptions={plan.generatedCaptionsEnabled}
            signedIn={signedIn}
            onAutoCaption={handleAutoCaption}
            onToggleExclusion={toggleExclusion}
            onUpdateCaption={updateCaption}
          />
        )}
        {page === 'audit' && <AuditPage audit={audit} estimatedCaptionCredits={estimatedCaptionCredits} />}
        {page === 'export' && (
          <ExportPage
            settings={settings}
            setSettings={setSettings}
            onExportDocx={handleDocxExport}
            isExportingDocx={isExportingDocx}
            hasDocument={Boolean(file)}
            signedIn={signedIn}
            remainingExports={getRemainingExports(user)}
          />
        )}
        {page === 'account' && (
          <AccountPage
            user={user}
            authMode={authMode}
            pendingVerificationEmail={pendingVerificationEmail}
            email={email}
            password={password}
            confirmPassword={confirmPassword}
            setAuthMode={setAuthMode}
            setPendingVerificationEmail={setPendingVerificationEmail}
            setEmail={setEmail}
            setPassword={setPassword}
            setConfirmPassword={setConfirmPassword}
            onSignIn={handleSignIn}
            onSignUp={handleSignUp}
            onGoogleSignIn={handleGoogleSignIn}
            onSignOut={handleSignOut}
          />
        )}
        {page === 'privacy' && <PrivacyPage />}
      </main>
      {showAuthPrompt && !signedIn && (
        <AuthPromptModal onClose={() => setShowAuthPrompt(false)}>
          <AccountPage
            user={user}
            authMode={authMode}
            pendingVerificationEmail={pendingVerificationEmail}
            email={email}
            password={password}
            confirmPassword={confirmPassword}
            setAuthMode={setAuthMode}
            setPendingVerificationEmail={setPendingVerificationEmail}
            setEmail={setEmail}
            setPassword={setPassword}
            setConfirmPassword={setConfirmPassword}
            onSignIn={handleSignIn}
            onSignUp={handleSignUp}
            onGoogleSignIn={handleGoogleSignIn}
            onSignOut={handleSignOut}
          />
        </AuthPromptModal>
      )}
    </div>
  );
}

function LandingPage({ signedIn, captionCredits, onStart }: { signedIn: boolean; captionCredits: number; onStart: () => void }) {
  return (
    <section className="space-y-10">
      <div className="grid min-h-[calc(100vh-12rem)] items-center gap-8 lg:grid-cols-[1.05fr_0.95fr]">
        <div>
          <p className="text-sm font-bold text-blue-600 uppercase tracking-widest">Document automation for Word</p>
          <h1 className="mt-4 max-w-4xl text-4xl font-bold leading-tight sm:text-5xl lg:text-6xl">Upload a report. Generate captions for every figure and table. Export a clean Word file.</h1>
          <p className="mt-6 max-w-2xl text-base text-slate-600 sm:text-lg">
            DocuCaption fixes captions, Tables of Figures, Tables of Tables, and report structure before submission.
          </p>
          <div className="mt-8 flex flex-wrap gap-3">
            <button onClick={onStart} className="primary-button">
              <Upload className="w-4 h-4" /> {signedIn ? 'Start processing' : 'Start free'}
            </button>
          </div>
          <div className="mt-6 flex flex-wrap gap-3 text-sm">
            <span className="badge">{signedIn ? `${captionCredits} caption credits left` : 'Free account required'}</span>
            <span className="badge">Word-native SEQ fields</span>
            <span className="badge">No PDF downgrade</span>
          </div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
          <div className="rounded-lg bg-blue-50 border border-blue-100 p-5">
            <div className="text-xs font-bold uppercase tracking-wider text-blue-700">Automation sequence</div>
            <div className="mt-4 grid gap-3">
              {['Detect figures and tables', 'Generate captions', 'Prevent duplicate captions', 'Build TOC + figure/table lists', 'Export corrected DOCX'].map((item, index) => (
                <div key={item} className="flex min-w-0 items-center gap-3 rounded border border-blue-100 bg-white p-3">
                  <span className="w-7 h-7 rounded-full bg-blue-600 text-white flex items-center justify-center text-xs font-bold">{index + 1}</span>
                  <span className="min-w-0 text-sm font-semibold">{item}</span>
                  {index < 4 && <ArrowRight className="ml-auto w-4 h-4 text-blue-300" />}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
      <div className="grid gap-4 md:grid-cols-3">
        <ValueCard title="Automated First" value="Generated captions" detail="The full workflow is available without a paid plan." />
        <ValueCard title="Submission Ready" value="DOCX fields" detail="Captions and lists are Word-native, not pasted text." />
        <ValueCard title="Free Tool" value="$0" detail="Use document detection, caption generation, and Word export at no cost." />
      </div>
    </section>
  );
}

function ValueCard({ title, value, detail }: { title: string; value: string; detail: string }) {
  return (
    <div className="bg-white border border-slate-200 rounded-lg p-5 shadow-sm">
      <div className="text-xs font-bold uppercase tracking-wider text-slate-400">{title}</div>
      <div className="mt-2 text-2xl font-bold">{value}</div>
      <p className="mt-2 text-sm text-slate-600">{detail}</p>
    </div>
  );
}

function CheckIcon() {
  return <CheckCircle2 className="h-5 w-5" />;
}

function AuthPromptModal({ children, onClose }: { children: ReactNode; onClose: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-slate-950/40 px-3 py-4 backdrop-blur-sm sm:px-4 sm:py-8" role="dialog" aria-modal="true" aria-label="Account sign in">
      <div className="relative w-full max-w-5xl rounded-2xl bg-slate-50 p-3 shadow-2xl ring-1 ring-white/30 sm:p-5">
        <button
          type="button"
          onClick={onClose}
          className="absolute right-4 top-4 z-10 rounded-md border border-slate-200 bg-white px-3 py-2 text-xs font-bold text-slate-500 shadow-sm hover:text-slate-900"
          aria-label="Continue without signing in for now"
        >
          Later
        </button>
        <div className="mb-4 overflow-hidden rounded-xl border border-blue-100 bg-white">
          <div className="flex flex-col gap-4 bg-blue-50 px-4 py-4 pr-20 sm:flex-row sm:items-center sm:justify-between sm:px-5">
            <div>
              <p className="text-xs font-bold uppercase tracking-wider text-blue-700">Free account</p>
              <p className="mt-1 text-sm text-blue-950">Sign in once to process documents, generate captions, and export corrected Word files.</p>
            </div>
            <div className="hidden gap-2 text-[10px] font-bold uppercase tracking-wider text-emerald-700 md:flex">
              <span className="badge">Secure workspace</span>
              <span className="badge">Editable captions</span>
            </div>
          </div>
        </div>
        {children}
      </div>
    </div>
  );
}

function ProcessPage(props: {
  fileInputRef: RefObject<HTMLInputElement | null>;
  onUpload: (event: ChangeEvent<HTMLInputElement>) => void;
  isProcessing: boolean;
  isCaptioning: boolean;
  assets: Extract<DocElement, { type: 'image' | 'table' }>[];
  activeAssets: Extract<DocElement, { type: 'image' | 'table' }>[];
  hasScannedDocument: boolean;
  canUseGeneratedCaptions: boolean;
  signedIn: boolean;
  onAutoCaption: () => void;
  onToggleExclusion: (id: string) => void;
  onUpdateCaption: (id: string, caption: string) => void;
}) {
  return (
    <section className="grid gap-6 xl:grid-cols-[320px_minmax(0,1fr)]">
      <aside className="space-y-4 xl:sticky xl:top-24 xl:self-start">
        <SectionTitle title="Upload / Process" subtitle="Import a Word document, review detected assets, then generate or edit captions." />
        <input ref={props.fileInputRef} type="file" accept=".docx" onChange={props.onUpload} className="hidden" />
        <button onClick={() => props.fileInputRef.current?.click()} className="primary-button w-full justify-center" disabled={props.isProcessing}>
          {props.isProcessing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Upload className="w-4 h-4" />}
          Import .docx
        </button>
        <button onClick={props.onAutoCaption} disabled={props.activeAssets.length === 0 || props.isCaptioning} className="export-button">
          {props.isCaptioning ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldCheck className="w-4 h-4 text-blue-500" />}
          {props.canUseGeneratedCaptions ? 'Generate captions' : 'Generated captions are unavailable'}
        </button>
        {!props.signedIn && <p className="text-xs text-slate-500">Sign in to keep your work available across devices.</p>}
        <PrivacyNotice compact />
      </aside>
      <div className="min-w-0 space-y-5">
        {props.assets.length === 0 ? (
          <EmptyState
            label={
              props.hasScannedDocument
                ? 'No figures or tables were detected in this document.'
                : 'Upload a .docx to detect figures, tables, captions, and Word fields.'
            }
          />
        ) : (
          props.assets.map((element) => (
            <article key={element.id} className="overflow-hidden rounded-lg border border-slate-200 bg-white p-3 shadow-sm sm:p-4">
              <div className="flex items-start justify-between gap-3 sm:items-center">
                <div className="flex min-w-0 flex-wrap items-center gap-2 text-sm font-bold uppercase text-slate-500">
                  {element.type === 'image' ? <ImageIcon className="w-4 h-4" /> : <TableIcon className="w-4 h-4" />}
                  {element.type === 'image' ? `Figure ${element.order + 1}` : `Table ${element.order + 1}`}
                  {element.hasExistingCaption && <span className="badge">Existing caption</span>}
                </div>
                <button onClick={() => props.onToggleExclusion(element.id)} className="icon-button" title="Exclude from export">
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
              <AssetPreview element={element} />
              <input
                value={element.caption ?? ''}
                onChange={(event) => props.onUpdateCaption(element.id, event.target.value)}
                className="mt-4 w-full border-b border-dashed border-slate-300 focus:border-blue-600 focus:border-solid outline-none text-sm py-2 bg-transparent text-center font-serif italic"
                placeholder={element.type === 'image' ? 'Figure: Add description' : 'Table: Add description'}
              />
            </article>
          ))
        )}
      </div>
    </section>
  );
}

function AuditPage({ audit, estimatedCaptionCredits }: { audit: DocumentAudit | null; estimatedCaptionCredits: number }) {
  if (!audit) return <EmptyState label="No document audit yet. Upload a .docx first." />;
  const rows = [
    ['Total figures', audit.totalFigures],
    ['Figures with captions', audit.figuresWithCaptions],
    ['Figures missing captions', audit.figuresMissingCaptions],
    ['Total tables', audit.totalTables],
    ['Tables with captions', audit.tablesWithCaptions],
    ['Tables missing captions', audit.tablesMissingCaptions],
    ['TOC found', audit.hasToc ? 'Yes' : 'No'],
    ['Table of Figures found', audit.hasTableOfFigures ? 'Yes' : 'No'],
    ['Table of Tables found', audit.hasTableOfTables ? 'Yes' : 'No'],
    ['Estimated caption credits needed', estimatedCaptionCredits],
  ];
  return (
    <section className="space-y-6">
      <SectionTitle title="Document Audit" subtitle="Counts and field detection used for limits, credits, and export decisions." />
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-3">
        {rows.map(([label, value]) => (
          <div key={label} className="metric-card">
            <span>{label}</span>
            <strong>{value}</strong>
          </div>
        ))}
      </div>
    </section>
  );
}

function ExportPage(props: {
  settings: AppSettings;
  setSettings: Dispatch<SetStateAction<AppSettings>>;
  onExportDocx: () => void;
  isExportingDocx: boolean;
  hasDocument: boolean;
  signedIn: boolean;
  remainingExports: number;
}) {
  const exportDisabled = !props.hasDocument || props.remainingExports <= 0;

  return (
    <section className="grid gap-6 lg:grid-cols-[minmax(0,1fr)_360px]">
      <div className="min-w-0">
        <SectionTitle title="Export" subtitle="Create a corrected Word copy with Word-native captions and fields." />
        <div className="bg-white border border-slate-200 rounded-lg p-6 space-y-4">
          <Toggle label="Create/update Table of Contents" checked={props.settings.toc} onChange={(toc) => props.setSettings((current) => ({ ...current, toc }))} />
          <Toggle label="Create/update Table of Figures" checked={props.settings.listFigures} onChange={(listFigures) => props.setSettings((current) => ({ ...current, listFigures }))} />
          <Toggle label="Create/update Table of Tables" checked={props.settings.listTables} onChange={(listTables) => props.setSettings((current) => ({ ...current, listTables }))} />
          <div className="grid gap-3 sm:grid-cols-2">
            <SelectSetting label="Figure captions" value={props.settings.figureCaptionPlacement} options={[['below', 'Below'], ['above', 'Above']]} onChange={(value) => props.setSettings((current) => ({ ...current, figureCaptionPlacement: value }))} />
            <SelectSetting label="Table captions" value={props.settings.tableCaptionPlacement} options={[['above', 'Above'], ['below', 'Below']]} onChange={(value) => props.setSettings((current) => ({ ...current, tableCaptionPlacement: value }))} />
          </div>
          <button onClick={props.onExportDocx} disabled={exportDisabled || props.isExportingDocx} className="primary-button justify-center">
            {props.isExportingDocx ? <Loader2 className="w-4 h-4 animate-spin" /> : <Download className="w-4 h-4" />}
            Export captioned .docx
          </button>
          {!props.signedIn && <p className="text-xs text-slate-500">Sign in to keep your exports in your account history.</p>}
          {!props.hasDocument && <p className="text-xs text-slate-500">Upload a .docx before exporting.</p>}
          {props.hasDocument && props.remainingExports <= 0 && (
            <div className="rounded-md border border-red-200 bg-red-50 p-3 text-xs text-red-700">
              <p>You have reached your monthly export limit.</p>
            </div>
          )}
        </div>
      </div>
      <div className="metric-card">
        <span>Remaining exports</span>
        <strong>{props.remainingExports}</strong>
      </div>
    </section>
  );
}

function AccountPage(props: {
  user: UserProfile;
  authMode: 'signin' | 'signup';
  pendingVerificationEmail: string;
  email: string;
  password: string;
  confirmPassword: string;
  setAuthMode: (mode: 'signin' | 'signup') => void;
  setPendingVerificationEmail: (email: string) => void;
  setEmail: (email: string) => void;
  setPassword: (password: string) => void;
  setConfirmPassword: (password: string) => void;
  onSignIn: () => void;
  onSignUp: () => void;
  onGoogleSignIn: () => void;
  onSignOut: () => void;
}) {
  const plan = getPlan(props.user.planId);
  const signedIn = isAuthenticated(props.user);
  return (
    <section className="grid items-start gap-6 lg:grid-cols-[minmax(320px,420px)_minmax(0,1fr)]">
      <div className="auth-card">
        <div>
          <div className="w-11 h-11 rounded-lg bg-blue-600 text-white flex items-center justify-center mb-5">
            <User className="w-5 h-5" />
          </div>
          <h2 className="text-2xl font-bold">{signedIn ? 'Your account' : props.authMode === 'signup' ? 'Create your free account' : 'Sign in'}</h2>
          <p className="mt-2 text-sm text-slate-600">
            {signedIn
              ? 'Manage your document automation usage.'
              : props.authMode === 'signup'
                ? 'Set up your email and password to unlock document processing.'
                : 'Continue to your document workspace.'}
          </p>
        </div>

        {signedIn ? (
          <div className="space-y-4">
            <div className="rounded-lg border border-emerald-200 bg-emerald-50 p-4">
              <div className="text-xs font-bold uppercase tracking-wider text-emerald-700">Signed in</div>
              <div className="mt-1 text-sm font-semibold text-emerald-950">{props.user.email}</div>
            </div>
            <button onClick={props.onSignOut} className="toolbar-button justify-center w-full">
              Sign out
            </button>
          </div>
        ) : (
          <div className="space-y-4">
            {props.pendingVerificationEmail && (
              <div className="rounded-xl border border-blue-200 bg-blue-50 p-5 text-blue-950">
                <div className="flex items-start gap-3">
                  <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-blue-600 text-white">
                    <CheckIcon />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold">One quick step</h3>
                    <p className="mt-1 text-sm leading-6">
                      We sent a secure sign-in link to <strong>{props.pendingVerificationEmail}</strong>. Open it to finish setting up your account.
                    </p>
                    <div className="mt-3 flex flex-wrap gap-2">
                      <button
                        type="button"
                        onClick={() => props.setPendingVerificationEmail('')}
                        className="rounded-md border border-blue-200 bg-white px-3 py-2 text-xs font-semibold text-blue-800 hover:bg-blue-50"
                      >
                        Use a different email
                      </button>
                      <button
                        type="button"
                        onClick={() => props.setAuthMode('signin')}
                        className="rounded-md bg-blue-600 px-3 py-2 text-xs font-semibold text-white hover:bg-blue-700"
                      >
                        I verified, sign in
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}
            <button onClick={props.onGoogleSignIn} className="google-button">
              Continue with Google
            </button>
            <div className="auth-divider"><span>or use email</span></div>
            <div className="grid grid-cols-2 gap-2 rounded-lg bg-slate-100 p-1">
              <button
                onClick={() => props.setAuthMode('signup')}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${props.authMode === 'signup' ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-500'}`}
              >
                Sign up
              </button>
              <button
                onClick={() => props.setAuthMode('signin')}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${props.authMode === 'signin' ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-500'}`}
              >
                Sign in
              </button>
            </div>
            <label className="form-label">
              Email
              <input value={props.email} onChange={(event) => props.setEmail(event.target.value)} className="form-control auth-input" autoComplete="email" />
            </label>
            <label className="form-label">
              Password
              <input
                type="password"
                value={props.password}
                onChange={(event) => props.setPassword(event.target.value)}
                className="form-control auth-input"
                autoComplete={props.authMode === 'signup' ? 'new-password' : 'current-password'}
              />
            </label>
            {props.authMode === 'signup' && (
              <label className="form-label">
                Confirm password
                <input
                  type="password"
                  value={props.confirmPassword}
                  onChange={(event) => props.setConfirmPassword(event.target.value)}
                  className="form-control auth-input"
                  autoComplete="new-password"
                />
              </label>
            )}
            <button onClick={props.authMode === 'signup' ? props.onSignUp : props.onSignIn} className="primary-button justify-center w-full py-3">
              <User className="w-4 h-4" /> {props.authMode === 'signup' ? 'Create account' : 'Sign in'}
            </button>
            <p className="text-xs text-slate-500 text-center">No payment required.</p>
          </div>
        )}
      </div>

      {signedIn ? (
        <div className="grid gap-4 sm:grid-cols-2">
          <div className="metric-card">
            <span>Access</span>
            <strong>{plan.name}</strong>
          </div>
          <div className="metric-card">
            <span>Documents remaining</span>
            <strong>{getRemainingDocuments(props.user)}</strong>
          </div>
          <div className="metric-card">
            <span>Generated captions remaining</span>
            <strong>{getRemainingCaptionCredits(props.user)}</strong>
          </div>
          <div className="metric-card">
            <span>Exports remaining</span>
            <strong>{getRemainingExports(props.user)}</strong>
          </div>
        </div>
      ) : (
        <div className="bg-white border border-slate-200 rounded-lg p-6 shadow-sm">
          <p className="text-sm font-bold text-blue-600 uppercase tracking-widest">Document automation for Word</p>
          <h1 className="mt-4 text-3xl font-bold leading-tight sm:text-4xl">Sign in to start fixing Word reports.</h1>
          <p className="mt-4 text-slate-600">
            Your account keeps document usage, caption generation, and exports connected securely across sessions.
          </p>
          <div className="mt-6 grid gap-3">
            {['Detect figures and tables', 'Generate editable captions', 'Export corrected DOCX'].map((item, index) => (
              <div key={item} className="flex items-center gap-3 rounded bg-blue-50 border border-blue-100 p-3">
                <span className="w-7 h-7 rounded-full bg-blue-600 text-white flex items-center justify-center text-xs font-bold">{index + 1}</span>
                <span className="text-sm font-semibold">{item}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </section>
  );
}

function PrivacyPage() {
  return (
    <section className="space-y-6">
      <SectionTitle title="Privacy" subtitle="Designed for sensitive academic and project documents." />
      <PrivacyNotice />
      <div className="grid gap-4 md:grid-cols-3">
        {['Documents are processed in the browser for parsing/export.', 'Files are not stored permanently by this MVP.', 'Generated captions use only minimal figure/table context, not the full document.'].map((item) => (
          <div key={item} className="bg-white border border-slate-200 rounded-lg p-5 text-sm text-slate-600">
            <ShieldCheck className="w-5 h-5 text-emerald-600 mb-3" />
            {item}
          </div>
        ))}
      </div>
    </section>
  );
}

function AssetPreview({ element }: { element: Extract<DocElement, { type: 'image' | 'table' }> }) {
  if (element.type === 'image') {
    return (
      <div className="mt-4 flex justify-center rounded border border-slate-100 bg-slate-50 p-3 sm:p-4">
        {element.src ? <img src={element.src} alt={`Figure ${element.order + 1}`} className="max-h-[55vh] max-w-full object-contain" /> : <span className="text-sm text-slate-400">Image preview unavailable</span>}
      </div>
    );
  }
  return (
    <div className="mt-4 overflow-x-auto rounded border border-slate-100 bg-slate-50">
      <table className="min-w-full border-collapse text-xs">
        <tbody>
          {element.rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, cellIndex) => (
                <td key={cellIndex} className="border border-slate-200 p-2">
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function PrivacyNotice({ compact = false }: { compact?: boolean }) {
  return (
    <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-4 text-sm text-emerald-900">
      <div className="flex items-center gap-2 font-bold">
        <Lock className="w-4 h-4" />
        Privacy-first processing
      </div>
      {!compact && <p className="mt-2">Files are deleted after processing. Use manual/default captions any time. Generated captions use only minimal figure/table context.</p>}
      {compact && <p className="mt-2">Files are deleted after processing. Generated captions are optional.</p>}
    </div>
  );
}

function SectionTitle({ title, subtitle }: { title: string; subtitle: string }) {
  return (
    <div>
      <h2 className="text-2xl font-bold capitalize">{title}</h2>
      <p className="mt-1 text-sm text-slate-500">{subtitle}</p>
    </div>
  );
}

function EmptyState({ label = 'Upload a .docx to detect figures, tables, captions, and Word fields.' }: { label?: string }) {
  return (
    <div className="rounded-lg border border-dashed border-slate-300 bg-white p-8 text-center text-slate-500 sm:p-12">
      <FileText className="w-10 h-10 mx-auto mb-3 text-slate-300" />
      {label}
    </div>
  );
}

function Toggle({ label, checked, onChange }: { label: string; checked: boolean; onChange: (checked: boolean) => void }) {
  return (
    <label className="flex items-center justify-between text-sm text-slate-600">
      {label}
      <input type="checkbox" checked={checked} onChange={(event) => onChange(event.target.checked)} className="accent-blue-600" />
    </label>
  );
}

function SelectSetting<T extends string>({ label, value, options, onChange }: { label: string; value: T; options: [T, string][]; onChange: (value: T) => void }) {
  return (
    <label className="form-label">
      {label}
      <select value={value} onChange={(event) => onChange(event.target.value as T)} className="form-control">
        {options.map(([optionValue, text]) => (
          <option key={optionValue} value={optionValue}>
            {text}
          </option>
        ))}
      </select>
    </label>
  );
}
