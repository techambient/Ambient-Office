import { useState, useRef, useEffect } from 'react';
import { 
  FileText, 
  Table, 
  Presentation, 
  FolderOpen, 
  Save, 
  Download, 
  Printer, 
  HelpCircle, 
  Sparkles, 
  Check, 
  ChevronDown,
  Layers,
  FileCheck,
  Eye,
  Undo2,
  Redo2,
  FileSpreadsheet
} from 'lucide-react';
import { OfficeTab, OfficeDocType } from '../types/office';

interface HeaderProps {
  activeTab: OfficeTab;
  onSaveAs: () => void;
  onQuickSave: () => void;
  onUploadClick: () => void;
  onNewDoc: (type: OfficeDocType) => void;
  onPrint: () => void;
  onOpenSamples: () => void;
  onUpdateTitle: (title: string) => void;
}

export const Header = ({
  activeTab,
  onSaveAs,
  onQuickSave,
  onUploadClick,
  onNewDoc,
  onPrint,
  onOpenSamples,
  onUpdateTitle,
}: HeaderProps) => {
  const [activeMenu, setActiveMenu] = useState<string | null>(null);
  const [isEditingTitle, setIsEditingTitle] = useState(false);
  const [titleInput, setTitleInput] = useState(activeTab.title);
  const navRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setTitleInput(activeTab.title);
  }, [activeTab.title]);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (navRef.current && !navRef.current.contains(e.target as Node)) {
        setActiveMenu(null);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, []);

  const handleTitleCommit = () => {
    if (titleInput.trim()) {
      onUpdateTitle(titleInput.trim());
    }
    setIsEditingTitle(false);
  };

  const getDocTypeIcon = () => {
    switch (activeTab.type) {
      case 'document':
        return <FileText className="w-5 h-5 text-blue-600" />;
      case 'spreadsheet':
        return <Table className="w-5 h-5 text-emerald-600" />;
      case 'presentation':
        return <Presentation className="w-5 h-5 text-amber-600" />;
    }
  };

  return (
    <header id="ambient-header" className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 select-none">
      {/* Top Banner Row */}
      <div className="flex items-center justify-between px-4 py-2 gap-4">
        {/* Logo and App Title */}
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-600 to-indigo-700 text-white flex items-center justify-center shadow-sm">
              <Layers className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-sm tracking-tight text-slate-900 dark:text-white flex items-center gap-1.5">
                Ambient Office
                <span className="text-[10px] px-1.5 py-0.2 font-mono bg-blue-50 text-blue-700 dark:bg-blue-950 dark:text-blue-300 rounded font-normal">
                  PRO
                </span>
              </span>
            </div>
          </div>

          <div className="h-4 w-px bg-slate-200 dark:border-slate-700 mx-1" />

          {/* Active File Title & Extension */}
          <div className="flex items-center gap-2">
            {getDocTypeIcon()}
            {isEditingTitle ? (
              <input
                id="input-header-title"
                type="text"
                value={titleInput}
                onChange={(e) => setTitleInput(e.target.value)}
                onBlur={handleTitleCommit}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleTitleCommit();
                  if (e.key === 'Escape') setIsEditingTitle(false);
                }}
                autoFocus
                className="text-sm font-semibold text-slate-900 dark:text-slate-100 bg-slate-100 dark:bg-slate-800 px-2 py-0.5 rounded border border-blue-500 outline-none"
              />
            ) : (
              <span
                id="header-tab-title-display"
                onClick={() => setIsEditingTitle(true)}
                className="text-sm font-semibold text-slate-800 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 px-2 py-0.5 rounded cursor-pointer transition-colors"
                title="Click to rename"
              >
                {activeTab.title}
                <span className="text-xs text-slate-400 font-mono font-normal ml-1">
                  .{activeTab.extension}
                </span>
              </span>
            )}

            {/* Saved Indicator */}
            <span className="text-[11px] text-slate-400 flex items-center gap-1">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
              {activeTab.isModified ? 'Editing' : 'Saved'}
            </span>
          </div>
        </div>

        {/* Action Controls */}
        <div className="flex items-center gap-2">
          <button
            id="btn-sample-templates"
            onClick={onOpenSamples}
            className="hidden md:flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-300 hover:text-blue-600 px-2.5 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors"
            title="Load sample Word, Excel, or PowerPoint files"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-500" />
            <span>Templates & Samples</span>
          </button>

          <button
            id="btn-header-upload"
            onClick={onUploadClick}
            className="flex items-center gap-1.5 text-xs text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 transition-colors font-medium cursor-pointer"
          >
            <FolderOpen className="w-4 h-4 text-blue-600" />
            <span>Open File</span>
          </button>

          <button
            id="btn-header-quicksave"
            onClick={onQuickSave}
            className="flex items-center gap-1.5 text-xs text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 transition-colors font-medium cursor-pointer"
            title="Save file"
          >
            <Save className="w-4 h-4 text-slate-600 dark:text-slate-300" />
            <span className="hidden sm:inline">Save</span>
          </button>

          <button
            id="btn-header-saveas"
            onClick={onSaveAs}
            className="flex items-center gap-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 px-3.5 py-1.5 rounded-lg shadow-sm transition-all cursor-pointer"
            title="Save As Microsoft Office (.docx/.xlsx/.pptx), PDF, or Ambient (.aw/.as/.ap)"
          >
            <Download className="w-4 h-4" />
            <span>Save As...</span>
          </button>
        </div>
      </div>

      {/* Menu Bar Row (File, Edit, Insert, Format, Tools, Help) */}
      <div 
        ref={navRef}
        className="flex items-center gap-1 px-4 py-0.5 border-t border-slate-100 dark:border-slate-800/80 text-xs text-slate-600 dark:text-slate-300"
      >
        {/* File Menu */}
        <div className="relative">
          <button
            id="menu-btn-file"
            onClick={() => setActiveMenu(activeMenu === 'file' ? null : 'file')}
            className={`px-2.5 py-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 font-medium ${
              activeMenu === 'file' ? 'bg-slate-100 dark:bg-slate-800 text-blue-600' : ''
            }`}
          >
            File
          </button>
          {activeMenu === 'file' && (
            <div className="absolute left-0 top-full mt-1 w-64 bg-white dark:bg-slate-900 rounded-xl shadow-xl border border-slate-200 dark:border-slate-800 py-1.5 z-50 animate-in fade-in zoom-in-95">
              <div className="px-3 py-1 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                Create New
              </div>
              <button
                onClick={() => { onNewDoc('document'); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 dark:hover:bg-blue-950/40 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <FileText className="w-4 h-4 text-blue-600" />
                  <span>New Document (.docx / .aw)</span>
                </div>
              </button>
              <button
                onClick={() => { onNewDoc('spreadsheet'); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-emerald-50 dark:hover:bg-emerald-950/40 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <Table className="w-4 h-4 text-emerald-600" />
                  <span>New Spreadsheet (.xlsx / .as)</span>
                </div>
              </button>
              <button
                onClick={() => { onNewDoc('presentation'); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-amber-50 dark:hover:bg-amber-950/40 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <Presentation className="w-4 h-4 text-amber-600" />
                  <span>New Presentation (.pptx / .ap)</span>
                </div>
              </button>

              <div className="my-1 border-t border-slate-100 dark:border-slate-800" />

              <button
                onClick={() => { onUploadClick(); setActiveMenu(null); }}
                className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <FolderOpen className="w-4 h-4 text-slate-500" />
                <span>Open / Upload File...</span>
              </button>

              <button
                onClick={() => { onQuickSave(); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <Save className="w-4 h-4 text-slate-500" />
                  <span>Save</span>
                </div>
                <span className="text-[10px] text-slate-400">Ctrl+S</span>
              </button>

              <button
                onClick={() => { onSaveAs(); setActiveMenu(null); }}
                className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left font-medium text-blue-600 dark:text-blue-400"
              >
                <Download className="w-4 h-4 text-blue-600" />
                <span>Save As (Office, PDF, .aw/.as/.ap)...</span>
              </button>

              <div className="my-1 border-t border-slate-100 dark:border-slate-800" />

              <button
                onClick={() => { onPrint(); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <Printer className="w-4 h-4 text-slate-500" />
                  <span>Print Document</span>
                </div>
                <span className="text-[10px] text-slate-400">Ctrl+P</span>
              </button>
            </div>
          )}
        </div>

        {/* Edit Menu */}
        <div className="relative">
          <button
            id="menu-btn-edit"
            onClick={() => setActiveMenu(activeMenu === 'edit' ? null : 'edit')}
            className={`px-2.5 py-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 font-medium ${
              activeMenu === 'edit' ? 'bg-slate-100 dark:bg-slate-800 text-blue-600' : ''
            }`}
          >
            Edit
          </button>
          {activeMenu === 'edit' && (
            <div className="absolute left-0 top-full mt-1 w-52 bg-white dark:bg-slate-900 rounded-xl shadow-xl border border-slate-200 dark:border-slate-800 py-1.5 z-50">
              <button
                onClick={() => { document.execCommand('undo'); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <Undo2 className="w-4 h-4 text-slate-500" />
                  <span>Undo</span>
                </div>
                <span className="text-[10px] text-slate-400">Ctrl+Z</span>
              </button>
              <button
                onClick={() => { document.execCommand('redo'); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <div className="flex items-center gap-2">
                  <Redo2 className="w-4 h-4 text-slate-500" />
                  <span>Redo</span>
                </div>
                <span className="text-[10px] text-slate-400">Ctrl+Y</span>
              </button>
              <div className="my-1 border-t border-slate-100 dark:border-slate-800" />
              <button
                onClick={() => { document.execCommand('selectAll'); setActiveMenu(null); }}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <span>Select All</span>
                <span className="text-[10px] text-slate-400">Ctrl+A</span>
              </button>
            </div>
          )}
        </div>

        {/* View Menu */}
        <div className="relative">
          <button
            id="menu-btn-view"
            onClick={() => setActiveMenu(activeMenu === 'view' ? null : 'view')}
            className={`px-2.5 py-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 font-medium ${
              activeMenu === 'view' ? 'bg-slate-100 dark:bg-slate-800 text-blue-600' : ''
            }`}
          >
            View
          </button>
          {activeMenu === 'view' && (
            <div className="absolute left-0 top-full mt-1 w-52 bg-white dark:bg-slate-900 rounded-xl shadow-xl border border-slate-200 dark:border-slate-800 py-1.5 z-50">
              <button
                onClick={() => { onPrint(); setActiveMenu(null); }}
                className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <Eye className="w-4 h-4 text-slate-500" />
                <span>Print Layout / Preview</span>
              </button>
              <button
                onClick={() => { onOpenSamples(); setActiveMenu(null); }}
                className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 text-left"
              >
                <Sparkles className="w-4 h-4 text-amber-500" />
                <span>Explore Sample Files</span>
              </button>
            </div>
          )}
        </div>

        {/* Help Menu */}
        <div className="relative">
          <button
            id="menu-btn-help"
            onClick={() => setActiveMenu(activeMenu === 'help' ? null : 'help')}
            className={`px-2.5 py-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 font-medium ${
              activeMenu === 'help' ? 'bg-slate-100 dark:bg-slate-800 text-blue-600' : ''
            }`}
          >
            Help
          </button>
          {activeMenu === 'help' && (
            <div className="absolute left-0 top-full mt-1 w-72 bg-white dark:bg-slate-900 rounded-xl shadow-xl border border-slate-200 dark:border-slate-800 p-3 z-50 space-y-2">
              <div className="font-semibold text-xs text-slate-900 dark:text-slate-100">
                Ambient Office Formats
              </div>
              <div className="text-[11px] text-slate-500 space-y-1">
                <div>• <strong className="text-blue-600">.aw</strong>: Ambient Word (JavaScript document)</div>
                <div>• <strong className="text-emerald-600">.as</strong>: Ambient Spreadsheet (JavaScript sheet)</div>
                <div>• <strong className="text-amber-600">.ap</strong>: Ambient Presentation (JavaScript slides)</div>
                <div>• Full support for <strong>.docx</strong>, <strong>.xlsx</strong>, <strong>.pptx</strong> and <strong>.pdf</strong>.</div>
              </div>
              <div className="pt-2 border-t border-slate-100 dark:border-slate-800 text-[10px] text-slate-400">
                Built with a lightweight browser-based architecture.
              </div>
            </div>
          )}
        </div>
      </div>
    </header>
  );
};
