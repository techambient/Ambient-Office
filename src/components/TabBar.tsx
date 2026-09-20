import { useState, useRef, useEffect } from 'react';
import { 
  FileText, 
  Table, 
  Presentation, 
  X, 
  Plus, 
  Upload, 
  ChevronDown,
  Sparkles
} from 'lucide-react';
import { OfficeTab, OfficeDocType } from '../types/office';

interface TabBarProps {
  tabs: OfficeTab[];
  activeTabId: string;
  onSelectTab: (id: string) => void;
  onCloseTab: (id: string) => void;
  onNewTab: (type: OfficeDocType) => void;
  onUploadClick: () => void;
  onRenameTab: (id: string, newTitle: string) => void;
}

export const TabBar = ({
  tabs,
  activeTabId,
  onSelectTab,
  onCloseTab,
  onNewTab,
  onUploadClick,
  onRenameTab,
}: TabBarProps) => {
  const [showNewMenu, setShowNewMenu] = useState(false);
  const [editingTabId, setEditingTabId] = useState<string | null>(null);
  const [editingTitle, setEditingTitle] = useState('');
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setShowNewMenu(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const getTabIcon = (type: OfficeDocType, ext: string) => {
    switch (type) {
      case 'document':
        return <FileText className="w-4 h-4 text-blue-600 shrink-0" />;
      case 'spreadsheet':
        return <Table className="w-4 h-4 text-emerald-600 shrink-0" />;
      case 'presentation':
        return <Presentation className="w-4 h-4 text-amber-600 shrink-0" />;
    }
  };

  const getFormatBadge = (ext: string) => {
    const isAmbient = ext === 'aw' || ext === 'as' || ext === 'ap';
    return (
      <span className={`text-[10px] px-1.5 py-0.5 rounded font-mono font-medium ${
        isAmbient 
          ? 'bg-purple-100 text-purple-700 dark:bg-purple-900/40 dark:text-purple-300' 
          : 'bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-300'
      }`}>
        .{ext}
      </span>
    );
  };

  const startRenaming = (tab: OfficeTab, e: React.MouseEvent) => {
    e.stopPropagation();
    setEditingTabId(tab.id);
    setEditingTitle(tab.title);
  };

  const commitRenaming = () => {
    if (editingTabId && editingTitle.trim()) {
      onRenameTab(editingTabId, editingTitle.trim());
    }
    setEditingTabId(null);
  };

  return (
    <div id="ambient-tabbar" className="flex items-center bg-slate-100 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 px-2 pt-1 gap-1 select-none overflow-x-auto no-scrollbar">
      {/* Tab List */}
      <div className="flex items-center gap-1.5 min-w-0">
        {tabs.map((tab) => {
          const isActive = tab.id === activeTabId;
          return (
            <div
              key={tab.id}
              id={`tab-item-${tab.id}`}
              onClick={() => onSelectTab(tab.id)}
              onDoubleClick={(e) => startRenaming(tab, e)}
              className={`group relative flex items-center gap-2 px-3.5 py-2 text-xs font-medium rounded-t-lg transition-all border-t-2 cursor-pointer max-w-[240px] shrink-0 ${
                isActive
                  ? 'bg-white dark:bg-slate-950 text-slate-900 dark:text-slate-100 border-t-blue-600 shadow-sm'
                  : 'bg-slate-200/70 dark:bg-slate-800/60 text-slate-600 dark:text-slate-400 border-t-transparent hover:bg-slate-200 dark:hover:bg-slate-800 hover:text-slate-800 dark:hover:text-slate-200'
              }`}
            >
              {getTabIcon(tab.type, tab.extension)}

              {editingTabId === tab.id ? (
                <input
                  type="text"
                  value={editingTitle}
                  onChange={(e) => setEditingTitle(e.target.value)}
                  onBlur={commitRenaming}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') commitRenaming();
                    if (e.key === 'Escape') setEditingTabId(null);
                  }}
                  autoFocus
                  className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded text-xs outline-none border border-blue-500 w-28"
                  onClick={(e) => e.stopPropagation()}
                />
              ) : (
                <span className="truncate max-w-[120px] font-sans" title={`${tab.title} (.${tab.extension})`}>
                  {tab.title}
                </span>
              )}

              {getFormatBadge(tab.extension)}

              {tab.isModified && (
                <span className="w-1.5 h-1.5 rounded-full bg-blue-600 shrink-0" title="Unsaved changes" />
              )}

              {/* Close Tab Button */}
              {tabs.length > 1 && (
                <button
                  id={`btn-close-tab-${tab.id}`}
                  onClick={(e) => {
                    e.stopPropagation();
                    onCloseTab(tab.id);
                  }}
                  className="opacity-0 group-hover:opacity-100 p-0.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 transition-opacity ml-1"
                  title="Close tab"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          );
        })}
      </div>

      {/* New Tab Dropdown */}
      <div className="relative shrink-0" ref={menuRef}>
        <button
          id="btn-new-tab-menu"
          onClick={() => setShowNewMenu(!showNewMenu)}
          className="flex items-center justify-center w-8 h-8 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300 transition-colors"
          title="Create new document, sheet, or slide"
        >
          <Plus className="w-4 h-4" />
        </button>

        {showNewMenu && (
          <div className="absolute left-0 top-full mt-1 w-56 bg-white dark:bg-slate-900 rounded-xl shadow-xl border border-slate-200 dark:border-slate-800 py-1.5 z-50 animate-in fade-in zoom-in-95 duration-100">
            <div className="px-3 py-1 text-[11px] font-semibold text-slate-400 uppercase tracking-wider">
              Create New
            </div>
            
            <button
              id="btn-new-doc"
              onClick={() => {
                onNewTab('document');
                setShowNewMenu(false);
              }}
              className="w-full flex items-center gap-2.5 px-3 py-2 text-xs text-slate-700 dark:text-slate-200 hover:bg-blue-50 dark:hover:bg-blue-950/40 hover:text-blue-600 transition-colors"
            >
              <FileText className="w-4 h-4 text-blue-600" />
              <div className="text-left">
                <div className="font-medium">Word Document</div>
                <div className="text-[10px] text-slate-400">Microsoft .docx / .aw</div>
              </div>
            </button>

            <button
              id="btn-new-sheet"
              onClick={() => {
                onNewTab('spreadsheet');
                setShowNewMenu(false);
              }}
              className="w-full flex items-center gap-2.5 px-3 py-2 text-xs text-slate-700 dark:text-slate-200 hover:bg-emerald-50 dark:hover:bg-emerald-950/40 hover:text-emerald-600 transition-colors"
            >
              <Table className="w-4 h-4 text-emerald-600" />
              <div className="text-left">
                <div className="font-medium">Excel Spreadsheet</div>
                <div className="text-[10px] text-slate-400">Microsoft .xlsx / .as</div>
              </div>
            </button>

            <button
              id="btn-new-slides"
              onClick={() => {
                onNewTab('presentation');
                setShowNewMenu(false);
              }}
              className="w-full flex items-center gap-2.5 px-3 py-2 text-xs text-slate-700 dark:text-slate-200 hover:bg-amber-50 dark:hover:bg-amber-950/40 hover:text-amber-600 transition-colors"
            >
              <Presentation className="w-4 h-4 text-amber-600" />
              <div className="text-left">
                <div className="font-medium">PowerPoint Presentation</div>
                <div className="text-[10px] text-slate-400">Microsoft .pptx / .ap</div>
              </div>
            </button>
          </div>
        )}
      </div>

      {/* Quick Upload / Open File Button */}
      <div className="ml-auto flex items-center gap-2 pr-2 shrink-0">
        <button
          id="btn-quick-upload-tabbar"
          onClick={onUploadClick}
          className="flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white px-2.5 py-1 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 transition-colors font-medium"
          title="Open or Upload Document, Spreadsheet, or Presentation"
        >
          <Upload className="w-3.5 h-3.5 text-blue-600" />
          <span>Upload / Open</span>
        </button>
      </div>
    </div>
  );
};
