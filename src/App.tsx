import { useState, useRef, useEffect } from 'react';
import { 
  OfficeTab, 
  OfficeDocType, 
  DocumentData, 
  SpreadsheetData, 
  PresentationData 
} from './types/office';
import { defaultTabs, initialDocumentData, initialSpreadsheetData, initialPresentationData } from './sampleData';
import { 
  parseDocxFile, 
  parseAwFile, 
  parseExcelFile, 
  parseAsFile, 
  parsePptxFile, 
  parseApFile,
  downloadFile,
  exportToAw,
  exportToAs,
  exportToAp,
  exportDocxFromHtml,
  exportToExcel,
  exportToPptx
} from './utils/fileConverters';
import { Header } from './components/Header';
import { TabBar } from './components/TabBar';
import { DocumentEditor } from './components/document/DocumentEditor';
import { SpreadsheetEditor } from './components/spreadsheet/SpreadsheetEditor';
import { PresentationEditor } from './components/presentation/PresentationEditor';
import { SaveAsModal } from './components/SaveAsModal';
import { TemplatesModal } from './components/TemplatesModal';
import { Upload, FileUp, Sparkles } from 'lucide-react';

export default function App() {
  const [tabs, setTabs] = useState<OfficeTab[]>(defaultTabs);
  const [activeTabId, setActiveTabId] = useState<string>(defaultTabs[0].id);
  const [isSaveAsOpen, setIsSaveAsOpen] = useState(false);
  const [isTemplatesOpen, setIsTemplatesOpen] = useState(false);
  const [isDraggingOver, setIsDraggingOver] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const activeTab = tabs.find((t) => t.id === activeTabId) || tabs[0];

  // Helper to update active tab data
  const updateActiveTab = (updates: Partial<OfficeTab>) => {
    setTabs((prev) =>
      prev.map((tab) =>
        tab.id === activeTabId ? { ...tab, ...updates, isModified: true } : tab
      )
    );
  };

  // Switch tab
  const handleSelectTab = (id: string) => {
    setActiveTabId(id);
  };

  // Close tab
  const handleCloseTab = (id: string) => {
    if (tabs.length <= 1) return;
    const closedIndex = tabs.findIndex((t) => t.id === id);
    const newTabs = tabs.filter((t) => t.id !== id);
    setTabs(newTabs);
    if (activeTabId === id) {
      const nextTab = newTabs[Math.max(0, closedIndex - 1)];
      setActiveTabId(nextTab.id);
    }
  };

  // Rename tab
  const handleRenameTab = (id: string, newTitle: string) => {
    setTabs((prev) =>
      prev.map((t) => (t.id === id ? { ...t, title: newTitle, isModified: true } : t))
    );
  };

  // Create new tab
  const handleNewTab = (type: OfficeDocType) => {
    const newId = `tab-${Date.now()}`;
    let newTab: OfficeTab;

    if (type === 'document') {
      newTab = {
        id: newId,
        title: 'Untitled Document',
        type: 'document',
        extension: 'docx',
        isModified: false,
        documentData: {
          htmlContent: '<h1>Untitled Document</h1><p>Start drafting your notes or document here...</p>',
          margin: 'normal',
          paperSize: 'letter',
          orientation: 'portrait',
          zoom: 100,
        },
      };
    } else if (type === 'spreadsheet') {
      newTab = {
        id: newId,
        title: 'Untitled Spreadsheet',
        type: 'spreadsheet',
        extension: 'xlsx',
        isModified: false,
        spreadsheetData: {
          sheets: [
            {
              id: 'sheet-1',
              name: 'Sheet1',
              rowCount: 40,
              colCount: 20,
              data: {
                '0:0': { value: 'Item', bold: true },
                '0:1': { value: 'Amount', bold: true, align: 'right' },
              },
            },
          ],
          activeSheetId: 'sheet-1',
        },
      };
    } else {
      // Presentation
      newTab = {
        id: newId,
        title: 'Untitled Presentation',
        type: 'presentation',
        extension: 'pptx',
        isModified: false,
        presentationData: {
          slides: [
            {
              id: 'slide-1',
              title: 'Welcome',
              layout: 'title',
              background: '#ffffff',
              elements: [
                {
                  id: 'el-1',
                  type: 'heading',
                  x: 10,
                  y: 35,
                  width: 80,
                  height: 18,
                  content: 'Presentation Title',
                  fontSize: 36,
                  fontWeight: 'bold',
                  align: 'center',
                },
                {
                  id: 'el-2',
                  type: 'text',
                  x: 15,
                  y: 55,
                  width: 70,
                  height: 15,
                  content: 'Click here to add subtitle',
                  fontSize: 18,
                  color: '#64748b',
                  align: 'center',
                },
              ],
            },
          ],
          activeSlideId: 'slide-1',
          theme: 'Modern Slate',
          aspectRatio: '16:9',
        },
      };
    }

    setTabs((prev) => [...prev, newTab]);
    setActiveTabId(newId);
  };

  // Open file handler (Word, Excel, PowerPoint, .aw, .as, .ap)
  const handleProcessFile = async (file: File) => {
    const filename = file.name;
    const ext = filename.split('.').pop()?.toLowerCase() || '';
    const rawTitle = filename.substring(0, filename.lastIndexOf('.')) || filename;
    const newId = `tab-file-${Date.now()}`;

    try {
      if (ext === 'docx' || ext === 'doc') {
        const docData = await parseDocxFile(file);
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'document',
          extension: 'docx',
          isModified: false,
          documentData: docData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      } else if (ext === 'aw') {
        const text = await file.text();
        const docData = parseAwFile(text);
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'document',
          extension: 'aw',
          isModified: false,
          documentData: docData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      } else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
        const sheetData = await parseExcelFile(file);
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'spreadsheet',
          extension: ext === 'csv' ? 'csv' : 'xlsx',
          isModified: false,
          spreadsheetData: sheetData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      } else if (ext === 'as') {
        const text = await file.text();
        const sheetData = parseAsFile(text);
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'spreadsheet',
          extension: 'as',
          isModified: false,
          spreadsheetData: sheetData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      } else if (ext === 'pptx' || ext === 'ppt') {
        const presData = await parsePptxFile(file);
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'presentation',
          extension: 'pptx',
          isModified: false,
          presentationData: presData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      } else if (ext === 'ap') {
        const text = await file.text();
        const presData = parseApFile(text);
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'presentation',
          extension: 'ap',
          isModified: false,
          presentationData: presData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      } else {
        // Fallback text/markdown to document
        const text = await file.text();
        const docData: DocumentData = {
          htmlContent: `<p>${text.replace(/\n/g, '<br/>')}</p>`,
          margin: 'normal',
          paperSize: 'letter',
          orientation: 'portrait',
          zoom: 100,
        };
        const newTab: OfficeTab = {
          id: newId,
          title: rawTitle,
          type: 'document',
          extension: 'aw',
          isModified: false,
          documentData: docData,
        };
        setTabs((prev) => [...prev, newTab]);
        setActiveTabId(newId);
      }
    } catch (err) {
      console.error('Error opening file:', err);
    }
  };

  // Quick save (saves to current extension directly)
  const handleQuickSave = async () => {
    if (activeTab.type === 'document' && activeTab.documentData) {
      if (activeTab.extension === 'aw') {
        const content = exportToAw(activeTab.title, activeTab.documentData);
        downloadFile(content, `${activeTab.title}.aw`, 'application/javascript');
      } else {
        const blob = await exportDocxFromHtml(activeTab.title, activeTab.documentData.htmlContent);
        downloadFile(blob, `${activeTab.title}.docx`);
      }
    } else if (activeTab.type === 'spreadsheet' && activeTab.spreadsheetData) {
      if (activeTab.extension === 'as') {
        const content = exportToAs(activeTab.title, activeTab.spreadsheetData);
        downloadFile(content, `${activeTab.title}.as`, 'application/javascript');
      } else {
        const blob = exportToExcel(activeTab.spreadsheetData);
        downloadFile(blob, `${activeTab.title}.xlsx`);
      }
    } else if (activeTab.type === 'presentation' && activeTab.presentationData) {
      if (activeTab.extension === 'ap') {
        const content = exportToAp(activeTab.title, activeTab.presentationData);
        downloadFile(content, `${activeTab.title}.ap`, 'application/javascript');
      } else {
        await exportToPptx(activeTab.presentationData, activeTab.title);
      }
    }
    updateActiveTab({ isModified: false });
  };

  // Drag and drop listeners for whole window
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingOver(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      Array.from(files).forEach((file) => handleProcessFile(file));
    }
  };

  return (
    <div 
      id="ambient-office-root"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className="flex flex-col h-screen w-screen bg-slate-50 dark:bg-slate-950 overflow-hidden font-sans relative"
    >
      {/* Hidden File Input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={(e) => {
          const files = e.target.files;
          if (files && files.length > 0) {
            Array.from(files).forEach((f) => handleProcessFile(f));
          }
          if (fileInputRef.current) fileInputRef.current.value = '';
        }}
        accept=".docx,.doc,.xlsx,.xls,.pptx,.ppt,.aw,.as,.ap,.csv,.txt,.html"
        multiple
        className="hidden"
      />

      {/* Top Header & Menu Bar */}
      <Header
        activeTab={activeTab}
        onSaveAs={() => setIsSaveAsOpen(true)}
        onQuickSave={handleQuickSave}
        onUploadClick={() => fileInputRef.current?.click()}
        onNewDoc={(type) => handleNewTab(type)}
        onPrint={() => window.print()}
        onOpenSamples={() => setIsTemplatesOpen(true)}
        onUpdateTitle={(title) => updateActiveTab({ title })}
      />

      {/* Multi-Document Tab Bar */}
      <TabBar
        tabs={tabs}
        activeTabId={activeTabId}
        onSelectTab={handleSelectTab}
        onCloseTab={handleCloseTab}
        onNewTab={handleNewTab}
        onUploadClick={() => fileInputRef.current?.click()}
        onRenameTab={handleRenameTab}
      />

      {/* Active Editor Canvas */}
      <main className="flex-1 overflow-hidden relative">
        {activeTab.type === 'document' && activeTab.documentData && (
          <DocumentEditor
            key={activeTab.id}
            data={activeTab.documentData}
            onChange={(documentData) => updateActiveTab({ documentData })}
          />
        )}

        {activeTab.type === 'spreadsheet' && activeTab.spreadsheetData && (
          <SpreadsheetEditor
            key={activeTab.id}
            data={activeTab.spreadsheetData}
            onChange={(spreadsheetData) => updateActiveTab({ spreadsheetData })}
          />
        )}

        {activeTab.type === 'presentation' && activeTab.presentationData && (
          <PresentationEditor
            key={activeTab.id}
            data={activeTab.presentationData}
            onChange={(presentationData) => updateActiveTab({ presentationData })}
          />
        )}
      </main>

      {/* Drag & Drop Visual Overlay */}
      {isDraggingOver && (
        <div className="fixed inset-0 z-50 bg-blue-600/20 backdrop-blur-xs flex items-center justify-center border-4 border-dashed border-blue-500 m-4 rounded-3xl pointer-events-none">
          <div className="bg-white dark:bg-slate-900 px-8 py-6 rounded-2xl shadow-2xl flex items-center gap-4 border border-blue-200">
            <FileUp className="w-10 h-10 text-blue-600 animate-bounce" />
            <div>
              <div className="text-base font-bold text-slate-900 dark:text-slate-100">
                Drop Office Files Here
              </div>
              <div className="text-xs text-slate-500">
                Supports Word (.docx), Excel (.xlsx), PowerPoint (.pptx), .aw, .as, and .ap
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Save As Modal Dialog */}
      <SaveAsModal
        tab={activeTab}
        isOpen={isSaveAsOpen}
        onClose={() => setIsSaveAsOpen(false)}
        onSaved={(newTitle, newExt) => {
          updateActiveTab({ title: newTitle, extension: newExt, isModified: false });
        }}
      />

      {/* Sample Templates Modal */}
      <TemplatesModal
        isOpen={isTemplatesOpen}
        onClose={() => setIsTemplatesOpen(false)}
        onLoadTemplate={(newTab) => {
          setTabs((prev) => [...prev, newTab]);
          setActiveTabId(newTab.id);
        }}
      />
    </div>
  );
}
