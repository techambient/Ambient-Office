import { useState, useRef, useEffect, useCallback } from 'react';
import { 
  Bold, 
  Italic, 
  Underline, 
  Strikethrough, 
  AlignLeft, 
  AlignCenter, 
  AlignRight, 
  AlignJustify, 
  List, 
  ListOrdered, 
  Table as TableIcon, 
  Image as ImageIcon, 
  Link as LinkIcon, 
  Highlighter, 
  Type, 
  Heading1, 
  Heading2, 
  Heading3, 
  Quote, 
  Minus, 
  RotateCcw, 
  RotateCw, 
  ZoomIn, 
  ZoomOut, 
  FileText,
  Calendar,
  CheckSquare,
  Sparkles,
  Maximize2
} from 'lucide-react';
import { DocumentData } from '../../types/office';

interface DocumentEditorProps {
  data: DocumentData;
  onChange: (newData: DocumentData) => void;
}

export const DocumentEditor = ({ data, onChange }: DocumentEditorProps) => {
  const editorRef = useRef<HTMLDivElement>(null);
  const [activeTab, setActiveTab] = useState<'home' | 'insert' | 'layout' | 'view'>('home');
  const [wordCount, setWordCount] = useState(0);
  const [charCount, setCharCount] = useState(0);
  const [zoom, setZoom] = useState(data.zoom || 100);
  const [margin, setMargin] = useState(data.margin || 'normal');
  const [paperSize, setPaperSize] = useState(data.paperSize || 'letter');
  const [textColor, setTextColor] = useState('#1e293b');
  const [highlightColor, setHighlightColor] = useState('#fef08a');
  const [isTableModalOpen, setIsTableModalOpen] = useState(false);
  const [tableRows, setTableRows] = useState(3);
  const [tableCols, setTableCols] = useState(3);

  // Initialize editor content once
  useEffect(() => {
    if (editorRef.current && data.htmlContent) {
      if (editorRef.current.innerHTML !== data.htmlContent) {
        editorRef.current.innerHTML = data.htmlContent;
      }
      updateStats();
    }
  }, []);

  const updateStats = () => {
    if (editorRef.current) {
      const text = editorRef.current.innerText || '';
      const words = text.trim() ? text.trim().split(/\s+/).length : 0;
      setWordCount(words);
      setCharCount(text.length);
    }
  };

  const handleInput = () => {
    if (editorRef.current) {
      const newHtml = editorRef.current.innerHTML;
      updateStats();
      onChange({
        ...data,
        htmlContent: newHtml,
        margin,
        paperSize,
        zoom,
      });
    }
  };

  const exec = (command: string, value: string = '') => {
    editorRef.current?.focus();
    document.execCommand(command, false, value);
    handleInput();
  };

  const handleInsertTable = () => {
    let tableHtml = '<table style="width: 100%; border-collapse: collapse; margin: 16px 0;"><tbody>';
    for (let r = 0; r < tableRows; r++) {
      tableHtml += '<tr>';
      for (let c = 0; c < tableCols; c++) {
        tableHtml += `<td style="border: 1px solid #cbd5e1; padding: 8px 12px; min-width: 60px;">${r === 0 ? `<strong>Header ${c + 1}</strong>` : 'Data'}</td>`;
      }
      tableHtml += '</tr>';
    }
    tableHtml += '</tbody></table><p><br/></p>';
    exec('insertHTML', tableHtml);
    setIsTableModalOpen(false);
  };

  const handleInsertImage = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.onchange = (e: any) => {
      const file = e.target?.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (loadEvt) => {
          const base64 = loadEvt.target?.result as string;
          exec('insertHTML', `<img src="${base64}" style="max-width: 100%; height: auto; border-radius: 8px; margin: 12px 0;" /><p><br/></p>`);
        };
        reader.readAsDataURL(file);
      }
    };
    input.click();
  };

  const getMarginClass = () => {
    switch (margin) {
      case 'narrow':
        return 'p-8 sm:p-10';
      case 'wide':
        return 'p-16 sm:p-24';
      default:
        return 'p-12 sm:p-16';
    }
  };

  return (
    <div id="ambient-word-editor" className="flex flex-col h-full bg-slate-100 dark:bg-slate-950 overflow-hidden select-none">
      {/* Ribbon Navigation Bar */}
      <div className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 shadow-xs z-10">
        {/* Ribbon Tabs Header */}
        <div className="flex items-center gap-1 px-4 pt-1.5 border-b border-slate-100 dark:border-slate-800/80">
          {(['home', 'insert', 'layout', 'view'] as const).map((tab) => (
            <button
              key={tab}
              id={`ribbon-tab-${tab}`}
              onClick={() => setActiveTab(tab)}
              className={`px-4 py-1 text-xs font-semibold uppercase tracking-wider transition-colors border-b-2 ${
                activeTab === tab
                  ? 'border-blue-600 text-blue-600 dark:text-blue-400'
                  : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
              }`}
            >
              {tab}
            </button>
          ))}
        </div>

        {/* Ribbon Tools Container */}
        <div className="px-4 py-2 flex items-center gap-2 overflow-x-auto no-scrollbar">
          {activeTab === 'home' && (
            <>
              {/* History */}
              <div className="flex items-center gap-0.5 border-r border-slate-200 dark:border-slate-700 pr-2">
                <button
                  id="btn-doc-undo"
                  onClick={() => exec('undo')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Undo (Ctrl+Z)"
                >
                  <RotateCcw className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-redo"
                  onClick={() => exec('redo')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Redo (Ctrl+Y)"
                >
                  <RotateCw className="w-4 h-4" />
                </button>
              </div>

              {/* Font Family & Size */}
              <div className="flex items-center gap-1.5 border-r border-slate-200 dark:border-slate-700 pr-2">
                <select
                  id="select-font-family"
                  onChange={(e) => exec('fontName', e.target.value)}
                  defaultValue="Arial"
                  className="text-xs bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded px-2 py-1 text-slate-800 dark:text-slate-200 outline-none"
                >
                  <option value="Arial">Arial</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Courier New">Courier New</option>
                  <option value="Verdana">Verdana</option>
                  <option value="Trebuchet MS">Trebuchet MS</option>
                </select>

                <select
                  id="select-font-size"
                  onChange={(e) => exec('fontSize', e.target.value)}
                  defaultValue="3"
                  className="text-xs bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded px-2 py-1 text-slate-800 dark:text-slate-200 outline-none"
                >
                  <option value="1">10 pt</option>
                  <option value="2">12 pt</option>
                  <option value="3">14 pt (Normal)</option>
                  <option value="4">18 pt</option>
                  <option value="5">24 pt (Heading)</option>
                  <option value="6">32 pt (Title)</option>
                  <option value="7">48 pt</option>
                </select>
              </div>

              {/* Text Styling */}
              <div className="flex items-center gap-0.5 border-r border-slate-200 dark:border-slate-700 pr-2">
                <button
                  id="btn-doc-bold"
                  onClick={() => exec('bold')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 font-bold"
                  title="Bold (Ctrl+B)"
                >
                  <Bold className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-italic"
                  onClick={() => exec('italic')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 italic"
                  title="Italic (Ctrl+I)"
                >
                  <Italic className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-underline"
                  onClick={() => exec('underline')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 underline"
                  title="Underline (Ctrl+U)"
                >
                  <Underline className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-strikethrough"
                  onClick={() => exec('strikeThrough')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Strikethrough"
                >
                  <Strikethrough className="w-4 h-4" />
                </button>
              </div>

              {/* Colors */}
              <div className="flex items-center gap-1 border-r border-slate-200 dark:border-slate-700 pr-2">
                <label className="flex items-center gap-1 p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer" title="Font Color">
                  <Type className="w-4 h-4 text-slate-700 dark:text-slate-200" />
                  <input
                    type="color"
                    value={textColor}
                    onChange={(e) => {
                      setTextColor(e.target.value);
                      exec('foreColor', e.target.value);
                    }}
                    className="w-4 h-4 rounded border-0 cursor-pointer p-0 bg-transparent"
                  />
                </label>

                <label className="flex items-center gap-1 p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer" title="Highlight Color">
                  <Highlighter className="w-4 h-4 text-amber-500" />
                  <input
                    type="color"
                    value={highlightColor}
                    onChange={(e) => {
                      setHighlightColor(e.target.value);
                      exec('hiliteColor', e.target.value);
                    }}
                    className="w-4 h-4 rounded border-0 cursor-pointer p-0 bg-transparent"
                  />
                </label>
              </div>

              {/* Headings */}
              <div className="flex items-center gap-0.5 border-r border-slate-200 dark:border-slate-700 pr-2">
                <button
                  id="btn-doc-h1"
                  onClick={() => exec('formatBlock', '<h1>')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Heading 1"
                >
                  <Heading1 className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-h2"
                  onClick={() => exec('formatBlock', '<h2>')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Heading 2"
                >
                  <Heading2 className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-h3"
                  onClick={() => exec('formatBlock', '<h3>')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Heading 3"
                >
                  <Heading3 className="w-4 h-4" />
                </button>
                <button
                  id="btn-doc-p"
                  onClick={() => exec('formatBlock', '<p>')}
                  className="px-2 py-1 text-xs rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 font-medium"
                  title="Normal Paragraph"
                >
                  Normal
                </button>
              </div>

              {/* Alignment */}
              <div className="flex items-center gap-0.5 border-r border-slate-200 dark:border-slate-700 pr-2">
                <button
                  onClick={() => exec('justifyLeft')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Align Left"
                >
                  <AlignLeft className="w-4 h-4" />
                </button>
                <button
                  onClick={() => exec('justifyCenter')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Align Center"
                >
                  <AlignCenter className="w-4 h-4" />
                </button>
                <button
                  onClick={() => exec('justifyRight')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Align Right"
                >
                  <AlignRight className="w-4 h-4" />
                </button>
                <button
                  onClick={() => exec('justifyFull')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Justify"
                >
                  <AlignJustify className="w-4 h-4" />
                </button>
              </div>

              {/* Lists */}
              <div className="flex items-center gap-0.5">
                <button
                  onClick={() => exec('insertUnorderedList')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Bullet List"
                >
                  <List className="w-4 h-4" />
                </button>
                <button
                  onClick={() => exec('insertOrderedList')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Numbered List"
                >
                  <ListOrdered className="w-4 h-4" />
                </button>
                <button
                  onClick={() => exec('formatBlock', '<blockquote>')}
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
                  title="Quote Block"
                >
                  <Quote className="w-4 h-4" />
                </button>
              </div>
            </>
          )}

          {activeTab === 'insert' && (
            <div className="flex items-center gap-3">
              {/* Insert Table */}
              <button
                id="btn-insert-table"
                onClick={() => setIsTableModalOpen(true)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg border border-slate-200 dark:border-slate-700"
              >
                <TableIcon className="w-4 h-4 text-blue-600" />
                <span>Insert Table</span>
              </button>

              {/* Insert Image */}
              <button
                id="btn-insert-image"
                onClick={handleInsertImage}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg border border-slate-200 dark:border-slate-700"
              >
                <ImageIcon className="w-4 h-4 text-emerald-600" />
                <span>Insert Image</span>
              </button>

              {/* Horizontal Line */}
              <button
                onClick={() => exec('insertHorizontalRule')}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg border border-slate-200 dark:border-slate-700"
              >
                <Minus className="w-4 h-4 text-slate-500" />
                <span>Divider Line</span>
              </button>

              {/* Insert Date */}
              <button
                onClick={() => exec('insertHTML', `<span>${new Date().toLocaleDateString(undefined, { dateStyle: 'long' })}</span>`)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg border border-slate-200 dark:border-slate-700"
              >
                <Calendar className="w-4 h-4 text-purple-600" />
                <span>Insert Date</span>
              </button>
            </div>
          )}

          {activeTab === 'layout' && (
            <div className="flex items-center gap-4 text-xs">
              {/* Margins */}
              <div className="flex items-center gap-1.5">
                <span className="font-semibold text-slate-600 dark:text-slate-300">Margins:</span>
                {(['normal', 'narrow', 'wide'] as const).map((m) => (
                  <button
                    key={m}
                    onClick={() => {
                      setMargin(m);
                      onChange({ ...data, margin: m });
                    }}
                    className={`px-2.5 py-1 rounded capitalize ${
                      margin === m
                        ? 'bg-blue-600 text-white font-medium'
                        : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300'
                    }`}
                  >
                    {m}
                  </button>
                ))}
              </div>

              {/* Paper Size */}
              <div className="flex items-center gap-1.5 border-l border-slate-200 dark:border-slate-700 pl-3">
                <span className="font-semibold text-slate-600 dark:text-slate-300">Paper:</span>
                {(['letter', 'a4'] as const).map((p) => (
                  <button
                    key={p}
                    onClick={() => {
                      setPaperSize(p);
                      onChange({ ...data, paperSize: p });
                    }}
                    className={`px-2.5 py-1 rounded uppercase ${
                      paperSize === p
                        ? 'bg-blue-600 text-white font-medium'
                        : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300'
                    }`}
                  >
                    {p}
                  </button>
                ))}
              </div>
            </div>
          )}

          {activeTab === 'view' && (
            <div className="flex items-center gap-4 text-xs">
              <div className="flex items-center gap-2">
                <span className="font-semibold text-slate-600 dark:text-slate-300">Zoom:</span>
                <button
                  onClick={() => setZoom(Math.max(50, zoom - 10))}
                  className="p-1 rounded bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300"
                >
                  <ZoomOut className="w-3.5 h-3.5" />
                </button>
                <span className="font-mono text-xs w-10 text-center">{zoom}%</span>
                <button
                  onClick={() => setZoom(Math.min(200, zoom + 10))}
                  className="p-1 rounded bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300"
                >
                  <ZoomIn className="w-3.5 h-3.5" />
                </button>
                <button
                  onClick={() => setZoom(100)}
                  className="px-2 py-1 rounded bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 text-[11px]"
                >
                  Reset 100%
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Paginated Paper Canvas View */}
      <div className="flex-1 overflow-y-auto p-4 sm:p-8 flex justify-center items-start">
        <div 
          style={{ transform: `scale(${zoom / 100})`, transformOrigin: 'top center' }}
          className="transition-transform duration-100"
        >
          {/* Simulated Office Document Paper */}
          <div
            id="document-page-canvas"
            className={`w-[800px] min-h-[1050px] bg-white text-slate-900 shadow-xl border border-slate-200/80 rounded-sm relative ${getMarginClass()}`}
          >
            {/* Header watermark/ruler indicator */}
            <div className="absolute top-3 left-12 right-12 flex items-center justify-between text-[10px] text-slate-300 font-mono select-none pointer-events-none border-b border-slate-100 pb-1">
              <span>Ambient Word Document (.docx / .aw)</span>
              <span>Confidential & Proprietary</span>
            </div>

            {/* Editable Content Surface */}
            <div
              ref={editorRef}
              id="document-content-editable"
              contentEditable
              onInput={handleInput}
              onKeyUp={updateStats}
              className="outline-none min-h-[900px] prose prose-slate max-w-none focus:outline-none"
              style={{
                fontFamily: 'Arial, sans-serif',
                fontSize: '11pt',
                lineHeight: '1.6',
              }}
              suppressContentEditableWarning
            />

            {/* Page Footer */}
            <div className="absolute bottom-3 left-12 right-12 flex items-center justify-between text-[10px] text-slate-400 font-mono select-none pointer-events-none border-t border-slate-100 pt-1">
              <span>Page 1</span>
              <span>{wordCount} words</span>
            </div>
          </div>
        </div>
      </div>

      {/* Document Status Bar */}
      <div className="bg-white dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 px-4 py-1.5 flex items-center justify-between text-xs text-slate-500 select-none">
        <div className="flex items-center gap-3">
          <span>Page 1 of 1</span>
          <span>•</span>
          <span>{wordCount} Words</span>
          <span>•</span>
          <span>{charCount} Characters</span>
          <span>•</span>
          <span>~{Math.ceil(wordCount / 200)} min read</span>
        </div>

        <div className="flex items-center gap-3">
          <span className="capitalize">{margin} Margins</span>
          <span>•</span>
          <span className="uppercase">{paperSize}</span>
          <span>•</span>
          <div className="flex items-center gap-1.5">
            <button 
              onClick={() => setZoom(Math.max(50, zoom - 10))}
              className="hover:text-slate-900 dark:hover:text-white"
            >
              -
            </button>
            <span className="font-mono text-[11px]">{zoom}%</span>
            <button 
              onClick={() => setZoom(Math.min(200, zoom + 10))}
              className="hover:text-slate-900 dark:hover:text-white"
            >
              +
            </button>
          </div>
        </div>
      </div>

      {/* Insert Table Modal */}
      {isTableModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="bg-white dark:bg-slate-900 rounded-xl p-5 shadow-2xl border border-slate-200 dark:border-slate-800 w-80 space-y-4">
            <h3 className="font-semibold text-sm text-slate-900 dark:text-slate-100">
              Insert Table
            </h3>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs font-medium text-slate-500">Rows</label>
                <input
                  type="number"
                  min={1}
                  max={20}
                  value={tableRows}
                  onChange={(e) => setTableRows(Number(e.target.value))}
                  className="w-full px-2.5 py-1.5 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg"
                />
              </div>
              <div>
                <label className="text-xs font-medium text-slate-500">Columns</label>
                <input
                  type="number"
                  min={1}
                  max={10}
                  value={tableCols}
                  onChange={(e) => setTableCols(Number(e.target.value))}
                  className="w-full px-2.5 py-1.5 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg"
                />
              </div>
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setIsTableModalOpen(false)}
                className="px-3 py-1.5 text-xs text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg"
              >
                Cancel
              </button>
              <button
                id="btn-confirm-insert-table"
                onClick={handleInsertTable}
                className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg"
              >
                Insert
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
