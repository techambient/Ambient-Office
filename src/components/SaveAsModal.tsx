import { useState } from 'react';
import { 
  X, 
  Download, 
  FileText, 
  Table, 
  Presentation, 
  FileCheck, 
  Check, 
  Sparkles,
  Info
} from 'lucide-react';
import { OfficeTab } from '../types/office';
import { 
  exportDocxFromHtml, 
  exportToAw, 
  exportHtmlToPdf, 
  exportToExcel, 
  exportToAs, 
  exportSpreadsheetToPdf, 
  exportToPptx, 
  exportToAp, 
  exportPresentationToPdf,
  downloadFile
} from '../utils/fileConverters';

interface SaveAsModalProps {
  tab: OfficeTab;
  isOpen: boolean;
  onClose: () => void;
  onSaved: (newTitle: string, newExt: string) => void;
}

export const SaveAsModal = ({ tab, isOpen, onClose, onSaved }: SaveAsModalProps) => {
  const [fileName, setFileName] = useState(tab.title);
  const [selectedFormat, setSelectedFormat] = useState<string>(() => {
    if (tab.type === 'document') return 'docx';
    if (tab.type === 'spreadsheet') return 'xlsx';
    return 'pptx';
  });
  const [isExporting, setIsExporting] = useState(false);

  if (!isOpen) return null;

  const handleExport = async () => {
    setIsExporting(true);
    try {
      const cleanName = fileName.trim() || tab.title;

      if (tab.type === 'document' && tab.documentData) {
        if (selectedFormat === 'docx') {
          const blob = await exportDocxFromHtml(cleanName, tab.documentData.htmlContent);
          downloadFile(blob, `${cleanName}.docx`);
          onSaved(cleanName, 'docx');
        } else if (selectedFormat === 'pdf') {
          await exportHtmlToPdf('document-page-canvas', cleanName);
          onSaved(cleanName, 'pdf');
        } else if (selectedFormat === 'aw') {
          const content = exportToAw(cleanName, tab.documentData);
          downloadFile(content, `${cleanName}.aw`, 'application/javascript');
          onSaved(cleanName, 'aw');
        } else if (selectedFormat === 'html') {
          downloadFile(tab.documentData.htmlContent, `${cleanName}.html`, 'text/html');
          onSaved(cleanName, 'html');
        }
      } else if (tab.type === 'spreadsheet' && tab.spreadsheetData) {
        if (selectedFormat === 'xlsx') {
          const blob = exportToExcel(tab.spreadsheetData);
          downloadFile(blob, `${cleanName}.xlsx`);
          onSaved(cleanName, 'xlsx');
        } else if (selectedFormat === 'pdf') {
          exportSpreadsheetToPdf(tab.spreadsheetData, cleanName);
          onSaved(cleanName, 'pdf');
        } else if (selectedFormat === 'as') {
          const content = exportToAs(cleanName, tab.spreadsheetData);
          downloadFile(content, `${cleanName}.as`, 'application/javascript');
          onSaved(cleanName, 'as');
        }
      } else if (tab.type === 'presentation' && tab.presentationData) {
        if (selectedFormat === 'pptx') {
          await exportToPptx(tab.presentationData, cleanName);
          onSaved(cleanName, 'pptx');
        } else if (selectedFormat === 'pdf') {
          await exportPresentationToPdf(tab.presentationData, cleanName);
          onSaved(cleanName, 'pdf');
        } else if (selectedFormat === 'ap') {
          const content = exportToAp(cleanName, tab.presentationData);
          downloadFile(content, `${cleanName}.ap`, 'application/javascript');
          onSaved(cleanName, 'ap');
        }
      }
      onClose();
    } catch (err) {
      console.error('Export error:', err);
    } finally {
      setIsExporting(false);
    }
  };

  const getFormatOptions = () => {
    if (tab.type === 'document') {
      return [
        {
          id: 'docx',
          title: 'Microsoft Word Document',
          ext: '.docx',
          badge: 'Office Compatible',
          badgeColor: 'bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300',
          desc: 'Industry standard editable Word format compatible with MS 365, Google Docs, and LibreOffice.',
        },
        {
          id: 'pdf',
          title: 'Standard PDF Document',
          ext: '.pdf',
          badge: 'Portable Standard',
          badgeColor: 'bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300',
          desc: 'Universal fixed-layout vector format for printing, sharing, and archiving.',
        },
        {
          id: 'aw',
          title: 'Ambient Word Format (JavaScript)',
          ext: '.aw',
          badge: 'Native JavaScript',
          badgeColor: 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300',
          desc: 'Fast, lightweight JSON/JS document format for Ambient Office with 100% layout fidelity.',
        },
        {
          id: 'html',
          title: 'Web Document (HTML)',
          ext: '.html',
          badge: 'Web Standard',
          badgeColor: 'bg-slate-100 text-slate-800 dark:bg-slate-800 dark:text-slate-300',
          desc: 'Standard web page format readable by any browser.',
        },
      ];
    }

    if (tab.type === 'spreadsheet') {
      return [
        {
          id: 'xlsx',
          title: 'Microsoft Excel Workbook',
          ext: '.xlsx',
          badge: 'Office Compatible',
          badgeColor: 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300',
          desc: 'Standard multi-sheet workbook compatible with Microsoft Excel, Google Sheets, and Numbers.',
        },
        {
          id: 'pdf',
          title: 'Standard PDF Document',
          ext: '.pdf',
          badge: 'Portable Standard',
          badgeColor: 'bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300',
          desc: 'Formatted table layout ready for formal financial reporting and printing.',
        },
        {
          id: 'as',
          title: 'Ambient Spreadsheet Format (JavaScript)',
          ext: '.as',
          badge: 'Native JavaScript',
          badgeColor: 'bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300',
          desc: 'Lightweight JSON/JS format containing formulas, cell formats, styles, and sheet matrix.',
        },
      ];
    }

    // Presentation
    return [
      {
        id: 'pptx',
        title: 'Microsoft PowerPoint Presentation',
        ext: '.pptx',
        badge: 'Office Compatible',
        badgeColor: 'bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-300',
        desc: 'Standard presentation deck with slides, shapes, text formatting, and keynote layouts.',
      },
      {
        id: 'pdf',
        title: 'Standard PDF Slides',
        ext: '.pdf',
        badge: 'Portable Standard',
        badgeColor: 'bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300',
        desc: 'Clean landscape presentation slides export for handout distribution and presenting.',
      },
      {
        id: 'ap',
        title: 'Ambient Presentation Format (JavaScript)',
        ext: '.ap',
        badge: 'Native JavaScript',
        badgeColor: 'bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300',
        desc: 'Browser-native JSON/JS slide format preserving all layout coordinates, themes, and shapes.',
      },
    ];
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-xs p-4 animate-in fade-in duration-150">
      <div 
        id="save-as-dialog"
        className="w-full max-w-xl bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden flex flex-col max-h-[90vh]"
      >
        {/* Modal Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-blue-50 dark:bg-blue-950/60 text-blue-600 flex items-center justify-center">
              <Download className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">
                Save / Export File
              </h2>
              <p className="text-xs text-slate-500">
                Choose format: Microsoft Office, Standard PDF, or Ambient Native (.aw, .as, .ap)
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 space-y-5 overflow-y-auto">
          {/* File Name input */}
          <div>
            <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">
              File Name
            </label>
            <div className="flex items-center gap-2">
              <input
                id="input-save-filename"
                type="text"
                value={fileName}
                onChange={(e) => setFileName(e.target.value)}
                placeholder="Enter file name"
                className="flex-1 px-3.5 py-2 text-sm bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg text-slate-900 dark:text-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
              <span className="text-xs font-mono font-bold text-slate-500 px-3 py-2 bg-slate-100 dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700">
                .{selectedFormat}
              </span>
            </div>
          </div>

          {/* Format selection */}
          <div>
            <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-2">
              Select Output Format
            </label>
            <div className="space-y-2.5">
              {getFormatOptions().map((opt) => {
                const isSelected = selectedFormat === opt.id;
                return (
                  <div
                    key={opt.id}
                    id={`format-option-${opt.id}`}
                    onClick={() => setSelectedFormat(opt.id)}
                    className={`p-3.5 rounded-xl border-2 transition-all cursor-pointer flex items-start gap-3.5 ${
                      isSelected
                        ? 'border-blue-600 bg-blue-50/50 dark:bg-blue-950/20'
                        : 'border-slate-200 dark:border-slate-800 hover:border-slate-300 dark:hover:border-slate-700 bg-white dark:bg-slate-850'
                    }`}
                  >
                    <div className={`mt-0.5 w-5 h-5 rounded-full border flex items-center justify-center shrink-0 ${
                      isSelected
                        ? 'border-blue-600 bg-blue-600 text-white'
                        : 'border-slate-300 dark:border-slate-600'
                    }`}>
                      {isSelected && <Check className="w-3.5 h-3.5 stroke-[3]" />}
                    </div>

                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 mb-0.5">
                        <span className="font-semibold text-sm text-slate-900 dark:text-slate-100">
                          {opt.title}
                        </span>
                        <span className="text-xs font-mono font-bold text-slate-400">
                          {opt.ext}
                        </span>
                        <span className={`text-[10px] px-2 py-0.5 rounded-full font-medium ${opt.badgeColor}`}>
                          {opt.badge}
                        </span>
                      </div>
                      <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                        {opt.desc}
                      </p>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Format Info Note */}
          <div className="p-3 bg-slate-50 dark:bg-slate-800/60 rounded-xl flex items-center gap-2.5 text-xs text-slate-500 dark:text-slate-400 border border-slate-200/60 dark:border-slate-700/60">
            <Info className="w-4 h-4 text-blue-500 shrink-0" />
            <span>
              All exports are generated client-side with native browser performance. Files will download directly to your computer.
            </span>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-4 border-t border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50 flex items-center justify-end gap-3">
          <button
            onClick={onClose}
            className="px-4 py-2 text-xs font-medium text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            Cancel
          </button>
          <button
            id="btn-confirm-export"
            onClick={handleExport}
            disabled={isExporting}
            className="flex items-center gap-2 px-5 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-sm transition-all disabled:opacity-50 cursor-pointer"
          >
            <Download className="w-4 h-4" />
            <span>{isExporting ? 'Generating...' : `Download .${selectedFormat}`}</span>
          </button>
        </div>
      </div>
    </div>
  );
};
