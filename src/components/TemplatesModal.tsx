import { 
  X, 
  FileText, 
  Table, 
  Presentation, 
  Sparkles, 
  ArrowRight,
  CheckCircle2
} from 'lucide-react';
import { OfficeDocType, OfficeTab } from '../types/office';
import { initialDocumentData, initialSpreadsheetData, initialPresentationData } from '../sampleData';

interface TemplatesModalProps {
  isOpen: boolean;
  onClose: () => void;
  onLoadTemplate: (tab: OfficeTab) => void;
}

export const TemplatesModal = ({ isOpen, onClose, onLoadTemplate }: TemplatesModalProps) => {
  if (!isOpen) return null;

  const templates = [
    {
      id: 'doc-report',
      type: 'document' as OfficeDocType,
      ext: 'docx',
      title: 'Executive Whitepaper & Report',
      desc: 'Complete corporate document featuring executive summaries, styled typography, data comparison tables, and formal sections.',
      badge: 'Word (.docx / .aw)',
      badgeColor: 'bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300',
      icon: <FileText className="w-5 h-5 text-blue-600" />,
      create: (): OfficeTab => ({
        id: `tab-sample-doc-${Date.now()}`,
        title: 'Executive Whitepaper',
        type: 'document',
        extension: 'docx',
        isModified: false,
        documentData: JSON.parse(JSON.stringify(initialDocumentData)),
      }),
    },
    {
      id: 'sheet-finance',
      type: 'spreadsheet' as OfficeDocType,
      ext: 'xlsx',
      title: '2026 Financial Projection & Margins',
      desc: 'Interactive financial workbook with multi-quarter revenue forecasts, live =SUM() and arithmetic formulas, and currency formatting.',
      badge: 'Excel (.xlsx / .as)',
      badgeColor: 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300',
      icon: <Table className="w-5 h-5 text-emerald-600" />,
      create: (): OfficeTab => ({
        id: `tab-sample-sheet-${Date.now()}`,
        title: '2026 Financial Model',
        type: 'spreadsheet',
        extension: 'xlsx',
        isModified: false,
        spreadsheetData: JSON.parse(JSON.stringify(initialSpreadsheetData)),
      }),
    },
    {
      id: 'pres-keynote',
      type: 'presentation' as OfficeDocType,
      ext: 'pptx',
      title: 'Product Strategy & Architecture Keynote',
      desc: 'Widescreen presentation deck with title branding, two-column feature blocks, custom shapes, badges, and slide show mode.',
      badge: 'PowerPoint (.pptx / .ap)',
      badgeColor: 'bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-300',
      icon: <Presentation className="w-5 h-5 text-amber-600" />,
      create: (): OfficeTab => ({
        id: `tab-sample-pres-${Date.now()}`,
        title: 'Product Strategy Keynote',
        type: 'presentation',
        extension: 'pptx',
        isModified: false,
        presentationData: JSON.parse(JSON.stringify(initialPresentationData)),
      }),
    },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-xs p-4 animate-in fade-in duration-150 select-none">
      <div className="w-full max-w-2xl bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden flex flex-col">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-amber-50 dark:bg-amber-950/60 text-amber-600 flex items-center justify-center">
              <Sparkles className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">
                Templates & Sample Files
              </h2>
              <p className="text-xs text-slate-500">
                Load fully working sample files for Word, Excel, and PowerPoint
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 rounded-lg"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* List of Templates */}
        <div className="p-6 space-y-3.5">
          {templates.map((tpl) => (
            <div
              key={tpl.id}
              onClick={() => {
                onLoadTemplate(tpl.create());
                onClose();
              }}
              className="group p-4 rounded-xl border border-slate-200 dark:border-slate-800 hover:border-blue-500 dark:hover:border-blue-500 hover:bg-blue-50/20 dark:hover:bg-blue-950/10 transition-all cursor-pointer flex items-start gap-4"
            >
              <div className="p-2.5 rounded-xl bg-slate-100 dark:bg-slate-800 group-hover:bg-white dark:group-hover:bg-slate-700 transition-colors shadow-2xs">
                {tpl.icon}
              </div>

              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 mb-1">
                  <h3 className="text-sm font-semibold text-slate-900 dark:text-slate-100 group-hover:text-blue-600 dark:group-hover:text-blue-400 transition-colors">
                    {tpl.title}
                  </h3>
                  <span className={`text-[10px] px-2 py-0.5 rounded-full font-medium ${tpl.badgeColor}`}>
                    {tpl.badge}
                  </span>
                </div>
                <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                  {tpl.desc}
                </p>
              </div>

              <div className="self-center">
                <span className="p-2 rounded-lg group-hover:bg-blue-600 group-hover:text-white text-slate-400 transition-colors flex items-center justify-center">
                  <ArrowRight className="w-4 h-4" />
                </span>
              </div>
            </div>
          ))}
        </div>

        {/* Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 text-xs font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-800 rounded-lg"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};
