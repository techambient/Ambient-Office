import { useState, useRef, useEffect, useMemo } from 'react';
import { 
  SpreadsheetData, 
  SpreadsheetSheet, 
  SpreadsheetCell 
} from '../../types/office';
import { colIndexToLetter, colLetterToIndex } from '../../utils/fileConverters';
import { 
  Bold, 
  Italic, 
  AlignLeft, 
  AlignCenter, 
  AlignRight, 
  DollarSign, 
  Percent, 
  Sigma, 
  Plus, 
  Trash2, 
  BarChart2, 
  ArrowUpDown, 
  Type, 
  Palette, 
  ChevronRight,
  Sparkles,
  Download,
  X
} from 'lucide-react';

interface SpreadsheetEditorProps {
  data: SpreadsheetData;
  onChange: (newData: SpreadsheetData) => void;
}

export const SpreadsheetEditor = ({ data, onChange }: SpreadsheetEditorProps) => {
  const activeSheetIndex = data.sheets.findIndex((s) => s.id === data.activeSheetId);
  const activeSheet = data.sheets[activeSheetIndex] || data.sheets[0];

  const [selectedRow, setSelectedRow] = useState<number>(0);
  const [selectedCol, setSelectedCol] = useState<number>(0);
  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [editValue, setEditValue] = useState<string>('');
  const [formulaBarValue, setFormulaBarValue] = useState<string>('');
  const [showChartModal, setShowChartModal] = useState<boolean>(false);
  const [chartType, setChartType] = useState<'bar' | 'line' | 'pie'>('bar');
  const [cellBgColor, setCellBgColor] = useState<string>('#ffffff');
  const [cellTextColor, setCellTextColor] = useState<string>('#1e293b');

  const cellInputRef = useRef<HTMLInputElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);

  const selectedCellKey = `${selectedRow}:${selectedCol}`;
  const currentCell = activeSheet.data[selectedCellKey];

  // Sync formula bar value when selection moves
  useEffect(() => {
    const rawVal = currentCell?.formula || currentCell?.value || '';
    setFormulaBarValue(rawVal);
    setEditValue(rawVal);
  }, [selectedRow, selectedCol, activeSheet]);

  // Evaluator function for cell formulas
  const evaluateCell = useMemo(() => {
    return (rawFormula: string, currentData: Record<string, SpreadsheetCell>): string => {
      if (!rawFormula.startsWith('=')) return rawFormula;

      const formula = rawFormula.substring(1).toUpperCase().trim();

      // Handle SUM(A1:B5)
      const sumMatch = formula.match(/^SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/);
      if (sumMatch) {
        const startCol = colLetterToIndex(sumMatch[1]);
        const startRow = parseInt(sumMatch[2], 10) - 1;
        const endCol = colLetterToIndex(sumMatch[3]);
        const endRow = parseInt(sumMatch[4], 10) - 1;

        let sum = 0;
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
          for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
            const cell = currentData[`${r}:${c}`];
            if (cell && cell.value) {
              const num = parseFloat(String(cell.value).replace(/[^0-9.-]+/g, ''));
              if (!isNaN(num)) sum += num;
            }
          }
        }
        return String(sum);
      }

      // Handle AVERAGE(A1:B5)
      const avgMatch = formula.match(/^AVERAGE\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/);
      if (avgMatch) {
        const startCol = colLetterToIndex(avgMatch[1]);
        const startRow = parseInt(avgMatch[2], 10) - 1;
        const endCol = colLetterToIndex(avgMatch[3]);
        const endRow = parseInt(avgMatch[4], 10) - 1;

        let sum = 0;
        let count = 0;
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
          for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
            const cell = currentData[`${r}:${c}`];
            if (cell && cell.value) {
              const num = parseFloat(String(cell.value).replace(/[^0-9.-]+/g, ''));
              if (!isNaN(num)) {
                sum += num;
                count++;
              }
            }
          }
        }
        return count > 0 ? (sum / count).toFixed(2) : '0';
      }

      // Handle simple math e.g. B4+B5 or F7-F9
      try {
        // Replace cell references with values
        const parsedExpr = formula.replace(/([A-Z]+)(\d+)/g, (_, colStr, rowStr) => {
          const c = colLetterToIndex(colStr);
          const r = parseInt(rowStr, 10) - 1;
          const target = currentData[`${r}:${c}`];
          if (!target || !target.value) return '0';
          const val = parseFloat(String(target.value).replace(/[^0-9.-]+/g, ''));
          return isNaN(val) ? '0' : String(val);
        });

        // Safe simple arithmetic evaluator
        if (/^[\d\s+\-*/().]+$/.test(parsedExpr)) {
          // eslint-disable-next-line no-new-func
          const result = Function(`"use strict"; return (${parsedExpr})`)();
          return String(result);
        }
      } catch (err) {
        return '#VALUE!';
      }

      return rawFormula;
    };
  }, []);

  const formatCellValue = (cell: SpreadsheetCell | undefined): string => {
    if (!cell) return '';
    const displayVal = cell.formula ? evaluateCell(cell.formula, activeSheet.data) : cell.value;
    if (!displayVal) return '';

    const num = parseFloat(String(displayVal).replace(/[^0-9.-]+/g, ''));
    if (isNaN(num)) return displayVal;

    if (cell.format === 'currency') {
      return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(num);
    }
    if (cell.format === 'percent') {
      return `${(num * 100).toFixed(1)}%`;
    }
    if (cell.format === 'number') {
      return new Intl.NumberFormat('en-US').format(num);
    }

    return String(displayVal);
  };

  const updateCell = (row: number, col: number, updates: Partial<SpreadsheetCell>) => {
    const key = `${row}:${col}`;
    const existing = activeSheet.data[key] || { value: '' };
    const updatedCell: SpreadsheetCell = { ...existing, ...updates };

    // If formula is provided, calculate value
    if (updates.formula !== undefined) {
      if (updates.formula.startsWith('=')) {
        updatedCell.formula = updates.formula;
        updatedCell.value = evaluateCell(updates.formula, activeSheet.data);
      } else {
        updatedCell.formula = undefined;
        updatedCell.value = updates.formula;
      }
    }

    const newSheetData = {
      ...activeSheet.data,
      [key]: updatedCell,
    };

    const newSheets = data.sheets.map((sheet) =>
      sheet.id === activeSheet.id ? { ...sheet, data: newSheetData } : sheet
    );

    onChange({
      ...data,
      sheets: newSheets,
    });
  };

  const commitEdit = () => {
    if (editValue.startsWith('=')) {
      updateCell(selectedRow, selectedCol, { formula: editValue });
    } else {
      updateCell(selectedRow, selectedCol, { value: editValue, formula: undefined });
    }
    setIsEditing(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (isEditing) {
      if (e.key === 'Enter') {
        commitEdit();
        setSelectedRow((r) => Math.min(r + 1, activeSheet.rowCount - 1));
      } else if (e.key === 'Escape') {
        setIsEditing(false);
        setEditValue(currentCell?.formula || currentCell?.value || '');
      }
      return;
    }

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        setSelectedRow((r) => Math.max(0, r - 1));
        break;
      case 'ArrowDown':
        e.preventDefault();
        setSelectedRow((r) => Math.min(activeSheet.rowCount - 1, r + 1));
        break;
      case 'ArrowLeft':
        e.preventDefault();
        setSelectedCol((c) => Math.max(0, c - 1));
        break;
      case 'ArrowRight':
        e.preventDefault();
        setSelectedCol((c) => Math.min(activeSheet.colCount - 1, c + 1));
        break;
      case 'Enter':
        e.preventDefault();
        setIsEditing(true);
        break;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        updateCell(selectedRow, selectedCol, { value: '', formula: undefined });
        break;
      default:
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
          setIsEditing(true);
          setEditValue(e.key);
        }
        break;
    }
  };

  const handleAutoSum = () => {
    // Look upwards from selected cell to find numbers
    let topRow = selectedRow - 1;
    while (topRow >= 0 && activeSheet.data[`${topRow}:${selectedCol}`]?.value) {
      topRow--;
    }
    const startR = topRow + 1;
    const endR = selectedRow - 1;
    if (endR >= startR) {
      const colLetter = colIndexToLetter(selectedCol);
      const formula = `=SUM(${colLetter}${startR + 1}:${colLetter}${endR + 1})`;
      updateCell(selectedRow, selectedCol, { formula, bold: true });
    }
  };

  // Add sheet
  const handleAddSheet = () => {
    const newId = `sheet-${data.sheets.length + 1}`;
    const newSheet: SpreadsheetSheet = {
      id: newId,
      name: `Sheet ${data.sheets.length + 1}`,
      data: {},
      rowCount: 50,
      colCount: 26,
    };
    onChange({
      ...data,
      sheets: [...data.sheets, newSheet],
      activeSheetId: newId,
    });
  };

  // Chart data extraction (first 10 rows with label col and numeric col)
  const chartPoints = useMemo(() => {
    const points: { label: string; value: number }[] = [];
    for (let r = 0; r < 12; r++) {
      const labelCell = activeSheet.data[`${r}:0`];
      const valCell = activeSheet.data[`${r}:1`];
      if (labelCell?.value && valCell?.value) {
        const num = parseFloat(String(valCell.value).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(num)) {
          points.push({
            label: String(labelCell.value).substring(0, 16),
            value: num,
          });
        }
      }
    }
    return points;
  }, [activeSheet]);

  return (
    <div 
      id="ambient-spreadsheet-editor" 
      className="flex flex-col h-full bg-white dark:bg-slate-900 overflow-hidden outline-none select-none"
      tabIndex={0}
      onKeyDown={handleKeyDown}
    >
      {/* Ribbon Toolbar */}
      <div className="bg-slate-50 dark:bg-slate-900/80 border-b border-slate-200 dark:border-slate-800 px-3 py-1.5 flex items-center gap-1.5 flex-wrap">
        {/* Font Style */}
        <div className="flex items-center gap-0.5 border-r border-slate-200 dark:border-slate-700 pr-2">
          <button
            id="btn-sheet-bold"
            onClick={() => updateCell(selectedRow, selectedCol, { bold: !currentCell?.bold })}
            className={`p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCell?.bold ? 'bg-slate-200 dark:bg-slate-800 text-blue-600 font-bold' : 'text-slate-700 dark:text-slate-300'
            }`}
            title="Bold (Ctrl+B)"
          >
            <Bold className="w-4 h-4" />
          </button>
          <button
            id="btn-sheet-italic"
            onClick={() => updateCell(selectedRow, selectedCol, { italic: !currentCell?.italic })}
            className={`p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCell?.italic ? 'bg-slate-200 dark:bg-slate-800 text-blue-600 italic' : 'text-slate-700 dark:text-slate-300'
            }`}
            title="Italic (Ctrl+I)"
          >
            <Italic className="w-4 h-4" />
          </button>
        </div>

        {/* Alignment */}
        <div className="flex items-center gap-0.5 border-r border-slate-200 dark:border-slate-700 pr-2">
          <button
            onClick={() => updateCell(selectedRow, selectedCol, { align: 'left' })}
            className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300"
            title="Align Left"
          >
            <AlignLeft className="w-4 h-4" />
          </button>
          <button
            onClick={() => updateCell(selectedRow, selectedCol, { align: 'center' })}
            className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300"
            title="Align Center"
          >
            <AlignCenter className="w-4 h-4" />
          </button>
          <button
            onClick={() => updateCell(selectedRow, selectedCol, { align: 'right' })}
            className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300"
            title="Align Right"
          >
            <AlignRight className="w-4 h-4" />
          </button>
        </div>

        {/* Formats */}
        <div className="flex items-center gap-1 border-r border-slate-200 dark:border-slate-700 pr-2">
          <button
            id="btn-format-currency"
            onClick={() => updateCell(selectedRow, selectedCol, { format: 'currency', align: 'right' })}
            className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 flex items-center gap-0.5 text-xs font-semibold"
            title="Format as Currency ($)"
          >
            <DollarSign className="w-4 h-4 text-emerald-600" />
          </button>
          <button
            id="btn-format-percent"
            onClick={() => updateCell(selectedRow, selectedCol, { format: 'percent', align: 'right' })}
            className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 flex items-center gap-0.5 text-xs font-semibold"
            title="Format as Percent (%)"
          >
            <Percent className="w-4 h-4 text-blue-600" />
          </button>
        </div>

        {/* Colors */}
        <div className="flex items-center gap-1.5 border-r border-slate-200 dark:border-slate-700 pr-2">
          <label className="flex items-center gap-1 cursor-pointer p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800" title="Cell Fill Color">
            <Palette className="w-4 h-4 text-slate-700 dark:text-slate-300" />
            <input
              type="color"
              value={cellBgColor}
              onChange={(e) => {
                setCellBgColor(e.target.value);
                updateCell(selectedRow, selectedCol, { bg: e.target.value });
              }}
              className="w-4 h-4 border-0 p-0 cursor-pointer bg-transparent"
            />
          </label>
          <label className="flex items-center gap-1 cursor-pointer p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800" title="Text Color">
            <Type className="w-4 h-4 text-slate-700 dark:text-slate-300" />
            <input
              type="color"
              value={cellTextColor}
              onChange={(e) => {
                setCellTextColor(e.target.value);
                updateCell(selectedRow, selectedCol, { color: e.target.value });
              }}
              className="w-4 h-4 border-0 p-0 cursor-pointer bg-transparent"
            />
          </label>
        </div>

        {/* AutoSum & Chart */}
        <div className="flex items-center gap-1">
          <button
            id="btn-sheet-autosum"
            onClick={handleAutoSum}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold bg-white dark:bg-slate-800 hover:bg-blue-50 dark:hover:bg-blue-950/40 text-blue-600 rounded border border-slate-200 dark:border-slate-700 shadow-2xs"
            title="AutoSum Formula (=SUM)"
          >
            <Sigma className="w-3.5 h-3.5" />
            <span>AutoSum</span>
          </button>

          <button
            id="btn-sheet-chart"
            onClick={() => setShowChartModal(true)}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold bg-white dark:bg-slate-800 hover:bg-emerald-50 dark:hover:bg-emerald-950/40 text-emerald-600 rounded border border-slate-200 dark:border-slate-700 shadow-2xs"
            title="Create Visual Chart from Sheet Data"
          >
            <BarChart2 className="w-3.5 h-3.5" />
            <span>Chart</span>
          </button>
        </div>
      </div>

      {/* Formula Bar */}
      <div className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 px-3 py-1 flex items-center gap-2">
        {/* Selected Cell Coordinate Label */}
        <div className="w-14 text-center font-mono text-xs font-semibold text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 py-1 rounded border border-slate-200 dark:border-slate-700">
          {colIndexToLetter(selectedCol)}{selectedRow + 1}
        </div>

        <span className="text-slate-400 font-mono text-xs italic font-bold select-none">
          fx
        </span>

        <input
          id="input-formula-bar"
          ref={formulaInputRef}
          type="text"
          value={formulaBarValue}
          onChange={(e) => setFormulaBarValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              if (formulaBarValue.startsWith('=')) {
                updateCell(selectedRow, selectedCol, { formula: formulaBarValue });
              } else {
                updateCell(selectedRow, selectedCol, { value: formulaBarValue, formula: undefined });
              }
            }
          }}
          placeholder="Enter a value or formula like =SUM(B2:B5) or =B2*1.15"
          className="flex-1 text-xs font-mono bg-slate-50 dark:bg-slate-800 px-3 py-1 rounded border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-slate-200 outline-none focus:border-blue-500"
        />
      </div>

      {/* Grid Container */}
      <div className="flex-1 overflow-auto bg-slate-100 dark:bg-slate-950">
        <table className="border-collapse table-fixed select-text">
          <thead>
            <tr>
              {/* Corner Header */}
              <th className="w-12 h-6 bg-slate-200 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 sticky top-0 left-0 z-30 text-[11px] text-slate-500 font-normal">
                ◢
              </th>
              {/* Column Headers (A, B, C...) */}
              {Array.from({ length: activeSheet.colCount }).map((_, c) => (
                <th
                  key={c}
                  className={`h-6 bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 sticky top-0 z-20 text-[11px] font-semibold text-slate-600 dark:text-slate-300 w-28 min-w-28 ${
                    selectedCol === c ? 'bg-blue-100 dark:bg-blue-900/60 text-blue-700' : ''
                  }`}
                >
                  {colIndexToLetter(c)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: activeSheet.rowCount }).map((_, r) => (
              <tr key={r}>
                {/* Row Header (1, 2, 3...) */}
                <td
                  className={`h-6 bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 sticky left-0 z-10 text-center text-[11px] font-semibold text-slate-600 dark:text-slate-300 ${
                    selectedRow === r ? 'bg-blue-100 dark:bg-blue-900/60 text-blue-700' : ''
                  }`}
                >
                  {r + 1}
                </td>

                {/* Cells */}
                {Array.from({ length: activeSheet.colCount }).map((_, c) => {
                  const isSelected = selectedRow === r && selectedCol === c;
                  const cell = activeSheet.data[`${r}:${c}`];
                  const formatted = formatCellValue(cell);

                  return (
                    <td
                      key={c}
                      id={`cell-${r}-${c}`}
                      onClick={() => {
                        setSelectedRow(r);
                        setSelectedCol(c);
                        setIsEditing(false);
                      }}
                      onDoubleClick={() => {
                        setSelectedRow(r);
                        setSelectedCol(c);
                        setIsEditing(true);
                      }}
                      style={{
                        backgroundColor: cell?.bg,
                        color: cell?.color,
                        textAlign: cell?.align || 'left',
                        fontWeight: cell?.bold ? 'bold' : 'normal',
                        fontStyle: cell?.italic ? 'italic' : 'normal',
                      }}
                      className={`h-6 px-1.5 text-xs border border-slate-200 dark:border-slate-800 overflow-hidden text-ellipsis whitespace-nowrap relative cursor-cell ${
                        isSelected
                          ? 'outline-2 outline-blue-600 -outline-offset-2 z-10 bg-blue-50/20'
                          : 'hover:bg-slate-50 dark:hover:bg-slate-800/50'
                      }`}
                    >
                      {isSelected && isEditing ? (
                        <input
                          ref={cellInputRef}
                          type="text"
                          value={editValue}
                          onChange={(e) => {
                            setEditValue(e.target.value);
                            setFormulaBarValue(e.target.value);
                          }}
                          onBlur={commitEdit}
                          autoFocus
                          className="absolute inset-0 w-full h-full px-1.5 text-xs bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 outline-none border-2 border-blue-600"
                        />
                      ) : (
                        <span>{formatted}</span>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Sheet Tabs Bar (Bottom) */}
      <div className="bg-slate-100 dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 px-3 py-1 flex items-center justify-between text-xs">
        <div className="flex items-center gap-1 overflow-x-auto no-scrollbar">
          {data.sheets.map((sheet) => {
            const isActive = sheet.id === data.activeSheetId;
            return (
              <button
                key={sheet.id}
                id={`sheet-tab-${sheet.id}`}
                onClick={() => onChange({ ...data, activeSheetId: sheet.id })}
                className={`px-3 py-1 rounded font-medium transition-colors ${
                  isActive
                    ? 'bg-white dark:bg-slate-800 text-emerald-600 dark:text-emerald-400 shadow-2xs'
                    : 'text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-800'
                }`}
              >
                {sheet.name}
              </button>
            );
          })}

          <button
            id="btn-add-sheet"
            onClick={handleAddSheet}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500"
            title="Add New Sheet"
          >
            <Plus className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Quick Calc Indicator */}
        <div className="text-slate-400 text-[11px] font-mono flex items-center gap-3">
          <span>Formula Engine Active</span>
          <span>•</span>
          <span>{activeSheet.rowCount} Rows × {activeSheet.colCount} Columns</span>
        </div>
      </div>

      {/* Chart Visualizer Modal */}
      {showChartModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white dark:bg-slate-900 rounded-2xl p-6 shadow-2xl border border-slate-200 dark:border-slate-800 w-full max-w-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <BarChart2 className="w-5 h-5 text-emerald-600" />
                <h3 className="font-semibold text-base text-slate-900 dark:text-slate-100">
                  Data Chart Visualization
                </h3>
              </div>
              <button
                onClick={() => setShowChartModal(false)}
                className="p-1 rounded-lg text-slate-400 hover:text-slate-600"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Chart Type Selector */}
            <div className="flex items-center gap-2">
              <button
                onClick={() => setChartType('bar')}
                className={`px-3 py-1.5 text-xs rounded-lg font-medium ${
                  chartType === 'bar' ? 'bg-emerald-600 text-white' : 'bg-slate-100 dark:bg-slate-800 text-slate-600'
                }`}
              >
                Bar Chart
              </button>
              <button
                onClick={() => setChartType('line')}
                className={`px-3 py-1.5 text-xs rounded-lg font-medium ${
                  chartType === 'line' ? 'bg-emerald-600 text-white' : 'bg-slate-100 dark:bg-slate-800 text-slate-600'
                }`}
              >
                Line Trend
              </button>
            </div>

            {/* Chart Canvas / Render */}
            <div className="h-64 bg-slate-50 dark:bg-slate-800/40 rounded-xl p-4 flex items-end justify-between gap-2 border border-slate-100 dark:border-slate-800">
              {chartPoints.length === 0 ? (
                <div className="m-auto text-xs text-slate-400">
                  No numeric columns detected in first rows. Enter values in Column B to see live charts.
                </div>
              ) : (
                chartPoints.map((pt, idx) => {
                  const maxVal = Math.max(...chartPoints.map((p) => p.value), 1);
                  const heightPct = Math.min(100, Math.max(10, Math.round((pt.value / maxVal) * 100)));

                  return (
                    <div key={idx} className="flex-1 flex flex-col items-center gap-1.5 h-full justify-end group">
                      <div className="text-[10px] text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity font-mono">
                        ${pt.value.toLocaleString()}
                      </div>
                      <div
                        style={{ height: `${heightPct}%` }}
                        className="w-full bg-gradient-to-t from-emerald-600 to-teal-500 rounded-t-md transition-all duration-300 hover:brightness-110"
                      />
                      <div className="text-[10px] text-slate-500 truncate w-full text-center" title={pt.label}>
                        {pt.label}
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            <div className="flex justify-end pt-2">
              <button
                onClick={() => setShowChartModal(false)}
                className="px-4 py-2 text-xs font-semibold text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 rounded-lg"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
