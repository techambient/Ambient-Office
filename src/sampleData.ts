import { OfficeTab, DocumentData, SpreadsheetData, PresentationData } from './types/office';

export const initialDocumentData: DocumentData = {
  htmlContent: `
    <h1 style="color: #1e293b; font-size: 26pt; font-weight: 700; margin-bottom: 8px;">Ambient Office — Product Whitepaper</h1>
    <p style="color: #64748b; font-size: 11pt; margin-bottom: 24px; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px;">
      <strong>Author:</strong> Ambient Architecture Group &nbsp;|&nbsp; <strong>Date:</strong> 2026 Edition &nbsp;|&nbsp; <strong>Status:</strong> Approved
    </p>

    <h2 style="color: #2563eb; font-size: 16pt; font-weight: 600; margin-top: 16px; margin-bottom: 8px;">1. Executive Summary</h2>
    <p style="font-size: 11pt; line-height: 1.6; color: #334155; margin-bottom: 14px;">
      Ambient Office is a lightweight, zero-latency browser-native office suite designed for modern knowledge workers. 
      It integrates full document authoring (Word / .aw), reactive calculation spreadsheets (Excel / .as), 
      and high-impact slide presentations (PowerPoint / .ap) directly inside any modern web browser without requiring heavy installations or native desktop runtimes.
    </p>

    <h2 style="color: #2563eb; font-size: 16pt; font-weight: 600; margin-top: 16px; margin-bottom: 8px;">2. Native & Universal Format Support</h2>
    <p style="font-size: 11pt; line-height: 1.6; color: #334155; margin-bottom: 12px;">
      Ambient Office bridges legacy productivity suites and browser computing with complete multi-format compatibility:
    </p>

    <table style="width: 100%; border-collapse: collapse; margin-bottom: 18px; font-size: 10pt;">
      <thead>
        <tr style="background-color: #f1f5f9; text-align: left;">
          <th style="padding: 8px 12px; border: 1px solid #cbd5e1;">Office Discipline</th>
          <th style="padding: 8px 12px; border: 1px solid #cbd5e1;">Microsoft Office</th>
          <th style="padding: 8px 12px; border: 1px solid #cbd5e1;">Ambient Native</th>
          <th style="padding: 8px 12px; border: 1px solid #cbd5e1;">Standard Portable</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1; font-weight: 600; color: #2563eb;">Word Processing</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1;">.docx, .doc</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1; font-weight: 600; color: #059669;">.aw (Ambient Word)</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1;">.pdf, .txt, .html</td>
        </tr>
        <tr style="background-color: #f8fafc;">
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1; font-weight: 600; color: #16a34a;">Spreadsheets</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1;">.xlsx, .xls</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1; font-weight: 600; color: #059669;">.as (Ambient Sheet)</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1;">.pdf, .csv</td>
        </tr>
        <tr>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1; font-weight: 600; color: #ea580c;">Presentations</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1;">.pptx, .ppt</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1; font-weight: 600; color: #059669;">.ap (Ambient Pres)</td>
          <td style="padding: 8px 12px; border: 1px solid #cbd5e1;">.pdf, Slide Show</td>
        </tr>
      </tbody>
    </table>

    <h2 style="color: #2563eb; font-size: 16pt; font-weight: 600; margin-top: 16px; margin-bottom: 8px;">3. Core Capabilities</h2>
    <ul style="margin-left: 20px; line-height: 1.8; color: #334155; font-size: 11pt;">
      <li><strong>Multi-Tab Editor:</strong> Work on documents, financial worksheets, and keynote slides concurrently.</li>
      <li><strong>Instant Client-Side Engine:</strong> Zero server latency, works offline, privacy-first storage.</li>
      <li><strong>One-Click Export:</strong> Export your files directly to genuine Microsoft Word (.docx), Excel (.xlsx), PowerPoint (.pptx), vector PDFs, or standard JavaScript formats (.aw, .as, .ap).</li>
    </ul>

    <p style="font-size: 10pt; color: #94a3b8; font-style: italic; margin-top: 24px;">
      Try typing directly in this page, inserting tables, changing typography, or downloading this file using the Save menu!
    </p>
  `,
  margin: 'normal',
  paperSize: 'letter',
  orientation: 'portrait',
  zoom: 100,
};

export const initialSpreadsheetData: SpreadsheetData = {
  sheets: [
    {
      id: 'sheet-1',
      name: 'Q1-Q4 Financials',
      rowCount: 35,
      colCount: 16,
      data: {
        '0:0': { value: 'Ambient Corp - 2026 Financial Projection', bold: true, color: '#1e293b' },
        
        // Table Headers
        '2:0': { value: 'Department / Revenue Stream', bold: true, bg: '#f1f5f9', align: 'left' },
        '2:1': { value: 'Q1 (USD)', bold: true, bg: '#f1f5f9', align: 'right' },
        '2:2': { value: 'Q2 (USD)', bold: true, bg: '#f1f5f9', align: 'right' },
        '2:3': { value: 'Q3 (USD)', bold: true, bg: '#f1f5f9', align: 'right' },
        '2:4': { value: 'Q4 (USD)', bold: true, bg: '#f1f5f9', align: 'right' },
        '2:5': { value: 'Annual Total', bold: true, bg: '#e2e8f0', color: '#1d4ed8', align: 'right' },

        // Rows
        '3:0': { value: 'Enterprise SaaS Subscriptions' },
        '3:1': { value: '145000', format: 'currency', align: 'right' },
        '3:2': { value: '172000', format: 'currency', align: 'right' },
        '3:3': { value: '210000', format: 'currency', align: 'right' },
        '3:4': { value: '255000', format: 'currency', align: 'right' },
        '3:5': { value: '782000', formula: '=SUM(B4:E4)', bold: true, format: 'currency', align: 'right' },

        '4:0': { value: 'Professional Services & Advisory' },
        '4:1': { value: '38000', format: 'currency', align: 'right' },
        '4:2': { value: '42000', format: 'currency', align: 'right' },
        '4:3': { value: '49000', format: 'currency', align: 'right' },
        '4:4': { value: '54000', format: 'currency', align: 'right' },
        '4:5': { value: '183000', formula: '=SUM(B5:E5)', bold: true, format: 'currency', align: 'right' },

        '5:0': { value: 'Developer API Credits' },
        '5:1': { value: '22000', format: 'currency', align: 'right' },
        '5:2': { value: '31000', format: 'currency', align: 'right' },
        '5:3': { value: '44000', format: 'currency', align: 'right' },
        '5:4': { value: '62000', format: 'currency', align: 'right' },
        '5:5': { value: '159000', formula: '=SUM(B6:E6)', bold: true, format: 'currency', align: 'right' },

        // Total Row
        '6:0': { value: 'Gross Revenue', bold: true, bg: '#f8fafc' },
        '6:1': { value: '205000', formula: '=SUM(B4:B6)', bold: true, format: 'currency', align: 'right' },
        '6:2': { value: '245000', formula: '=SUM(C4:C6)', bold: true, format: 'currency', align: 'right' },
        '6:3': { value: '303000', formula: '=SUM(D4:D6)', bold: true, format: 'currency', align: 'right' },
        '6:4': { value: '371000', formula: '=SUM(E4:E6)', bold: true, format: 'currency', align: 'right' },
        '6:5': { value: '1124000', formula: '=SUM(F4:F6)', bold: true, bg: '#dbeafe', color: '#1e40af', format: 'currency', align: 'right' },

        // Expenses
        '8:0': { value: 'Operating Costs & Infrastructure', italic: true, color: '#64748b' },
        '8:1': { value: '82000', format: 'currency', align: 'right' },
        '8:2': { value: '91000', format: 'currency', align: 'right' },
        '8:3': { value: '105000', format: 'currency', align: 'right' },
        '8:4': { value: '118000', format: 'currency', align: 'right' },
        '8:5': { value: '396000', formula: '=SUM(B9:E9)', bold: true, format: 'currency', align: 'right' },

        // Net Margin
        '10:0': { value: 'Net Operating Income', bold: true, bg: '#ecfdf5', color: '#047857' },
        '10:1': { value: '123000', formula: '=B7-B9', bold: true, format: 'currency', align: 'right' },
        '10:2': { value: '154000', formula: '=C7-C9', bold: true, format: 'currency', align: 'right' },
        '10:3': { value: '198000', formula: '=D7-D9', bold: true, format: 'currency', align: 'right' },
        '10:4': { value: '253000', formula: '=E7-E9', bold: true, format: 'currency', align: 'right' },
        '10:5': { value: '728000', formula: '=F7-F9', bold: true, bg: '#d1fae5', color: '#065f46', format: 'currency', align: 'right' },
      },
    },
    {
      id: 'sheet-2',
      name: 'Headcount & Staffing',
      rowCount: 20,
      colCount: 10,
      data: {
        '0:0': { value: 'Team Headcount by Quarter', bold: true },
        '2:0': { value: 'Role', bold: true, bg: '#f1f5f9' },
        '2:1': { value: 'Q1', bold: true, bg: '#f1f5f9', align: 'center' },
        '2:2': { value: 'Q2', bold: true, bg: '#f1f5f9', align: 'center' },
        '2:3': { value: 'Q3', bold: true, bg: '#f1f5f9', align: 'center' },
        '2:4': { value: 'Q4', bold: true, bg: '#f1f5f9', align: 'center' },
        '3:0': { value: 'Software Engineering' },
        '3:1': { value: '12', align: 'center' },
        '3:2': { value: '16', align: 'center' },
        '3:3': { value: '22', align: 'center' },
        '3:4': { value: '28', align: 'center' },
        '4:0': { value: 'Product & Design' },
        '4:1': { value: '4', align: 'center' },
        '4:2': { value: '6', align: 'center' },
        '4:3': { value: '8', align: 'center' },
        '4:4': { value: '10', align: 'center' },
      },
    },
  ],
  activeSheetId: 'sheet-1',
};

export const initialPresentationData: PresentationData = {
  theme: 'Modern Slate',
  aspectRatio: '16:9',
  activeSlideId: 'slide-1',
  slides: [
    {
      id: 'slide-1',
      title: 'Ambient Office Overview',
      layout: 'title',
      background: '#0f172a',
      elements: [
        {
          id: 'el-p1',
          type: 'badge',
          x: 40,
          y: 20,
          width: 20,
          height: 7,
          content: 'AMBIENT SUITE 2026',
          backgroundColor: '#3b82f6',
          color: '#ffffff',
          borderRadius: 20,
          align: 'center',
        },
        {
          id: 'el-p2',
          type: 'heading',
          x: 10,
          y: 32,
          width: 80,
          height: 22,
          content: 'Ambient Office Suite',
          fontSize: 44,
          fontWeight: 'bold',
          color: '#ffffff',
          align: 'center',
        },
        {
          id: 'el-p3',
          type: 'text',
          x: 15,
          y: 58,
          width: 70,
          height: 14,
          content: 'Professional In-Browser Word, Excel, and PowerPoint Productivity',
          fontSize: 20,
          color: '#94a3b8',
          align: 'center',
        },
      ],
    },
    {
      id: 'slide-2',
      title: 'Universal File Compatibility',
      layout: 'title-content',
      background: '#ffffff',
      elements: [
        {
          id: 'el-s2-h',
          type: 'heading',
          x: 8,
          y: 8,
          width: 84,
          height: 14,
          content: 'Full Office File Format Support',
          fontSize: 30,
          fontWeight: 'bold',
          color: '#0f172a',
        },
        {
          id: 'el-s2-t1',
          type: 'text',
          x: 8,
          y: 26,
          width: 40,
          height: 60,
          content: 'Microsoft Office Formats:\n• Word Documents (.docx, .doc)\n• Excel Spreadsheets (.xlsx, .xls)\n• PowerPoint Slides (.pptx, .ppt)\n\nPortable Standard:\n• Vector PDF Documents (.pdf)\n• CSV Data Sheets (.csv)',
          fontSize: 17,
          color: '#334155',
        },
        {
          id: 'el-s2-t2',
          type: 'text',
          x: 52,
          y: 26,
          width: 40,
          height: 60,
          content: 'Ambient Native Formats:\n• .aw — Ambient Word (JSON/JS)\n• .as — Ambient Spreadsheet\n• .ap — Ambient Presentation\n\nLightweight Architecture:\n• Zero server delays\n• Fully client-side rendering\n• Instant tab switching',
          fontSize: 17,
          color: '#334155',
        },
      ],
    },
    {
      id: 'slide-3',
      title: 'Feature Architecture',
      layout: 'two-column',
      background: '#f8fafc',
      elements: [
        {
          id: 'el-s3-h',
          type: 'heading',
          x: 8,
          y: 10,
          width: 84,
          height: 14,
          content: 'Designed For Professional Workflows',
          fontSize: 32,
          fontWeight: 'bold',
          color: '#1e293b',
        },
        {
          id: 'el-s3-c1',
          type: 'shape',
          x: 8,
          y: 30,
          width: 26,
          height: 52,
          content: '',
          backgroundColor: '#eff6ff',
          borderColor: '#bfdbfe',
          borderWidth: 2,
          borderRadius: 8,
        },
        {
          id: 'el-s3-t1',
          type: 'text',
          x: 10,
          y: 34,
          width: 22,
          height: 44,
          content: 'Ambient Word (.aw)\n\n• Paginated paper simulator\n• Rich styling, headings, lists\n• Live word/character counts\n• Real .docx & PDF export',
          fontSize: 15,
          color: '#1e40af',
        },
        {
          id: 'el-s3-c2',
          type: 'shape',
          x: 37,
          y: 30,
          width: 26,
          height: 52,
          content: '',
          backgroundColor: '#f0fdf4',
          borderColor: '#bbf7d0',
          borderWidth: 2,
          borderRadius: 8,
        },
        {
          id: 'el-s3-t2',
          type: 'text',
          x: 39,
          y: 34,
          width: 22,
          height: 44,
          content: 'Ambient Sheet (.as)\n\n• Live formula engine (SUM, etc)\n• Multi-sheet workbook tabs\n• Currency, %, decimal formatting\n• Native Excel .xlsx & PDF export',
          fontSize: 15,
          color: '#166534',
        },
        {
          id: 'el-s3-c3',
          type: 'shape',
          x: 66,
          y: 30,
          width: 26,
          height: 52,
          content: '',
          backgroundColor: '#fff7ed',
          borderColor: '#fed7aa',
          borderWidth: 2,
          borderRadius: 8,
        },
        {
          id: 'el-s3-t3',
          type: 'text',
          x: 68,
          y: 34,
          width: 22,
          height: 44,
          content: 'Ambient Slide (.ap)\n\n• Full drag & style canvas\n• Slide navigator & reordering\n• Presentation Slide Show mode\n• Genuine .pptx & PDF export',
          fontSize: 15,
          color: '#9a3412',
        },
      ],
    },
  ],
};

export const defaultTabs: OfficeTab[] = [
  {
    id: 'tab-1',
    title: 'Product Whitepaper',
    type: 'document',
    extension: 'docx',
    isModified: false,
    documentData: initialDocumentData,
  },
  {
    id: 'tab-2',
    title: '2026 Financial Projection',
    type: 'spreadsheet',
    extension: 'xlsx',
    isModified: false,
    spreadsheetData: initialSpreadsheetData,
  },
  {
    id: 'tab-3',
    title: 'Product Launch Keynote',
    type: 'presentation',
    extension: 'pptx',
    isModified: false,
    presentationData: initialPresentationData,
  },
];
