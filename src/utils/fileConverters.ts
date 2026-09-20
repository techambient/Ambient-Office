import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import JSZip from 'jszip';
import pptxgen from 'pptxgenjs';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import {
  DocumentData,
  SpreadsheetData,
  SpreadsheetSheet,
  SpreadsheetCell,
  PresentationData,
  Slide,
  SlideElement,
  AmbientWordFile,
  AmbientSpreadsheetFile,
  AmbientPresentationFile,
} from '../types/office';

// Helper to download blob or text file
export function downloadFile(content: Blob | string, filename: string, mimeType?: string) {
  const blob = typeof content === 'string' 
    ? new Blob([content], { type: mimeType || 'application/octet-stream' }) 
    : content;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Convert Excel column index to letter (0 -> A, 27 -> AB)
export function colIndexToLetter(index: number): string {
  let temp = index;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

// Convert Excel column letter to index (A -> 0, AB -> 27)
export function colLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

// ==========================================
// DOCUMENT (WORD / .DOCX / .AW) CONVERTERS
// ==========================================

export async function parseDocxFile(file: File): Promise<DocumentData> {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    const html = result.value || '<p>Empty document</p>';
    return {
      htmlContent: html,
      margin: 'normal',
      paperSize: 'letter',
      orientation: 'portrait',
      zoom: 100,
    };
  } catch (err) {
    console.error('Failed to parse docx with mammoth, trying text fallback', err);
    const text = await file.text();
    return {
      htmlContent: `<p>${text.replace(/\n/g, '<br/>')}</p>`,
      margin: 'normal',
      paperSize: 'letter',
      orientation: 'portrait',
      zoom: 100,
    };
  }
}

export function parseAwFile(jsonStr: string): DocumentData {
  try {
    const parsed = JSON.parse(jsonStr) as AmbientWordFile;
    if (parsed.data && parsed.data.htmlContent !== undefined) {
      return parsed.data;
    }
  } catch (err) {
    console.error('Failed to parse .aw file', err);
  }
  return {
    htmlContent: '<p>Document restored</p>',
    margin: 'normal',
    paperSize: 'letter',
    orientation: 'portrait',
    zoom: 100,
  };
}

export function exportToAw(title: string, data: DocumentData): string {
  const fileData: AmbientWordFile = {
    format: 'ambient-word',
    version: '1.0',
    metadata: {
      title,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
    data,
  };
  return JSON.stringify(fileData, null, 2);
}

export async function exportDocxFromHtml(title: string, htmlContent: string): Promise<Blob> {
  // Dynamically import docx to keep bundle lightweight
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('docx');
  
  // Simple parser from HTML string to Docx elements
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = htmlContent;

  const docParagraphs: any[] = [];

  const processNode = (node: Node) => {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      const tagName = el.tagName.toLowerCase();

      if (tagName === 'h1') {
        docParagraphs.push(
          new Paragraph({
            text: el.innerText,
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 },
          })
        );
      } else if (tagName === 'h2') {
        docParagraphs.push(
          new Paragraph({
            text: el.innerText,
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          })
        );
      } else if (tagName === 'h3') {
        docParagraphs.push(
          new Paragraph({
            text: el.innerText,
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 160, after: 80 },
          })
        );
      } else if (tagName === 'p' || tagName === 'div' || tagName === 'li') {
        const textRuns: any[] = [];
        
        el.childNodes.forEach((child) => {
          if (child.nodeType === Node.TEXT_NODE) {
            textRuns.push(new TextRun({ text: child.textContent || '' }));
          } else if (child.nodeType === Node.ELEMENT_NODE) {
            const childEl = child as HTMLElement;
            const isBold = childEl.tagName === 'B' || childEl.tagName === 'STRONG' || childEl.style.fontWeight === 'bold';
            const isItalic = childEl.tagName === 'I' || childEl.tagName === 'EM' || childEl.style.fontStyle === 'italic';
            const isUnderline = childEl.tagName === 'U';
            textRuns.push(
              new TextRun({
                text: childEl.innerText || '',
                bold: isBold,
                italics: isItalic,
                underline: isUnderline ? {} : undefined,
              })
            );
          }
        });

        docParagraphs.push(
          new Paragraph({
            children: textRuns.length > 0 ? textRuns : [new TextRun(el.innerText || '')],
            spacing: { after: 120 },
            bullet: tagName === 'li' ? { level: 0 } : undefined,
          })
        );
      } else {
        // Fallback for other tags
        Array.from(el.childNodes).forEach(processNode);
      }
    }
  };

  Array.from(tempDiv.childNodes).forEach(processNode);

  if (docParagraphs.length === 0) {
    docParagraphs.push(new Paragraph({ text: tempDiv.innerText || 'Empty Document' }));
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: docParagraphs,
      },
    ],
  });

  return await Packer.toBlob(doc);
}

export async function exportHtmlToPdf(elementId: string, filename: string) {
  const el = document.getElementById(elementId);
  if (!el) return;

  const canvas = await html2canvas(el, {
    scale: 2,
    useCORS: true,
    logging: false,
    backgroundColor: '#ffffff',
  });

  const imgData = canvas.toDataURL('image/jpeg', 0.95);
  const pdf = new jsPDF('p', 'mm', 'a4');
  const imgWidth = 210; // A4 width in mm
  const pageHeight = 297; // A4 height in mm
  const imgHeight = (canvas.height * imgWidth) / canvas.width;
  let heightLeft = imgHeight;
  let position = 0;

  pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
  heightLeft -= pageHeight;

  while (heightLeft >= 0) {
    position = heightLeft - imgHeight;
    pdf.addPage();
    pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
    heightLeft -= pageHeight;
  }

  pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
}

// ==========================================
// SPREADSHEET (EXCEL / .XLSX / .AS) CONVERTERS
// ==========================================

export async function parseExcelFile(file: File): Promise<SpreadsheetData> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellFormula: true, cellStyles: true });

  const sheets: SpreadsheetSheet[] = workbook.SheetNames.map((sheetName, idx) => {
    const ws = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:M30');
    const cellData: Record<string, SpreadsheetCell> = {};

    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[cellAddress];
        if (cell) {
          cellData[`${R}:${C}`] = {
            value: cell.w || (cell.v !== undefined ? String(cell.v) : ''),
            formula: cell.f ? `=${cell.f}` : undefined,
          };
        }
      }
    }

    return {
      id: `sheet-${idx + 1}`,
      name: sheetName,
      data: cellData,
      rowCount: Math.max(range.e.r + 10, 40),
      colCount: Math.max(range.e.c + 5, 20),
    };
  });

  return {
    sheets: sheets.length > 0 ? sheets : [createDefaultSheet('Sheet1')],
    activeSheetId: sheets[0]?.id || 'sheet-1',
  };
}

export function parseAsFile(jsonStr: string): SpreadsheetData {
  try {
    const parsed = JSON.parse(jsonStr) as AmbientSpreadsheetFile;
    if (parsed.data && parsed.data.sheets) {
      return parsed.data;
    }
  } catch (err) {
    console.error('Failed to parse .as file', err);
  }
  return {
    sheets: [createDefaultSheet('Sheet1')],
    activeSheetId: 'sheet-1',
  };
}

export function exportToAs(title: string, data: SpreadsheetData): string {
  const fileData: AmbientSpreadsheetFile = {
    format: 'ambient-spreadsheet',
    version: '1.0',
    metadata: {
      title,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
    data,
  };
  return JSON.stringify(fileData, null, 2);
}

export function exportToExcel(data: SpreadsheetData): Blob {
  const wb = XLSX.utils.book_new();

  data.sheets.forEach((sheet) => {
    // Convert matrix to array of arrays
    const aoa: any[][] = [];
    const maxRow = sheet.rowCount || 30;
    const maxCol = sheet.colCount || 15;

    for (let r = 0; r < maxRow; r++) {
      const rowArr: any[] = [];
      for (let c = 0; c < maxCol; c++) {
        const cell = sheet.data[`${r}:${c}`];
        if (cell) {
          rowArr.push(cell.value || '');
        } else {
          rowArr.push('');
        }
      }
      aoa.push(rowArr);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, sheet.name.substring(0, 31));
  });

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export function exportSpreadsheetToPdf(data: SpreadsheetData, filename: string) {
  const pdf = new jsPDF('l', 'mm', 'a4'); // Landscape for spreadsheets
  const activeSheet = data.sheets.find((s) => s.id === data.activeSheetId) || data.sheets[0];

  pdf.setFontSize(14);
  pdf.text(`${filename} - ${activeSheet.name}`, 14, 15);

  const startX = 14;
  let startY = 25;
  const colWidth = 24;
  const rowHeight = 8;
  const maxCols = Math.min(activeSheet.colCount || 10, 10);
  const maxRows = Math.min(activeSheet.rowCount || 25, 25);

  // Draw Headers (A, B, C...)
  pdf.setFontSize(8);
  pdf.setFillColor(240, 240, 240);
  pdf.rect(startX, startY, colWidth * (maxCols + 1), rowHeight, 'F');
  pdf.text('#', startX + 2, startY + 5);

  for (let c = 0; c < maxCols; c++) {
    pdf.text(colIndexToLetter(c), startX + (c + 1) * colWidth + 2, startY + 5);
  }
  startY += rowHeight;

  // Draw rows
  for (let r = 0; r < maxRows; r++) {
    pdf.rect(startX, startY, colWidth, rowHeight);
    pdf.text(String(r + 1), startX + 2, startY + 5);

    for (let c = 0; c < maxCols; c++) {
      const cell = activeSheet.data[`${r}:${c}`];
      pdf.rect(startX + (c + 1) * colWidth, startY, colWidth, rowHeight);
      if (cell && cell.value) {
        const text = String(cell.value).substring(0, 12);
        pdf.text(text, startX + (c + 1) * colWidth + 2, startY + 5);
      }
    }
    startY += rowHeight;
    if (startY > 185) break; // page boundary
  }

  pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
}

export function createDefaultSheet(name = 'Sheet1', id = 'sheet-1'): SpreadsheetSheet {
  return {
    id,
    name,
    data: {},
    rowCount: 50,
    colCount: 26,
  };
}

// ==========================================
// PRESENTATION (POWERPOINT / .PPTX / .AP) CONVERTERS
// ==========================================

export async function parsePptxFile(file: File): Promise<PresentationData> {
  try {
    const zip = await JSZip.loadAsync(file);
    const slideFiles: string[] = [];

    // Find all slide XML files in ppt/slides/
    zip.forEach((relativePath) => {
      if (relativePath.match(/^ppt\/slides\/slide\d+\.xml$/)) {
        slideFiles.push(relativePath);
      }
    });

    // Sort by slide number
    slideFiles.sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)![0]);
      const numB = parseInt(b.match(/\d+/)![0]);
      return numA - numB;
    });

    const slides: Slide[] = [];

    for (let i = 0; i < slideFiles.length; i++) {
      const xmlStr = await zip.file(slideFiles[i])?.async('text');
      if (xmlStr) {
        // Extract text tags <a:t>text</a:t>
        const matches = xmlStr.match(/<a:t>(.*?)<\/a:t>/g) || [];
        const texts = matches.map((m) => m.replace(/<\/?a:t>/g, '').trim()).filter(Boolean);

        const title = texts[0] || `Slide ${i + 1}`;
        const contentTexts = texts.slice(1);

        const elements: SlideElement[] = [
          {
            id: `el-title-${i}`,
            type: 'heading',
            x: 8,
            y: 12,
            width: 84,
            height: 18,
            content: title,
            fontSize: 28,
            fontWeight: 'bold',
            color: '#1e293b',
          },
        ];

        if (contentTexts.length > 0) {
          elements.push({
            id: `el-content-${i}`,
            type: 'text',
            x: 8,
            y: 35,
            width: 84,
            height: 50,
            content: contentTexts.join('\n• '),
            fontSize: 18,
            color: '#334155',
          });
        }

        slides.push({
          id: `slide-${i + 1}`,
          title,
          layout: 'title-content',
          background: '#ffffff',
          elements,
        });
      }
    }

    if (slides.length > 0) {
      return {
        slides,
        activeSlideId: slides[0].id,
        theme: 'Modern Light',
        aspectRatio: '16:9',
      };
    }
  } catch (err) {
    console.error('Failed to parse .pptx archive', err);
  }

  return {
    slides: [createDefaultSlide(1)],
    activeSlideId: 'slide-1',
    theme: 'Modern Light',
    aspectRatio: '16:9',
  };
}

export function parseApFile(jsonStr: string): PresentationData {
  try {
    const parsed = JSON.parse(jsonStr) as AmbientPresentationFile;
    if (parsed.data && parsed.data.slides) {
      return parsed.data;
    }
  } catch (err) {
    console.error('Failed to parse .ap file', err);
  }
  return {
    slides: [createDefaultSlide(1)],
    activeSlideId: 'slide-1',
    theme: 'Modern Light',
    aspectRatio: '16:9',
  };
}

export function exportToAp(title: string, data: PresentationData): string {
  const fileData: AmbientPresentationFile = {
    format: 'ambient-presentation',
    version: '1.0',
    metadata: {
      title,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
    data,
  };
  return JSON.stringify(fileData, null, 2);
}

export async function exportToPptx(data: PresentationData, filename: string) {
  const pres = new pptxgen();
  pres.layout = data.aspectRatio === '4:3' ? 'LAYOUT_4x3' : 'LAYOUT_16x9';

  data.slides.forEach((slide) => {
    const pptSlide = pres.addSlide();
    
    // Background color
    if (slide.background) {
      pptSlide.background = { color: slide.background.replace('#', '') };
    }

    slide.elements.forEach((el) => {
      // Convert % coordinates to inches (standard 16:9 slide is 10 x 5.625 inches)
      const slideW = 10;
      const slideH = 5.625;

      const x = (el.x / 100) * slideW;
      const y = (el.y / 100) * slideH;
      const w = (el.width / 100) * slideW;
      const h = (el.height / 100) * slideH;

      if (el.type === 'heading' || el.type === 'text') {
        pptSlide.addText(el.content || '', {
          x,
          y,
          w,
          h,
          fontSize: el.fontSize ? Math.round(el.fontSize * 0.9) : (el.type === 'heading' ? 24 : 14),
          bold: el.fontWeight === 'bold' || el.type === 'heading',
          color: (el.color || '#000000').replace('#', ''),
          align: el.align || 'left',
        });
      } else if (el.type === 'shape') {
        pptSlide.addShape(pres.ShapeType.rect, {
          x,
          y,
          w,
          h,
          fill: { color: (el.backgroundColor || '#3b82f6').replace('#', '') },
          line: { color: (el.borderColor || '#1d4ed8').replace('#', ''), width: el.borderWidth || 1 },
        });
      } else if (el.type === 'badge') {
        pptSlide.addText(el.content || 'BADGE', {
          x,
          y,
          w,
          h,
          fontSize: 12,
          bold: true,
          color: (el.color || '#ffffff').replace('#', ''),
          fill: { color: (el.backgroundColor || '#2563eb').replace('#', '') },
          align: 'center',
        });
      } else if (el.type === 'image' && el.content && el.content.startsWith('data:image')) {
        pptSlide.addImage({
          data: el.content,
          x,
          y,
          w,
          h,
        });
      }
    });
  });

  await pres.writeFile({ fileName: filename.endsWith('.pptx') ? filename : `${filename}.pptx` });
}

export async function exportPresentationToPdf(data: PresentationData, filename: string) {
  const pdf = new jsPDF('l', 'mm', 'a4'); // Landscape

  for (let i = 0; i < data.slides.length; i++) {
    if (i > 0) pdf.addPage();
    const slide = data.slides[i];
    
    // Page dimensions
    const pWidth = 297;
    const pHeight = 210;

    // Fill background
    pdf.setFillColor(slide.background || '#ffffff');
    pdf.rect(0, 0, pWidth, pHeight, 'F');

    slide.elements.forEach((el) => {
      const x = (el.x / 100) * pWidth;
      const y = (el.y / 100) * pHeight;
      const w = (el.width / 100) * pWidth;
      const h = (el.height / 100) * pHeight;

      if (el.type === 'heading' || el.type === 'text') {
        pdf.setFontSize(el.type === 'heading' ? 22 : (el.fontSize || 14));
        pdf.setTextColor(el.color || '#1e293b');
        pdf.text(el.content || '', x, y + 6, { maxWidth: w });
      } else if (el.type === 'shape' || el.type === 'badge') {
        pdf.setFillColor(el.backgroundColor || '#3b82f6');
        pdf.rect(x, y, w, h, 'F');
        if (el.content) {
          pdf.setFontSize(12);
          pdf.setTextColor(el.color || '#ffffff');
          pdf.text(el.content, x + w / 2, y + h / 2 + 2, { align: 'center' });
        }
      }
    });

    // Slide footer
    pdf.setFontSize(8);
    pdf.setTextColor(150, 150, 150);
    pdf.text(`${i + 1} / ${data.slides.length}`, pWidth - 20, pHeight - 10);
  }

  pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
}

export function createDefaultSlide(index: number): Slide {
  return {
    id: `slide-${index}`,
    title: index === 1 ? 'Ambient Office Suite' : `Slide ${index}`,
    layout: index === 1 ? 'title' : 'title-content',
    background: '#ffffff',
    elements: index === 1
      ? [
          {
            id: 'el-1',
            type: 'heading',
            x: 10,
            y: 28,
            width: 80,
            height: 20,
            content: 'Ambient Office Presentation',
            fontSize: 36,
            fontWeight: 'bold',
            color: '#0f172a',
            align: 'center',
          },
          {
            id: 'el-2',
            type: 'text',
            x: 15,
            y: 50,
            width: 70,
            height: 15,
            content: 'High performance browser-based office suite. Compatible with .pptx and .ap',
            fontSize: 18,
            color: '#64748b',
            align: 'center',
          },
          {
            id: 'el-3',
            type: 'badge',
            x: 40,
            y: 70,
            width: 20,
            height: 8,
            content: 'VERSION 2026',
            color: '#ffffff',
            backgroundColor: '#2563eb',
            borderRadius: 6,
          },
        ]
      : [
          {
            id: 'el-h',
            type: 'heading',
            x: 8,
            y: 10,
            width: 84,
            height: 15,
            content: `Overview & Highlights #${index}`,
            fontSize: 28,
            fontWeight: 'bold',
            color: '#0f172a',
          },
          {
            id: 'el-t',
            type: 'text',
            x: 8,
            y: 30,
            width: 84,
            height: 50,
            content: '• Real-time browser-based multi-tab editing\n• Export to Microsoft PowerPoint or Standard PDF\n• Full screen interactive presentation mode\n• Customizable shapes, typography, and color schemes',
            fontSize: 18,
            color: '#334155',
          },
        ],
  };
}
