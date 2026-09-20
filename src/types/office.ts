export type OfficeDocType = 'document' | 'spreadsheet' | 'presentation';

export interface OfficeTab {
  id: string;
  title: string;
  type: OfficeDocType;
  extension: string; // e.g., 'docx', 'aw', 'xlsx', 'as', 'pptx', 'ap'
  isModified: boolean;
  documentData?: DocumentData;
  spreadsheetData?: SpreadsheetData;
  presentationData?: PresentationData;
}

export interface DocumentData {
  htmlContent: string;
  margin: 'normal' | 'narrow' | 'wide';
  paperSize: 'letter' | 'a4';
  orientation: 'portrait' | 'landscape';
  zoom: number;
}

export interface SpreadsheetCell {
  value: string;
  formula?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  bg?: string;
  align?: 'left' | 'center' | 'right';
  format?: 'general' | 'currency' | 'percent' | 'number';
}

export interface SpreadsheetSheet {
  id: string;
  name: string;
  data: Record<string, SpreadsheetCell>; // key is "row:col" or "A1"
  colWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  rowCount: number;
  colCount: number;
}

export interface SpreadsheetData {
  sheets: SpreadsheetSheet[];
  activeSheetId: string;
}

export interface SlideElement {
  id: string;
  type: 'text' | 'shape' | 'image' | 'heading' | 'badge';
  x: number; // percentage or px
  y: number;
  width: number;
  height: number;
  content: string; // text, shape type, or image URL
  fontSize?: number;
  fontFamily?: string;
  fontWeight?: string;
  color?: string;
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
  borderRadius?: number;
  align?: 'left' | 'center' | 'right';
  rotation?: number;
}

export interface Slide {
  id: string;
  title: string;
  layout: 'title' | 'title-content' | 'two-column' | 'quote' | 'section' | 'blank';
  background: string;
  elements: SlideElement[];
  notes?: string;
}

export interface PresentationData {
  slides: Slide[];
  activeSlideId: string;
  theme: string;
  aspectRatio: '16:9' | '4:3';
}

export interface AmbientWordFile {
  format: 'ambient-word';
  version: '1.0';
  metadata: {
    title: string;
    createdAt: string;
    updatedAt: string;
  };
  data: DocumentData;
}

export interface AmbientSpreadsheetFile {
  format: 'ambient-spreadsheet';
  version: '1.0';
  metadata: {
    title: string;
    createdAt: string;
    updatedAt: string;
  };
  data: SpreadsheetData;
}

export interface AmbientPresentationFile {
  format: 'ambient-presentation';
  version: '1.0';
  metadata: {
    title: string;
    createdAt: string;
    updatedAt: string;
  };
  data: PresentationData;
}
