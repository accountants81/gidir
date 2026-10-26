export type Direction = 'sales' | 'purchases';
export type Severity = 'error' | 'warning';

export interface ValidationIssue {
  rowNumber: number;
  header: string;
  columnIndex: number;
  cellReference: string;
  severity: Severity;
  reason: string;
  reasonAr?: string;
}

export interface NormalizationNotice {
  code: string;
  count: number;
  reason: string;
  reasonAr: string;
}

export interface FormulaCell {
  cellReference: string;
  rowNumber: number;
  header: string;
}

export interface ParsedWorkbook {
  direction: Direction;
  headers: string[];
  rows: string[][];
  issues: ValidationIssue[];
  normalizations: NormalizationNotice[];
  formulaCells: FormulaCell[];
  sourceFilename: string;
  sourceRowNumbers?: number[];
  sourceSheetName?: string;
}
