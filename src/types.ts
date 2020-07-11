/** Text Run */
export interface WJSTextRun {
  t: "s";
  /** Text content */
  v: string;
}

export interface WJSTableCell {
  t: "c";
  /** Body */
  p: WJSPara[];
}

/** Table Row */
export interface WJSTableRow {
  t: "r";
  /** Cells */
  c: WJSTableCell[];
}

/** Table */
export interface WJSTable {
  t: "t";
  /** Rows */
  r: WJSTableRow[];
}

/** Children elements of a Paragraph */
export type WJSParaElement = WJSTextRun | WJSTable;

/** Paragraph */
export interface WJSPara {
  /** Children */
  elts: WJSParaElement[];
}

/** WordJS Document */
export interface WJSDoc {
  p: WJSPara[];
}