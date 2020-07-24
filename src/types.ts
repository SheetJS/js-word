/** Text Run */
export interface WJSTextRun {
  t: "s";
  /** Text content */
  v: string;
}

export interface WJSEndNote {
  t: "endnote";
  /** Id of the endnote element it's linked to */
  id: Number;
  /** Body */
  p: WJSPara[];
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
export type WJSParaElement = WJSTextRun | WJSTable | WJSEndNote;

/** Paragraph */
export interface WJSPara {
  /** Children */
  elts: WJSParaElement[];
}

/** WordJS Document */
export interface WJSDoc {
  p: WJSPara[];
  rels?: WJSRel[];
}

/** Relationship */
export interface WJSRel {
  t: "rel";
  id: String;
  type: String;
  target: String;
  targerMode?: String;
}
