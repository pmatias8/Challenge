// Declarações mínimas do ExcelScript (só o que os scripts usam), com as
// mesmas assinaturas da API oficial, para checar os tipos fora do Excel.
declare namespace ExcelScript {
  type Valor = string | number | boolean;
  enum HorizontalAlignment { general = "General", left = "Left", center = "Center", right = "Right" }
  interface Workbook {
    getWorksheets(): Worksheet[];
    getWorksheet(name: string): Worksheet | undefined;
    addWorksheet(name?: string): Worksheet;
  }
  interface Worksheet {
    getName(): string;
    getRange(address?: string): Range;
    getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
    getCell(row: number, column: number): Range;
    getUsedRange(valuesOnly?: boolean): Range;
    getFreezePanes(): WorksheetFreezePanes;
    delete(): void;
    activate(): void;
  }
  interface Range {
    getValues(): Valor[][];
    setValues(values: Valor[][]): void;
    getValue(): Valor;
    setValue(value: Valor): void;
    setNumberFormat(numberFormat: string): void;
    getFormat(): RangeFormat;
    getRowIndex(): number;
    getRowCount(): number;
  }
  interface RangeFormat {
    getFont(): RangeFont;
    getFill(): RangeFill;
    setWrapText(wrapText: boolean): void;
    setRowHeight(rowHeight: number): void;
    setColumnWidth(columnWidth: number): void;
    autofitColumns(): void;
    setHorizontalAlignment(horizontalAlignment: HorizontalAlignment): void;
  }
  interface RangeFont { setBold(bold: boolean): void; setColor(color: string): void; setSize(size: number): void; }
  interface RangeFill { setColor(color: string): void; }
  interface WorksheetFreezePanes { freezeRows(count?: number): void; }
}
