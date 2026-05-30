declare module 'xlsx-populate' {
  interface Cell {
    value(): unknown;
    value(v: unknown): Cell;
    address(): string;
    rowNumber(): number;
    columnNumber(): number;
    relativeCell(rowOffset: number, colOffset: number): Cell;
  }

  interface Row {
    cell(colNumber: number): Cell;
  }

  interface Sheet {
    name(): string;
    row(rowNumber: number): Row;
    cell(address: string): Cell;
    usedRange(): Range | undefined;
  }

  interface Range {
    forEach(callback: (cell: Cell, rowIndex: number, colIndex: number) => void): void;
  }

  interface Workbook {
    sheet(name: string): Sheet;
    sheet(index: number): Sheet;
    sheets(): Sheet[];
    outputAsync(options?: { type?: string }): Promise<Buffer>;
  }

  function fromFileAsync(path: string): Promise<Workbook>;
  function fromDataAsync(data: Buffer | ArrayBuffer): Promise<Workbook>;
}
