export type Color = number;
export type HorizontalAlignment = string;
export type VerticalAlignment = string;
export type BorderThickness = string;
export type FillPattern = string;

export interface CellStyle {
    setFont(name: string, size: number, color: Color | null, bold: boolean, italic: boolean, underline: boolean, strikethrough: boolean): void;
    setAlignment(horizontal: HorizontalAlignment, vertical: VerticalAlignment, wrapText: boolean): void;
    setBorderLeft(thickness: BorderThickness | null, color?: Color): void;
    setBorderRight(thickness: BorderThickness | null, color?: Color): void;
    setBorderTop(thickness: BorderThickness | null, color?: Color): void;
    setBorderBottom(thickness: BorderThickness | null, color?: Color): void;
    setBorderDiagonal(thickness: BorderThickness | null, color?: Color): void;
    setFill(pattern?: FillPattern, bgColor?: Color): void;
    fontsAreEquals(cellStyle: CellStyle): boolean;
    alignmentsAreEquals(cellStyle: CellStyle): boolean;
    bordersAreEquals(cellStyle: CellStyle): boolean;
    fillsAreEquals(cellStyle: CellStyle): boolean;
    hasAlignments(): boolean;
    hasBorders(): boolean;
    hasFills(): boolean;
    clone(): CellStyle;
};

export type Position = {
    x: number,
    y: number
};

export type Row = (values: Array<string>, styles: Array<CellStyle>, doComputeExtremes?: boolean) => void;

export interface WorksheetSource {
    getAuthor(): Promise<string>;
    getFrozenPosition(): Promise<Position>;
    getReadableStream(): Promise<ReadableStream<Row>>;
};

export type TableToSourceOptions = {
    author?: string,
    skipEmptyRows?: boolean
};

export declare function getColumnNameByIndex(n: number): string;
export declare function createSourceFromTableElement(tableElement: Element, options?: TableToSourceOptions): Streamer;
export declare function exportToExcel(source: WorksheetSource): Promise<Blob>;
