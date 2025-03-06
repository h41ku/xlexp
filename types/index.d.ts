export type Color = number;
export type HorizontalAlignment = string;
export type VerticalAlignment = string;
export type BorderThickness = string;
export type FillPattern = string;

export interface CellStyle {
    setFont(name: string, size: number, color: Color | null, bold: boolean, italic: boolean, underline: boolean, strikethrough: boolean): void;
    setAlignment(horizontal: HorizontalAlignment, vertical: VerticalAlignment, wrapText: boolean): void;
    setBorderLeft(thickness: BorderThickness | null, color: Color | undefined): void;
    setBorderRight(thickness: BorderThickness | null, color: Color | undefined): void;
    setBorderTop(thickness: BorderThickness | null, color: Color | undefined): void;
    setBorderBottom(thickness: BorderThickness | null, color: Color | undefined): void;
    setBorderDiagonal(thickness: BorderThickness | null, color: Color | undefined): void;
    setFill(pattern: FillPattern | undefined, bgColor: Color | undefined): void;
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

export type callbackFn = (values: Array<string>, styles: Array<CellStyle>, doComputeExtremes: boolean | undefined) => void;

export interface Streamer {
    frozenPosition(): Promise<Position>;
    streamAll(callback: callbackFn): Promise<void>;
};

export declare function getColumnNameByIndex(n: number): string;
export declare function createStreamerFromTableElement(tableElement: Element, skipEmptyRows: boolean | undefined): Streamer;
export declare function downloadAs(sourceAsBlob: Blob, fileName: string): void;
export declare function exportToExcel(streamer: Streamer): Promise<Blob>;
