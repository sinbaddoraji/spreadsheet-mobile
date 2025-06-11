export interface SheetCellValue {
    v?: string | number;
    m?: string;
    f?: string;
    ct?: { fa: string; t: string };
}

export interface SheetCellData {
    r: number;
    c: number;
    v: SheetCellValue;
}

export interface SheetData {
    id: string;
    name: string;
    plugin?: string;
    config?: {
        merge?: Record<string, any>;
        columnlen?: Record<string, number>;
        customWidth?: Record<string, number>;
    };
    celldata: SheetCellData[];
    calcChain?: any[];
}

export class SheetParser {
    static parse(content: string): SheetData[] {
        console.log('SheetParser: parsing content:', content);
        try {
            const data = JSON.parse(content);
            console.log('SheetParser: parsed JSON:', data);
            if (Array.isArray(data)) {
                const result = data.map(sheet => this.normalizeSheet(sheet));
                console.log('SheetParser: normalized sheets:', result);
                return result;
            }
            const result = [this.normalizeSheet(data)];
            console.log('SheetParser: normalized single sheet:', result);
            return result;
        } catch (error) {
            console.error('Failed to parse sheet:', error, 'Content:', content);
            return [];
        }
    }

    private static normalizeSheet(sheet: any): SheetData {
        return {
            id: sheet.id || 'sheet1',
            name: sheet.name || 'Sheet1',
            plugin: sheet.plugin,
            config: sheet.config || {},
            celldata: sheet.celldata || [],
            calcChain: sheet.calcChain
        };
    }

    static getCellValue(cellValue: SheetCellValue | null | undefined): string {
        if (!cellValue) return '';
        if (cellValue.m) return cellValue.m;
        if (cellValue.v !== undefined) return String(cellValue.v);
        return '';
    }

    static getSheetDimensions(celldata: SheetCellData[]): { rows: number; cols: number } {
        let maxRow = 0;
        let maxCol = 0;
        
        for (const cell of celldata) {
            maxRow = Math.max(maxRow, cell.r);
            maxCol = Math.max(maxCol, cell.c);
        }
        
        return { rows: maxRow + 1, cols: maxCol + 1 };
    }

    static createGrid(celldata: SheetCellData[]): {
        grid: (SheetCellValue | null)[][];
        rowMap: number[];
        colMap: number[];
    } {
        const { rows, cols } = this.getSheetDimensions(celldata);
        const fullGrid: (SheetCellValue | null)[][] = [];
        
        // Initialize full grid with nulls
        for (let r = 0; r < rows; r++) {
            fullGrid[r] = new Array(cols).fill(null);
        }
        
        // Fill in the actual data
        for (const cell of celldata) {
            if (cell.r < rows && cell.c < cols) {
                fullGrid[cell.r][cell.c] = cell.v;
            }
        }
        
        // Find non-empty rows and columns
        const nonEmptyRows: number[] = [];
        const nonEmptyCols: number[] = [];
        
        for (let r = 0; r < rows; r++) {
            if (fullGrid[r].some(cell => cell !== null && this.getCellValue(cell).trim() !== '')) {
                nonEmptyRows.push(r);
            }
        }
        
        for (let c = 0; c < cols; c++) {
            let hasContent = false;
            for (let r = 0; r < rows; r++) {
                if (fullGrid[r][c] !== null && this.getCellValue(fullGrid[r][c]).trim() !== '') {
                    hasContent = true;
                    break;
                }
            }
            if (hasContent) {
                nonEmptyCols.push(c);
            }
        }
        
        // Create compact grid with only non-empty rows/cols
        const compactGrid: (SheetCellValue | null)[][] = [];
        for (let r = 0; r < nonEmptyRows.length; r++) {
            compactGrid[r] = [];
            for (let c = 0; c < nonEmptyCols.length; c++) {
                compactGrid[r][c] = fullGrid[nonEmptyRows[r]][nonEmptyCols[c]];
            }
        }
        
        return {
            grid: compactGrid,
            rowMap: nonEmptyRows,
            colMap: nonEmptyCols
        };
    }
}