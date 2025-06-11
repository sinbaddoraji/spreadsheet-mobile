import { TextFileView, TFile, WorkspaceLeaf } from 'obsidian';
import { SheetParser, SheetData } from './sheetParser';

export const VIEW_TYPE_SHEET = 'sheet-view';

export class SheetView extends TextFileView {
    private sheetData: SheetData[] = [];

    constructor(leaf: WorkspaceLeaf) {
        super(leaf);
        console.log('SheetView: Constructor called');
    }

    getViewType(): string {
        return VIEW_TYPE_SHEET;
    }

    getDisplayText(): string {
        return this.file?.basename || 'Sheet';
    }

    getViewData(): string {
        return this.data;
    }

    setViewData(data: string, clear: boolean): void {
        console.log('SheetView: setViewData called with data:', data);
        this.data = data;
        this.sheetData = SheetParser.parse(data);
        console.log('SheetView: parsed sheet data:', this.sheetData);
        this.render();
    }

    clear(): void {
        this.data = '';
        this.sheetData = [];
        this.contentEl.empty();
    }

    private render(): void {
        console.log('SheetView: render called, sheetData:', this.sheetData);
        this.contentEl.empty();
        this.contentEl.addClass('mobile-sheet-viewer');

        if (this.sheetData.length === 0) {
            console.log('SheetView: No sheet data found');
            this.contentEl.createEl('div', { 
                text: 'No sheet data found',
                cls: 'sheet-error' 
            });
            return;
        }

        this.sheetData.forEach((sheet, index) => {
            console.log('SheetView: rendering sheet', index, sheet);
            this.renderSheet(sheet, index);
        });
    }

    private renderSheet(sheet: SheetData, index: number): void {
        const sheetContainer = this.contentEl.createEl('div', { 
            cls: 'sheet-container' 
        });

        if (this.sheetData.length > 1) {
            sheetContainer.createEl('h3', { 
                text: sheet.name,
                cls: 'sheet-title' 
            });
        }

        const tableContainer = sheetContainer.createEl('div', { 
            cls: 'sheet-table-container' 
        });

        const table = tableContainer.createEl('table', { 
            cls: 'sheet-table' 
        });

        this.renderTable(table, sheet);
    }

    private renderTable(table: HTMLTableElement, sheet: SheetData): void {
        const gridData = SheetParser.createGrid(sheet.celldata);
        const { grid, rowMap, colMap } = gridData;
        
        if (grid.length === 0 || (grid[0] && grid[0].length === 0)) {
            table.createEl('tr').createEl('td', {
                text: 'No data',
                cls: 'sheet-no-data'
            });
            return;
        }

        for (let r = 0; r < grid.length; r++) {
            const row = table.createEl('tr', { cls: 'sheet-row' });
            
            for (let c = 0; c < grid[r].length; c++) {
                const cellValue = SheetParser.getCellValue(grid[r][c]);
                
                const td = row.createEl('td', { 
                    cls: 'sheet-cell',
                    text: cellValue
                });

                // Apply header styling to first non-empty row/column
                if (r === 0 && rowMap[0] === 0) {
                    td.addClass('sheet-header-row');
                }
                if (c === 0 && colMap[0] === 0) {
                    td.addClass('sheet-header-col');
                }
            }
        }
    }
}