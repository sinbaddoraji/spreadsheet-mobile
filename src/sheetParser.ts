export interface SheetCellValue {
    v?: string | number;
    m?: string;
    f?: string;
    ct?: { fa: string; t: string };
    calculated?: boolean; // Indicates if this is a calculated value from a formula
    s?: CellFormat; // Cell formatting
}

export interface CellFormat {
    // Font properties
    fontFamily?: string;
    fontSize?: number;
    fontWeight?: 'normal' | 'bold' | 'bolder' | 'lighter' | number;
    fontStyle?: 'normal' | 'italic' | 'oblique';
    textDecoration?: 'none' | 'underline' | 'overline' | 'line-through';
    
    // Text properties
    color?: string;
    textAlign?: 'left' | 'center' | 'right' | 'justify';
    verticalAlign?: 'top' | 'middle' | 'bottom';
    textWrap?: boolean;
    
    // Background and borders
    backgroundColor?: string;
    border?: CellBorder;
    
    // Number formatting
    numberFormat?: NumberFormat;
    
    // Custom styling
    customCss?: string;
}

export interface CellBorder {
    top?: BorderStyle;
    right?: BorderStyle;
    bottom?: BorderStyle;
    left?: BorderStyle;
    all?: BorderStyle; // Shorthand for all sides
}

export interface BorderStyle {
    width?: number;
    style?: 'none' | 'solid' | 'dashed' | 'dotted' | 'double';
    color?: string;
}

export interface NumberFormat {
    type: 'general' | 'number' | 'currency' | 'percentage' | 'date' | 'time' | 'scientific' | 'fraction' | 'text' | 'custom';
    decimalPlaces?: number;
    thousandsSeparator?: boolean;
    currencySymbol?: string;
    dateFormat?: string;
    timeFormat?: string;
    customFormat?: string;
}

export interface FormatPreset {
    id: string;
    name: string;
    description: string;
    format: CellFormat;
    category: 'text' | 'number' | 'date' | 'currency' | 'highlight' | 'custom';
}

export interface CellState {
    originalValue: SheetCellValue | null;
    currentValue: SheetCellValue | null;
    isModified: boolean;
    isValid: boolean;
    validationError?: string;
    lastModified?: Date;
    originalFormat?: CellFormat | null;
    currentFormat?: CellFormat | null;
    isFormatModified: boolean;
}

export interface ValidationRule {
    type: 'required' | 'number' | 'text' | 'regex' | 'custom';
    message: string;
    validate?: (value: string) => boolean;
    pattern?: RegExp;
}

export interface CellValidation {
    rules: ValidationRule[];
    isValid: boolean;
    errors: string[];
}

import { FormulaEngine, CellReference, FormulaContext, FormulaResult } from './formulaEngine';

export interface SheetCellData {
    r: number;
    c: number;
    v: SheetCellValue;
}

export interface FormulaCell {
    row: number;
    col: number;
    formula: string;
    dependencies: string[]; // Array of cell references this formula depends on
    dependents: string[];   // Array of cell references that depend on this formula
}

export interface SheetData {
    id: string;
    name: string;
    plugin?: string;
    config?: {
        merge?: Record<string, any>;
        columnlen?: Record<string, number>;
        customWidth?: Record<string, number>;
        validationRules?: Record<string, ValidationRule[]>;
    };
    celldata: SheetCellData[];
    calcChain?: any[];
    formulas?: Map<string, FormulaCell>; // Key: "row,col"
}

export interface UndoRedoAction {
    type: 'cell_edit' | 'cell_delete' | 'row_insert' | 'row_delete' | 'column_insert' | 'column_delete';
    sheetId: string;
    timestamp: Date;
    data: {
        row?: number;
        col?: number;
        oldValue?: SheetCellValue | null;
        newValue?: SheetCellValue | null;
        affectedCells?: Array<{row: number, col: number, oldValue: SheetCellValue | null, newValue: SheetCellValue | null}>;
    };
}

export interface AutoSaveConfig {
    enabled: boolean;
    intervalMs: number;
    maxRetries: number;
    debounceMs: number;
}

export interface AutoSaveState {
    lastSaved: Date | null;
    isSaving: boolean;
    lastError: string | null;
    saveCount: number;
    retryCount: number;
}

export interface ConflictInfo {
    type: 'file_modified' | 'concurrent_edit' | 'version_mismatch';
    conflictId: string;
    timestamp: Date;
    conflictedCells: Array<{
        row: number;
        col: number;
        localValue: SheetCellValue | null;
        remoteValue: SheetCellValue | null;
        baseValue?: SheetCellValue | null; // Original value before any changes
    }>;
    fileChecksum?: string;
    lastKnownModified?: Date;
}

export interface ConflictResolution {
    strategy: 'keep_local' | 'keep_remote' | 'merge' | 'manual';
    resolvedCells: Array<{
        row: number;
        col: number;
        resolvedValue: SheetCellValue | null;
        reasoning?: string;
    }>;
    timestamp: Date;
}

export interface SheetStateManager {
    undoStack: UndoRedoAction[];
    redoStack: UndoRedoAction[];
    cellStates: Map<string, CellState>;
    validationRules: Map<string, ValidationRule[]>;
    isDirty: boolean;
    maxUndoSteps: number;
    autoSaveConfig: AutoSaveConfig;
    autoSaveState: AutoSaveState;
    conflictInfo?: ConflictInfo;
    lastFileChecksum?: string;
    lastFileModified?: Date;
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

    static updateCellValue(sheetData: SheetData, row: number, col: number, value: string): void {
        const existingCellIndex = sheetData.celldata.findIndex(cell => cell.r === row && cell.c === col);
        
        if (value.trim() === '') {
            if (existingCellIndex !== -1) {
                sheetData.celldata.splice(existingCellIndex, 1);
            }
            return;
        }

        const cellValue: SheetCellValue = {
            v: this.parseValue(value),
            m: value
        };

        if (existingCellIndex !== -1) {
            sheetData.celldata[existingCellIndex].v = cellValue;
        } else {
            sheetData.celldata.push({
                r: row,
                c: col,
                v: cellValue
            });
        }
    }

    static parseValue(value: string): string | number {
        const trimmed = value.trim();
        if (trimmed === '') return '';
        
        const num = Number(trimmed);
        if (!isNaN(num) && isFinite(num)) {
            return num;
        }
        return trimmed;
    }

    static serialize(sheets: SheetData[]): string {
        return JSON.stringify(sheets, null, 2);
    }

    static getCellAt(sheetData: SheetData, row: number, col: number): SheetCellData | null {
        return sheetData.celldata.find(cell => cell.r === row && cell.c === col) || null;
    }

    static validateCell(value: string, rules: ValidationRule[]): CellValidation {
        const validation: CellValidation = {
            rules,
            isValid: true,
            errors: []
        };

        for (const rule of rules) {
            let isValid = true;
            
            switch (rule.type) {
                case 'required':
                    isValid = value.trim().length > 0;
                    break;
                case 'number':
                    isValid = !isNaN(Number(value)) && isFinite(Number(value));
                    break;
                case 'text':
                    isValid = typeof value === 'string';
                    break;
                case 'regex':
                    isValid = rule.pattern ? rule.pattern.test(value) : true;
                    break;
                case 'custom':
                    isValid = rule.validate ? rule.validate(value) : true;
                    break;
            }

            if (!isValid) {
                validation.isValid = false;
                validation.errors.push(rule.message);
            }
        }

        return validation;
    }

    static getCellKey(row: number, col: number): string {
        return `${row},${col}`;
    }

    static createCellState(originalValue: SheetCellValue | null, currentValue: SheetCellValue | null): CellState {
        return {
            originalValue,
            currentValue,
            isModified: JSON.stringify(originalValue) !== JSON.stringify(currentValue),
            isValid: true,
            lastModified: new Date(),
            originalFormat: originalValue?.s || null,
            currentFormat: currentValue?.s || null,
            isFormatModified: JSON.stringify(originalValue?.s) !== JSON.stringify(currentValue?.s)
        };
    }

    // Formatting utility methods
    static formatNumberValue(value: number | string, format: NumberFormat): string {
        const numValue = typeof value === 'string' ? parseFloat(value) : value;
        if (isNaN(numValue)) return String(value);

        switch (format.type) {
            case 'number':
                return this.formatNumber(numValue, format.decimalPlaces || 2, format.thousandsSeparator || false);
            case 'currency':
                return this.formatCurrency(numValue, format.currencySymbol || '$', format.decimalPlaces || 2);
            case 'percentage':
                return this.formatPercentage(numValue, format.decimalPlaces || 2);
            case 'scientific':
                return numValue.toExponential(format.decimalPlaces || 2);
            case 'custom':
                return format.customFormat ? this.applyCustomFormat(numValue, format.customFormat) : String(numValue);
            default:
                return String(value);
        }
    }

    static formatDateValue(value: string | number | Date, format: NumberFormat): string {
        let date: Date;
        
        if (value instanceof Date) {
            date = value;
        } else if (typeof value === 'number') {
            date = new Date(value);
        } else {
            date = new Date(value);
        }

        if (isNaN(date.getTime())) return String(value);

        switch (format.type) {
            case 'date':
                return this.formatDate(date, format.dateFormat || 'MM/dd/yyyy');
            case 'time':
                return this.formatTime(date, format.timeFormat || 'HH:mm:ss');
            default:
                return date.toLocaleDateString();
        }
    }

    private static formatNumber(value: number, decimalPlaces: number, useThousandsSeparator: boolean): string {
        const formatted = value.toFixed(decimalPlaces);
        if (!useThousandsSeparator) return formatted;
        
        const parts = formatted.split('.');
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        return parts.join('.');
    }

    private static formatCurrency(value: number, symbol: string, decimalPlaces: number): string {
        const formatted = this.formatNumber(Math.abs(value), decimalPlaces, true);
        const prefix = value < 0 ? '-' : '';
        return `${prefix}${symbol}${formatted}`;
    }

    private static formatPercentage(value: number, decimalPlaces: number): string {
        return (value * 100).toFixed(decimalPlaces) + '%';
    }

    private static formatDate(date: Date, format: string): string {
        const formatMap: Record<string, string> = {
            'yyyy': date.getFullYear().toString(),
            'yy': date.getFullYear().toString().slice(-2),
            'MM': (date.getMonth() + 1).toString().padStart(2, '0'),
            'M': (date.getMonth() + 1).toString(),
            'dd': date.getDate().toString().padStart(2, '0'),
            'd': date.getDate().toString(),
            'EEEE': date.toLocaleDateString('en-US', { weekday: 'long' }),
            'EEE': date.toLocaleDateString('en-US', { weekday: 'short' })
        };

        let result = format;
        for (const [pattern, replacement] of Object.entries(formatMap)) {
            result = result.replace(new RegExp(pattern, 'g'), replacement);
        }
        return result;
    }

    private static formatTime(date: Date, format: string): string {
        const formatMap: Record<string, string> = {
            'HH': date.getHours().toString().padStart(2, '0'),
            'H': date.getHours().toString(),
            'hh': (date.getHours() % 12 || 12).toString().padStart(2, '0'),
            'h': (date.getHours() % 12 || 12).toString(),
            'mm': date.getMinutes().toString().padStart(2, '0'),
            'm': date.getMinutes().toString(),
            'ss': date.getSeconds().toString().padStart(2, '0'),
            's': date.getSeconds().toString(),
            'a': date.getHours() >= 12 ? 'PM' : 'AM'
        };

        let result = format;
        for (const [pattern, replacement] of Object.entries(formatMap)) {
            result = result.replace(new RegExp(pattern, 'g'), replacement);
        }
        return result;
    }

    private static applyCustomFormat(value: number, format: string): string {
        // Basic custom format support - can be extended
        return format.replace('#', value.toString());
    }

    static getFormattedCellValue(cellValue: SheetCellValue | null | undefined): string {
        if (!cellValue) return '';
        
        // First check if there's manual formatting (m property)
        if (cellValue.m) return cellValue.m;
        
        // If there's a value and formatting, apply formatting
        if (cellValue.v !== undefined && cellValue.s?.numberFormat) {
            if (typeof cellValue.v === 'number') {
                return this.formatNumberValue(cellValue.v, cellValue.s.numberFormat);
            } else if (cellValue.s.numberFormat.type === 'date' || cellValue.s.numberFormat.type === 'time') {
                return this.formatDateValue(cellValue.v, cellValue.s.numberFormat);
            }
        }
        
        // Default to string conversion
        if (cellValue.v !== undefined) return String(cellValue.v);
        
        return '';
    }

    static applyCellFormat(cellValue: SheetCellValue, format: CellFormat): SheetCellValue {
        return {
            ...cellValue,
            s: { ...cellValue.s, ...format }
        };
    }

    static mergeCellFormats(baseFormat: CellFormat | undefined, newFormat: CellFormat): CellFormat {
        return {
            ...baseFormat,
            ...newFormat,
            border: {
                ...baseFormat?.border,
                ...newFormat.border
            }
        };
    }

    static generateCellStyle(format: CellFormat | undefined): string {
        if (!format) return '';
        
        const styles: string[] = [];
        
        // Font properties
        if (format.fontFamily) styles.push(`font-family: ${format.fontFamily}`);
        if (format.fontSize) styles.push(`font-size: ${format.fontSize}px`);
        if (format.fontWeight) styles.push(`font-weight: ${format.fontWeight}`);
        if (format.fontStyle) styles.push(`font-style: ${format.fontStyle}`);
        if (format.textDecoration) styles.push(`text-decoration: ${format.textDecoration}`);
        
        // Text properties
        if (format.color) styles.push(`color: ${format.color}`);
        if (format.textAlign) styles.push(`text-align: ${format.textAlign}`);
        if (format.verticalAlign) styles.push(`vertical-align: ${format.verticalAlign}`);
        if (format.textWrap !== undefined) {
            styles.push(`white-space: ${format.textWrap ? 'pre-wrap' : 'nowrap'}`);
        }
        
        // Background
        if (format.backgroundColor) styles.push(`background-color: ${format.backgroundColor}`);
        
        // Borders
        if (format.border) {
            const borderStyles = this.generateBorderStyles(format.border);
            styles.push(...borderStyles);
        }
        
        // Custom CSS
        if (format.customCss) styles.push(format.customCss);
        
        return styles.join('; ');
    }

    private static generateBorderStyles(border: CellBorder): string[] {
        const styles: string[] = [];
        
        if (border.all) {
            const borderStr = this.formatBorderStyle(border.all);
            styles.push(`border: ${borderStr}`);
        } else {
            if (border.top) styles.push(`border-top: ${this.formatBorderStyle(border.top)}`);
            if (border.right) styles.push(`border-right: ${this.formatBorderStyle(border.right)}`);
            if (border.bottom) styles.push(`border-bottom: ${this.formatBorderStyle(border.bottom)}`);
            if (border.left) styles.push(`border-left: ${this.formatBorderStyle(border.left)}`);
        }
        
        return styles;
    }

    private static formatBorderStyle(border: BorderStyle): string {
        const width = border.width || 1;
        const style = border.style || 'solid';
        const color = border.color || '#000000';
        return `${width}px ${style} ${color}`;
    }
}

export class DataManager {
    private stateManager: SheetStateManager;
    private sheetData: SheetData[];
    private autoSaveTimer: number | null = null;
    private debounceTimer: number | null = null;
    private saveCallback: (() => Promise<void>) | null = null;

    constructor(sheetData: SheetData[], maxUndoSteps: number = 100, autoSaveConfig?: Partial<AutoSaveConfig>) {
        this.sheetData = sheetData;
        this.stateManager = {
            undoStack: [],
            redoStack: [],
            cellStates: new Map(),
            validationRules: new Map(),
            isDirty: false,
            maxUndoSteps,
            autoSaveConfig: {
                enabled: true,
                intervalMs: 30000, // 30 seconds
                maxRetries: 3,
                debounceMs: 2000, // 2 seconds after last change
                ...autoSaveConfig
            },
            autoSaveState: {
                lastSaved: null,
                isSaving: false,
                lastError: null,
                saveCount: 0,
                retryCount: 0
            }
        };
        this.initializeFormulas();
        this.initializeCellStates();
        this.startAutoSave();
    }

    private initializeFormulas(): void {
        for (const sheet of this.sheetData) {
            if (!sheet.formulas) {
                sheet.formulas = new Map();
            }
            
            // Scan for existing formulas
            for (const cell of sheet.celldata) {
                if (cell.v.f) {
                    const key = `${cell.r},${cell.c}`;
                    const formulaCell: FormulaCell = {
                        row: cell.r,
                        col: cell.c,
                        formula: cell.v.f,
                        dependencies: [],
                        dependents: []
                    };
                    sheet.formulas.set(key, formulaCell);
                }
            }
            
            // Calculate dependencies
            this.calculateAllDependencies(sheet);
        }
    }

    private calculateAllDependencies(sheet: SheetData): void {
        if (!sheet.formulas) return;

        // Clear existing dependencies
        for (const formulaCell of sheet.formulas.values()) {
            formulaCell.dependencies = [];
            formulaCell.dependents = [];
        }

        // Calculate new dependencies
        for (const formulaCell of sheet.formulas.values()) {
            const context = this.createFormulaContext(sheet, formulaCell.row, formulaCell.col);
            const result = FormulaEngine.parseFormula(formulaCell.formula, context);
            
            formulaCell.dependencies = result.dependencies.map(dep => `${dep.row},${dep.col}`);
            
            // Update dependents
            for (const depKey of formulaCell.dependencies) {
                const depFormula = sheet.formulas.get(depKey);
                if (depFormula) {
                    const currentCellKey = `${formulaCell.row},${formulaCell.col}`;
                    if (!depFormula.dependents.includes(currentCellKey)) {
                        depFormula.dependents.push(currentCellKey);
                    }
                }
            }
        }
    }

    private createFormulaContext(sheet: SheetData, currentRow: number, currentCol: number): FormulaContext {
        return {
            getCellValue: (ref: CellReference) => {
                const cell = SheetParser.getCellAt(sheet, ref.row, ref.col);
                if (!cell) return '';
                
                // If it's a formula cell, return the calculated value
                if (cell.v.f) {
                    return cell.v.v; // The calculated value
                }
                
                return SheetParser.getCellValue(cell.v);
            },
            getRangeValues: (range) => {
                const values = [];
                for (let r = range.start.row; r <= range.end.row; r++) {
                    for (let c = range.start.col; c <= range.end.col; c++) {
                        const cell = SheetParser.getCellAt(sheet, r, c);
                        values.push(cell ? SheetParser.getCellValue(cell.v) : '');
                    }
                }
                return values;
            },
            currentCell: { row: currentRow, col: currentCol }
        };
    }

    private initializeCellStates(): void {
        for (const sheet of this.sheetData) {
            for (const cell of sheet.celldata) {
                const key = `${sheet.id}_${SheetParser.getCellKey(cell.r, cell.c)}`;
                const state = SheetParser.createCellState(cell.v, cell.v);
                this.stateManager.cellStates.set(key, state);
            }
        }
    }

    updateCell(sheetId: string, row: number, col: number, newValue: string, validate: boolean = true): boolean {
        return this.updateCellWithFormat(sheetId, row, col, newValue, undefined, validate);
    }

    updateCellWithFormat(sheetId: string, row: number, col: number, newValue: string, format?: CellFormat, validate: boolean = true): boolean {
        const sheet = this.sheetData.find(s => s.id === sheetId);
        if (!sheet) return false;

        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        const currentCell = SheetParser.getCellAt(sheet, row, col);
        const oldValue = currentCell?.v || null;

        // Validate if required
        if (validate) {
            const rules = this.stateManager.validationRules.get(key) || [];
            const validation = SheetParser.validateCell(newValue, rules);
            if (!validation.isValid) {
                const state = this.stateManager.cellStates.get(key);
                if (state) {
                    state.isValid = false;
                    state.validationError = validation.errors.join(', ');
                }
                return false;
            }
        }

        // Create undo action
        const undoAction: UndoRedoAction = {
            type: 'cell_edit',
            sheetId,
            timestamp: new Date(),
            data: {
                row,
                col,
                oldValue,
                newValue: newValue.trim() === '' ? null : this.createCellValue(newValue)
            }
        };

        // Handle formula vs regular value
        if (newValue.startsWith('=')) {
            this.updateFormulaCell(sheet, row, col, newValue);
        } else {
            this.updateRegularCell(sheet, row, col, newValue);
        }

        // Update state
        let newCellValue = newValue.trim() === '' ? null : this.createCellValue(newValue);
        
        // Apply formatting if provided
        if (format && newCellValue) {
            newCellValue = SheetParser.applyCellFormat(newCellValue, format);
        }
        
        const state = this.stateManager.cellStates.get(key) || SheetParser.createCellState(oldValue, newCellValue);
        state.currentValue = newCellValue;
        state.isModified = JSON.stringify(state.originalValue) !== JSON.stringify(newCellValue);
        state.currentFormat = newCellValue?.s || null;
        state.isFormatModified = JSON.stringify(state.originalFormat) !== JSON.stringify(state.currentFormat);
        state.isValid = true;
        state.validationError = undefined;
        state.lastModified = new Date();
        this.stateManager.cellStates.set(key, state);

        // Recalculate dependent formulas
        this.recalculateDependents(sheet, row, col);

        // Add to undo stack
        this.addUndoAction(undoAction);
        this.stateManager.isDirty = true;

        // Trigger auto-save debounce
        this.scheduleAutoSave();

        return true;
    }

    private createCellValue(value: string, format?: CellFormat): SheetCellValue {
        if (value.startsWith('=')) {
            return {
                f: value,
                m: value,
                v: '', // Will be calculated
                calculated: true,
                s: format
            };
        } else {
            const parsedValue = SheetParser.parseValue(value);
            return {
                v: typeof parsedValue === 'boolean' ? String(parsedValue) : parsedValue,
                m: value,
                s: format
            };
        }
    }

    private updateFormulaCell(sheet: SheetData, row: number, col: number, formula: string): void {
        if (!sheet.formulas) {
            sheet.formulas = new Map();
        }

        const cellKey = `${row},${col}`;
        
        // Remove old formula if exists
        if (sheet.formulas.has(cellKey)) {
            this.removeFormulaDependencies(sheet, row, col);
        }

        // Calculate the formula
        const context = this.createFormulaContext(sheet, row, col);
        const result = FormulaEngine.parseFormula(formula, context);

        // Update cell data
        SheetParser.updateCellValue(sheet, row, col, String(result.value));
        const cell = SheetParser.getCellAt(sheet, row, col);
        if (cell) {
            cell.v.f = formula;
            cell.v.calculated = true;
            cell.v.v = result.value;
        }

        // Store formula metadata
        const formulaCell: FormulaCell = {
            row,
            col,
            formula,
            dependencies: result.dependencies.map(dep => `${dep.row},${dep.col}`),
            dependents: []
        };
        sheet.formulas.set(cellKey, formulaCell);

        // Update dependencies
        this.updateFormulaDependencies(sheet, formulaCell);
    }

    private updateRegularCell(sheet: SheetData, row: number, col: number, value: string): void {
        const cellKey = `${row},${col}`;
        
        // Remove formula if this cell had one
        if (sheet.formulas?.has(cellKey)) {
            this.removeFormulaDependencies(sheet, row, col);
            sheet.formulas.delete(cellKey);
        }

        // Update cell value
        SheetParser.updateCellValue(sheet, row, col, value);
    }

    private updateFormulaDependencies(sheet: SheetData, formulaCell: FormulaCell): void {
        const currentCellKey = `${formulaCell.row},${formulaCell.col}`;
        
        // Add this cell as dependent of its dependencies
        for (const depKey of formulaCell.dependencies) {
            const depFormula = sheet.formulas?.get(depKey);
            if (depFormula && !depFormula.dependents.includes(currentCellKey)) {
                depFormula.dependents.push(currentCellKey);
            }
        }
    }

    private removeFormulaDependencies(sheet: SheetData, row: number, col: number): void {
        const cellKey = `${row},${col}`;
        const formulaCell = sheet.formulas?.get(cellKey);
        
        if (!formulaCell) return;

        // Remove this cell from dependents of its dependencies
        for (const depKey of formulaCell.dependencies) {
            const depFormula = sheet.formulas?.get(depKey);
            if (depFormula) {
                depFormula.dependents = depFormula.dependents.filter(dep => dep !== cellKey);
            }
        }
    }

    private recalculateDependents(sheet: SheetData, row: number, col: number): void {
        const cellKey = `${row},${col}`;
        const formulaCell = sheet.formulas?.get(cellKey);
        
        if (!formulaCell) return;

        // Recalculate all dependent formulas
        const toRecalculate = [...formulaCell.dependents];
        const calculated = new Set<string>();

        while (toRecalculate.length > 0) {
            const depKey = toRecalculate.shift()!;
            if (calculated.has(depKey)) continue;

            const [depRowStr, depColStr] = depKey.split(',');
            const depRow = parseInt(depRowStr);
            const depCol = parseInt(depColStr);
            
            const depFormula = sheet.formulas?.get(depKey);
            if (depFormula) {
                // Recalculate this formula
                const context = this.createFormulaContext(sheet, depRow, depCol);
                const result = FormulaEngine.parseFormula(depFormula.formula, context);
                
                // Update cell value
                const cell = SheetParser.getCellAt(sheet, depRow, depCol);
                if (cell) {
                    cell.v.v = result.value;
                }

                // Add its dependents to the queue
                toRecalculate.push(...depFormula.dependents);
                calculated.add(depKey);
            }
        }
    }

    private addUndoAction(action: UndoRedoAction): void {
        this.stateManager.undoStack.push(action);
        this.stateManager.redoStack = []; // Clear redo stack on new action
        
        // Limit undo stack size
        if (this.stateManager.undoStack.length > this.stateManager.maxUndoSteps) {
            this.stateManager.undoStack.shift();
        }
    }

    undo(): boolean {
        const action = this.stateManager.undoStack.pop();
        if (!action) return false;

        const sheet = this.sheetData.find(s => s.id === action.sheetId);
        if (!sheet) return false;

        switch (action.type) {
            case 'cell_edit':
                if (action.data.row !== undefined && action.data.col !== undefined) {
                    const oldValueStr = action.data.oldValue ? SheetParser.getCellValue(action.data.oldValue) : '';
                    SheetParser.updateCellValue(sheet, action.data.row, action.data.col, oldValueStr);
                    
                    const key = `${action.sheetId}_${SheetParser.getCellKey(action.data.row, action.data.col)}`;
                    const state = this.stateManager.cellStates.get(key);
                    if (state) {
                        state.currentValue = action.data.oldValue || null;
                        state.isModified = JSON.stringify(state.originalValue) !== JSON.stringify(action.data.oldValue || null);
                        state.lastModified = new Date();
                    }
                }
                break;
        }

        this.stateManager.redoStack.push(action);
        this.updateDirtyState();
        return true;
    }

    redo(): boolean {
        const action = this.stateManager.redoStack.pop();
        if (!action) return false;

        const sheet = this.sheetData.find(s => s.id === action.sheetId);
        if (!sheet) return false;

        switch (action.type) {
            case 'cell_edit':
                if (action.data.row !== undefined && action.data.col !== undefined && action.data.newValue) {
                    const newValueStr = SheetParser.getCellValue(action.data.newValue);
                    SheetParser.updateCellValue(sheet, action.data.row, action.data.col, newValueStr);
                    
                    const key = `${action.sheetId}_${SheetParser.getCellKey(action.data.row, action.data.col)}`;
                    const state = this.stateManager.cellStates.get(key);
                    if (state) {
                        state.currentValue = action.data.newValue;
                        state.isModified = JSON.stringify(state.originalValue) !== JSON.stringify(action.data.newValue);
                        state.lastModified = new Date();
                    }
                }
                break;
        }

        this.stateManager.undoStack.push(action);
        this.updateDirtyState();
        return true;
    }

    private updateDirtyState(): void {
        let hasDirtyChanges = false;
        for (const state of this.stateManager.cellStates.values()) {
            if (state.isModified) {
                hasDirtyChanges = true;
                break;
            }
        }
        this.stateManager.isDirty = hasDirtyChanges;
    }

    getCellState(sheetId: string, row: number, col: number): CellState | null {
        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        return this.stateManager.cellStates.get(key) || null;
    }

    addValidationRule(sheetId: string, row: number, col: number, rule: ValidationRule): void {
        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        const rules = this.stateManager.validationRules.get(key) || [];
        rules.push(rule);
        this.stateManager.validationRules.set(key, rules);
    }

    removeValidationRule(sheetId: string, row: number, col: number, ruleType: string): void {
        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        const rules = this.stateManager.validationRules.get(key) || [];
        const filteredRules = rules.filter(rule => rule.type !== ruleType);
        this.stateManager.validationRules.set(key, filteredRules);
    }

    getValidationRules(sheetId: string, row: number, col: number): ValidationRule[] {
        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        return this.stateManager.validationRules.get(key) || [];
    }

    markClean(): void {
        for (const state of this.stateManager.cellStates.values()) {
            state.originalValue = state.currentValue;
            state.isModified = false;
        }
        this.stateManager.isDirty = false;
    }

    isDirty(): boolean {
        return this.stateManager.isDirty;
    }

    canUndo(): boolean {
        return this.stateManager.undoStack.length > 0;
    }

    canRedo(): boolean {
        return this.stateManager.redoStack.length > 0;
    }

    getModifiedCells(): Array<{sheetId: string, row: number, col: number, state: CellState}> {
        const modified: Array<{sheetId: string, row: number, col: number, state: CellState}> = [];
        
        for (const [key, state] of this.stateManager.cellStates.entries()) {
            if (state.isModified) {
                const [sheetId, coords] = key.split('_');
                const [row, col] = coords.split(',').map(Number);
                modified.push({ sheetId, row, col, state });
            }
        }
        
        return modified;
    }

    getSheetData(): SheetData[] {
        return this.sheetData;
    }

    // Cell formatting methods
    formatCell(sheetId: string, row: number, col: number, format: CellFormat): boolean {
        const sheet = this.sheetData.find(s => s.id === sheetId);
        if (!sheet) return false;

        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        const currentCell = SheetParser.getCellAt(sheet, row, col);

        // Create undo action for formatting
        const undoAction: UndoRedoAction = {
            type: 'cell_edit',
            sheetId,
            timestamp: new Date(),
            data: {
                row,
                col,
                oldValue: currentCell?.v || null,
                newValue: currentCell?.v || null // Value doesn't change, only format
            }
        };

        // Apply format to existing cell or create new cell
        if (currentCell) {
            currentCell.v.s = SheetParser.mergeCellFormats(currentCell.v.s, format);
        } else {
            // Create new cell with just formatting
            const newCellValue: SheetCellValue = {
                v: '',
                m: '',
                s: format
            };
            sheet.celldata.push({
                r: row,
                c: col,
                v: newCellValue
            });
        }

        // Update state
        const state = this.stateManager.cellStates.get(key) || SheetParser.createCellState(null, null);
        const updatedCell = SheetParser.getCellAt(sheet, row, col);
        state.currentValue = updatedCell?.v || null;
        state.currentFormat = updatedCell?.v.s || null;
        state.isFormatModified = JSON.stringify(state.originalFormat) !== JSON.stringify(state.currentFormat);
        state.lastModified = new Date();
        this.stateManager.cellStates.set(key, state);

        // Add to undo stack
        this.addUndoAction(undoAction);
        this.stateManager.isDirty = true;

        // Trigger auto-save debounce
        this.scheduleAutoSave();

        return true;
    }

    clearCellFormat(sheetId: string, row: number, col: number): boolean {
        const sheet = this.sheetData.find(s => s.id === sheetId);
        if (!sheet) return false;

        const currentCell = SheetParser.getCellAt(sheet, row, col);
        if (!currentCell) return false;

        const key = `${sheetId}_${SheetParser.getCellKey(row, col)}`;
        
        // Create undo action
        const undoAction: UndoRedoAction = {
            type: 'cell_edit',
            sheetId,
            timestamp: new Date(),
            data: {
                row,
                col,
                oldValue: currentCell.v,
                newValue: { ...currentCell.v, s: undefined }
            }
        };

        // Clear formatting
        delete currentCell.v.s;

        // Update state
        const state = this.stateManager.cellStates.get(key);
        if (state) {
            state.currentFormat = null;
            state.isFormatModified = JSON.stringify(state.originalFormat) !== JSON.stringify(null);
            state.lastModified = new Date();
        }

        // Add to undo stack
        this.addUndoAction(undoAction);
        this.stateManager.isDirty = true;
        this.scheduleAutoSave();

        return true;
    }

    formatRange(sheetId: string, startRow: number, startCol: number, endRow: number, endCol: number, format: CellFormat): boolean {
        let success = true;
        for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
                if (!this.formatCell(sheetId, row, col, format)) {
                    success = false;
                }
            }
        }
        return success;
    }

    getCellFormat(sheetId: string, row: number, col: number): CellFormat | null {
        const sheet = this.sheetData.find(s => s.id === sheetId);
        if (!sheet) return null;

        const cell = SheetParser.getCellAt(sheet, row, col);
        return cell?.v.s || null;
    }

    copyFormat(fromSheetId: string, fromRow: number, fromCol: number, toSheetId: string, toRow: number, toCol: number): boolean {
        const sourceFormat = this.getCellFormat(fromSheetId, fromRow, fromCol);
        if (!sourceFormat) return false;

        return this.formatCell(toSheetId, toRow, toCol, sourceFormat);
    }

    // Conflict detection and resolution
    generateFileChecksum(content: string): string {
        // Simple hash function for file content
        let hash = 0;
        for (let i = 0; i < content.length; i++) {
            const char = content.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return hash.toString(16);
    }

    detectConflicts(currentFileContent: string, fileModifiedTime?: Date): ConflictInfo | null {
        const currentChecksum = this.generateFileChecksum(currentFileContent);
        const hasFileChanged = this.stateManager.lastFileChecksum && 
                               this.stateManager.lastFileChecksum !== currentChecksum;
        
        const hasTimeChanged = fileModifiedTime && 
                              this.stateManager.lastFileModified && 
                              fileModifiedTime > this.stateManager.lastFileModified;

        if (hasFileChanged || hasTimeChanged) {
            // Parse the current file content to compare with local changes
            const remoteSheetData = SheetParser.parse(currentFileContent);
            const conflictedCells = this.findConflictedCells(remoteSheetData);

            if (conflictedCells.length > 0) {
                const conflict: ConflictInfo = {
                    type: hasTimeChanged ? 'file_modified' : 'version_mismatch',
                    conflictId: `conflict_${Date.now()}`,
                    timestamp: new Date(),
                    conflictedCells,
                    fileChecksum: currentChecksum,
                    lastKnownModified: fileModifiedTime
                };

                this.stateManager.conflictInfo = conflict;
                return conflict;
            }
        }

        // Update tracking info even if no conflicts
        this.stateManager.lastFileChecksum = currentChecksum;
        this.stateManager.lastFileModified = fileModifiedTime || new Date();

        return null;
    }

    private findConflictedCells(remoteSheetData: SheetData[]): Array<{
        row: number;
        col: number;
        localValue: SheetCellValue | null;
        remoteValue: SheetCellValue | null;
        baseValue?: SheetCellValue | null;
    }> {
        const conflicts: Array<{
            row: number;
            col: number;
            localValue: SheetCellValue | null;
            remoteValue: SheetCellValue | null;
            baseValue?: SheetCellValue | null;
        }> = [];

        const modifiedCells = this.getModifiedCells();

        for (const { sheetId, row, col, state } of modifiedCells) {
            // Find corresponding sheet in remote data
            const remoteSheet = remoteSheetData.find(s => s.id === sheetId);
            if (!remoteSheet) continue;

            const remoteCell = SheetParser.getCellAt(remoteSheet, row, col);
            const remoteCellValue = remoteCell ? remoteCell.v : null;
            const localCellValue = state.currentValue;
            const baseCellValue = state.originalValue;

            // Check if remote value differs from both local and base values
            const remoteValueStr = JSON.stringify(remoteCellValue);
            const localValueStr = JSON.stringify(localCellValue);
            const baseValueStr = JSON.stringify(baseCellValue);

            if (remoteValueStr !== baseValueStr && remoteValueStr !== localValueStr) {
                conflicts.push({
                    row,
                    col,
                    localValue: localCellValue,
                    remoteValue: remoteCellValue,
                    baseValue: baseCellValue
                });
            }
        }

        return conflicts;
    }

    async resolveConflicts(resolution: ConflictResolution): Promise<boolean> {
        if (!this.stateManager.conflictInfo) {
            return false;
        }

        try {
            const conflict = this.stateManager.conflictInfo;
            
            for (const resolvedCell of resolution.resolvedCells) {
                const { row, col, resolvedValue } = resolvedCell;
                
                // Find the corresponding sheet
                const sheet = this.sheetData.find(s => 
                    conflict.conflictedCells.some(cc => cc.row === row && cc.col === col)
                );
                
                if (!sheet) continue;

                // Apply the resolution
                if (resolvedValue) {
                    const cellValue = typeof resolvedValue.v === 'string' ? resolvedValue.v : String(resolvedValue.v || '');
                    this.updateCell(sheet.id, row, col, cellValue);
                } else {
                    // Clear the cell
                    this.updateCell(sheet.id, row, col, '');
                }
            }

            // Clear conflict info
            this.stateManager.conflictInfo = undefined;
            
            return true;
        } catch (error) {
            console.error('Failed to resolve conflicts:', error);
            return false;
        }
    }

    hasActiveConflicts(): boolean {
        return !!this.stateManager.conflictInfo;
    }

    getActiveConflict(): ConflictInfo | null {
        return this.stateManager.conflictInfo || null;
    }

    createBackup(): string {
        return JSON.stringify({
            timestamp: new Date().toISOString(),
            sheetData: this.sheetData,
            cellStates: Array.from(this.stateManager.cellStates.entries()),
            undoStack: this.stateManager.undoStack,
            checksum: this.stateManager.lastFileChecksum
        });
    }

    restoreFromBackup(backupData: string): boolean {
        try {
            const backup = JSON.parse(backupData);
            
            this.sheetData = backup.sheetData;
            this.stateManager.cellStates = new Map(backup.cellStates);
            this.stateManager.undoStack = backup.undoStack || [];
            this.stateManager.redoStack = []; // Clear redo stack after restore
            this.stateManager.lastFileChecksum = backup.checksum;
            
            return true;
        } catch (error) {
            console.error('Failed to restore from backup:', error);
            return false;
        }
    }

    // Auto-save functionality
    setSaveCallback(callback: () => Promise<void>): void {
        this.saveCallback = callback;
    }

    private startAutoSave(): void {
        if (!this.stateManager.autoSaveConfig.enabled) return;

        this.autoSaveTimer = window.setInterval(() => {
            this.performAutoSave();
        }, this.stateManager.autoSaveConfig.intervalMs);
    }

    private scheduleAutoSave(): void {
        if (!this.stateManager.autoSaveConfig.enabled || !this.saveCallback) return;

        // Clear existing debounce timer
        if (this.debounceTimer) {
            clearTimeout(this.debounceTimer);
        }

        // Schedule auto-save after debounce period
        this.debounceTimer = window.setTimeout(() => {
            this.performAutoSave();
        }, this.stateManager.autoSaveConfig.debounceMs);
    }

    private async performAutoSave(): Promise<void> {
        if (!this.stateManager.isDirty || 
            this.stateManager.autoSaveState.isSaving || 
            !this.saveCallback) return;

        this.stateManager.autoSaveState.isSaving = true;
        this.stateManager.autoSaveState.lastError = null;

        try {
            await this.saveCallback();
            
            // Success
            this.stateManager.autoSaveState.lastSaved = new Date();
            this.stateManager.autoSaveState.saveCount++;
            this.stateManager.autoSaveState.retryCount = 0;
            
            console.log('Auto-save successful');
            
        } catch (error) {
            console.error('Auto-save failed:', error);
            
            this.stateManager.autoSaveState.lastError = error instanceof Error ? error.message : 'Unknown error';
            this.stateManager.autoSaveState.retryCount++;
            
            // Retry logic
            if (this.stateManager.autoSaveState.retryCount < this.stateManager.autoSaveConfig.maxRetries) {
                const retryDelay = Math.min(1000 * Math.pow(2, this.stateManager.autoSaveState.retryCount), 10000);
                setTimeout(() => this.performAutoSave(), retryDelay);
            }
        } finally {
            this.stateManager.autoSaveState.isSaving = false;
        }
    }

    enableAutoSave(enabled: boolean = true): void {
        this.stateManager.autoSaveConfig.enabled = enabled;
        
        if (enabled) {
            this.startAutoSave();
        } else {
            this.stopAutoSave();
        }
    }

    private stopAutoSave(): void {
        if (this.autoSaveTimer) {
            clearInterval(this.autoSaveTimer);
            this.autoSaveTimer = null;
        }
        
        if (this.debounceTimer) {
            clearTimeout(this.debounceTimer);
            this.debounceTimer = null;
        }
    }

    getAutoSaveState(): AutoSaveState {
        return { ...this.stateManager.autoSaveState };
    }

    getAutoSaveConfig(): AutoSaveConfig {
        return { ...this.stateManager.autoSaveConfig };
    }

    updateAutoSaveConfig(config: Partial<AutoSaveConfig>): void {
        Object.assign(this.stateManager.autoSaveConfig, config);
        
        // Restart auto-save with new config
        if (this.stateManager.autoSaveConfig.enabled) {
            this.stopAutoSave();
            this.startAutoSave();
        }
    }

    forceAutoSave(): Promise<void> {
        return this.performAutoSave();
    }

    getTimeSinceLastSave(): number | null {
        if (!this.stateManager.autoSaveState.lastSaved) return null;
        return Date.now() - this.stateManager.autoSaveState.lastSaved.getTime();
    }

    // Clean up on destroy
    destroy(): void {
        this.stopAutoSave();
    }
}