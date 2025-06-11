import { TextFileView, TFile, WorkspaceLeaf } from 'obsidian';
import { SheetParser, SheetData, DataManager, ValidationRule, AutoSaveState, ConflictInfo, ConflictResolution } from './sheetParser';
import { FormulaEngine } from './formulaEngine';

export const VIEW_TYPE_SHEET = 'sheet-view';

export class SheetView extends TextFileView {
    private sheetData: SheetData[] = [];
    private dataManager: DataManager | null = null;
    private isEditing: boolean = false;
    private currentEditor: HTMLInputElement | HTMLTextAreaElement | null = null;
    private editingCell: { sheetIndex: number; row: number; col: number } | null = null;
    private selectedCell: { sheetIndex: number; row: number; col: number } | null = null;
    private touchStartPos: { x: number; y: number } | null = null;
    private longPressTimer: number | null = null;
    private isLongPress: boolean = false;
    private swipeThreshold: number = 50;
    private longPressDelay: number = 500;
    private gridDimensions: { rows: number; cols: number } = { rows: 0, cols: 0 };
    private fileWatcher: number | null = null;
    private lastFileContent: string = '';
    private conflictCheckInterval: number = 30000; // Check every 30 seconds

    constructor(leaf: WorkspaceLeaf) {
        super(leaf);
        console.log('SheetView: Constructor called');
        this.setupOrientationHandler();
        this.setupKeyboardHandlers();
    }

    getViewType(): string {
        return VIEW_TYPE_SHEET;
    }

    getDisplayText(): string {
        const baseName = this.file?.basename || 'Sheet';
        return this.dataManager?.isDirty() ? `${baseName} •` : baseName;
    }

    getViewData(): string {
        return this.data;
    }

    setViewData(data: string, clear: boolean): void {
        console.log('SheetView: setViewData called with data:', data);
        this.data = data;
        this.lastFileContent = data;
        this.sheetData = SheetParser.parse(data);
        this.dataManager = new DataManager(this.sheetData, 100, {
            enabled: true,
            intervalMs: 30000, // 30 seconds
            debounceMs: 2000,  // 2 seconds after last change
            maxRetries: 3
        });
        
        // Set up auto-save callback
        this.dataManager.setSaveCallback(() => this.performSave());
        
        // Initialize conflict detection
        this.initializeConflictDetection();
        
        console.log('SheetView: parsed sheet data:', this.sheetData);
        this.render();
    }

    clear(): void {
        this.data = '';
        this.sheetData = [];
        if (this.dataManager) {
            this.dataManager.destroy();
            this.dataManager = null;
        }
        this.stopConflictDetection();
        this.contentEl.empty();
    }

    private render(): void {
        console.log('SheetView: render called, sheetData:', this.sheetData);
        this.contentEl.empty();
        this.contentEl.addClass('mobile-sheet-viewer');

        this.renderToolbar();
        this.setupSwipeGestures();

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

    private renderToolbar(): void {
        const toolbar = this.contentEl.createEl('div', { cls: 'sheet-toolbar' });
        
        const saveBtn = toolbar.createEl('button', { 
            text: 'Save',
            cls: 'sheet-save-btn'
        });
        saveBtn.disabled = !this.dataManager?.isDirty();
        saveBtn.addEventListener('click', () => this.saveFile());

        const editToggle = toolbar.createEl('button', { 
            text: this.isEditing ? 'View Mode' : 'Edit Mode',
            cls: 'sheet-edit-toggle'
        });
        editToggle.addEventListener('click', () => this.toggleEditMode());

        // Undo/Redo buttons
        const undoBtn = toolbar.createEl('button', { 
            text: '↶ Undo',
            cls: 'sheet-undo-btn'
        });
        undoBtn.disabled = !this.dataManager?.canUndo();
        undoBtn.addEventListener('click', () => this.undo());

        const redoBtn = toolbar.createEl('button', { 
            text: '↷ Redo',
            cls: 'sheet-redo-btn'
        });
        redoBtn.disabled = !this.dataManager?.canRedo();
        redoBtn.addEventListener('click', () => this.redo());

        // Auto-save status and modified cells indicator
        if (this.dataManager) {
            const autoSaveState = this.dataManager.getAutoSaveState();
            const isDirty = this.dataManager.isDirty();
            const hasConflicts = this.dataManager.hasActiveConflicts();
            
            if (hasConflicts) {
                const conflictBtn = toolbar.createEl('button', {
                    text: 'Conflicts',
                    cls: 'sheet-conflict-indicator'
                });
                conflictBtn.addEventListener('click', () => this.showConflictResolutionDialog());
            } else if (autoSaveState.isSaving) {
                toolbar.createEl('span', {
                    text: 'Saving...',
                    cls: 'sheet-auto-save-indicator saving'
                });
            } else if (autoSaveState.lastError) {
                const errorSpan = toolbar.createEl('span', {
                    text: '⚠️ Save failed',
                    cls: 'sheet-auto-save-indicator error'
                });
                errorSpan.title = autoSaveState.lastError;
            } else if (isDirty) {
                const modifiedCount = this.dataManager.getModifiedCells().length;
                toolbar.createEl('span', {
                    text: `${modifiedCount} unsaved change${modifiedCount !== 1 ? 's' : ''}`,
                    cls: 'sheet-modified-indicator'
                });
            } else if (autoSaveState.lastSaved) {
                const timeSince = this.dataManager.getTimeSinceLastSave();
                const timeText = this.formatTimeSince(timeSince);
                toolbar.createEl('span', {
                    text: `✓ Saved ${timeText}`,
                    cls: 'sheet-auto-save-indicator saved'
                });
            }
        }
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
        const sheetIndex = this.sheetData.indexOf(sheet);
        
        // Update grid dimensions for navigation
        this.gridDimensions = {
            rows: grid.length,
            cols: grid.length > 0 ? grid[0].length : 0
        };
        
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
                const cellData = grid[r][c];
                const cellValue = SheetParser.getCellValue(cellData);
                const actualRow = rowMap[r];
                const actualCol = colMap[c];
                
                const td = row.createEl('td', { 
                    cls: 'sheet-cell',
                    text: cellValue
                });

                // Add data attributes for navigation
                td.setAttribute('data-row', actualRow.toString());
                td.setAttribute('data-col', actualCol.toString());
                td.setAttribute('data-sheet', sheetIndex.toString());
                td.setAttribute('tabindex', '0'); // Make focusable

                // Add selection styling
                if (this.selectedCell && 
                    this.selectedCell.sheetIndex === sheetIndex && 
                    this.selectedCell.row === actualRow && 
                    this.selectedCell.col === actualCol) {
                    td.addClass('sheet-cell-selected');
                }

                // Add formula indicator
                if (cellData?.f) {
                    td.addClass('sheet-cell-formula');
                    td.title = `Formula: ${cellData.f}`;
                    
                    // Add formula icon
                    const formulaIcon = td.createEl('span', { 
                        cls: 'sheet-formula-icon',
                        text: 'ƒ'
                    });
                }

                // Apply header styling to first non-empty row/column
                if (r === 0 && rowMap[0] === 0) {
                    td.addClass('sheet-header-row');
                }
                if (c === 0 && colMap[0] === 0) {
                    td.addClass('sheet-header-col');
                }

                // Add cell state styling
                if (this.dataManager) {
                    const cellState = this.dataManager.getCellState(sheet.id, actualRow, actualCol);
                    if (cellState) {
                        if (cellState.isModified) {
                            td.addClass('sheet-cell-modified');
                        }
                        if (!cellState.isValid) {
                            td.addClass('sheet-cell-invalid');
                            td.title = cellState.validationError || 'Invalid value';
                        }
                    }
                }

                // Add interaction handlers
                this.setupCellInteractionHandlers(td, sheetIndex, actualRow, actualCol);
            }
        }
    }

    private toggleEditMode(): void {
        this.isEditing = !this.isEditing;
        this.finishCellEdit();
        this.render();
    }

    private startCellEdit(cellEl: HTMLElement, sheetIndex: number, row: number, col: number): void {
        if (this.currentEditor) {
            this.finishCellEdit();
        }

        const sheet = this.sheetData[sheetIndex];
        const cell = SheetParser.getCellAt(sheet, row, col);
        
        // For formula cells, show the formula; for regular cells, show the displayed value
        const currentValue = (cell?.v.f) ? cell.v.f : (cellEl.textContent || '');
        const isMultiline = this.shouldUseMultilineInput(currentValue);
        
        cellEl.innerHTML = '';

        let input: HTMLInputElement | HTMLTextAreaElement;
        
        if (isMultiline) {
            input = cellEl.createEl('textarea', {
                value: currentValue,
                cls: 'sheet-cell-input sheet-cell-textarea'
            });
        } else {
            // Determine input type based on content
            const inputType = this.detectInputType(currentValue);
            input = cellEl.createEl('input', {
                type: inputType,
                value: currentValue,
                cls: 'sheet-cell-input'
            });
        }

        // Add formula helper for formula cells
        if (currentValue.startsWith('=')) {
            input.addClass('sheet-formula-input');
        }

        this.currentEditor = input;
        this.editingCell = { sheetIndex, row, col };

        // Mobile-optimized input handling
        this.setupMobileInput(input);

        input.focus();
        if (input instanceof HTMLInputElement) {
            input.select();
        } else {
            // For textarea, select all text
            input.setSelectionRange(0, input.value.length);
        }

        input.addEventListener('blur', () => this.finishCellEdit());

        // Add formula help for new formulas
        input.addEventListener('input', () => {
            if (input.value.startsWith('=')) {
                input.addClass('sheet-formula-input');
                this.showFormulaHelp(input);
            } else {
                input.removeClass('sheet-formula-input');
                this.hideFormulaHelp();
            }
        });
    }

    private shouldUseMultilineInput(value: string): boolean {
        return value.includes('\n') || value.length > 100;
    }

    private detectInputType(value: string): string {
        // Don't change type for formulas
        if (value.startsWith('=')) {
            return 'text';
        }

        // Check for numbers
        if (/^\d*\.?\d+$/.test(value.trim())) {
            return 'number';
        }

        // Check for dates (basic patterns)
        if (/^\d{4}-\d{2}-\d{2}$/.test(value.trim()) || 
            /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value.trim())) {
            return 'date';
        }

        // Check for emails
        if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim())) {
            return 'email';
        }

        // Check for URLs
        if (/^https?:\/\//.test(value.trim())) {
            return 'url';
        }

        return 'text';
    }

    private isMultilineInput(): boolean {
        return this.currentEditor instanceof HTMLTextAreaElement;
    }

    private showFormulaHelp(input: HTMLInputElement): void {
        // Remove existing help
        this.hideFormulaHelp();
        
        const help = document.createElement('div');
        help.className = 'sheet-formula-help';
        help.innerHTML = `
            <div class="formula-functions">
                <strong>Functions:</strong>
                SUM(A1:A5), COUNT(A1:A5), AVERAGE(A1:A5), MIN(A1:A5), MAX(A1:A5),
                IF(condition, true_value, false_value), CONCATENATE(text1, text2),
                LEN(text), UPPER(text), LOWER(text), ROUND(number, digits)
            </div>
            <div class="formula-examples">
                <strong>Examples:</strong>
                =A1+B1, =SUM(A1:A5), =IF(A1>10, "High", "Low")
            </div>
        `;
        
        const rect = input.getBoundingClientRect();
        help.style.position = 'fixed';
        help.style.top = (rect.bottom + 5) + 'px';
        help.style.left = rect.left + 'px';
        help.style.zIndex = '1001';
        
        document.body.appendChild(help);
    }

    private hideFormulaHelp(): void {
        const existing = document.querySelector('.sheet-formula-help');
        if (existing) {
            existing.remove();
        }
    }

    private setupMobileInput(input: HTMLInputElement | HTMLTextAreaElement): void {
        // Prevent zoom on iOS when focusing input
        input.style.fontSize = '16px';
        
        // Handle virtual keyboard
        input.addEventListener('focus', () => {
            this.handleVirtualKeyboard(true);
        });
        
        input.addEventListener('blur', () => {
            this.handleVirtualKeyboard(false);
        });

        // Auto-resize input based on content
        input.addEventListener('input', () => {
            if (input instanceof HTMLTextAreaElement) {
                this.autoResizeTextarea(input);
            } else {
                this.autoResizeInput(input);
            }
        });

        // Initial resize
        setTimeout(() => {
            if (input instanceof HTMLTextAreaElement) {
                this.autoResizeTextarea(input);
            } else {
                this.autoResizeInput(input);
            }
        }, 0);
    }

    private handleVirtualKeyboard(isShowing: boolean): void {
        const viewport = document.querySelector('meta[name=viewport]');
        if (!viewport) return;

        if (isShowing) {
            // Adjust viewport when keyboard shows
            viewport.setAttribute('content', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
            
            // Scroll to editing cell if needed
            if (this.currentEditor) {
                setTimeout(() => {
                    this.currentEditor?.scrollIntoView({ 
                        behavior: 'smooth', 
                        block: 'center' 
                    });
                }, 300);
            }
        } else {
            // Restore normal viewport
            viewport.setAttribute('content', 'width=device-width, initial-scale=1.0');
        }
    }

    private autoResizeInput(input: HTMLInputElement): void {
        // Create temporary span to measure text width
        const span = document.createElement('span');
        span.style.visibility = 'hidden';
        span.style.position = 'absolute';
        span.style.fontSize = window.getComputedStyle(input).fontSize;
        span.style.fontFamily = window.getComputedStyle(input).fontFamily;
        span.textContent = input.value || input.placeholder || '';
        
        document.body.appendChild(span);
        const textWidth = span.offsetWidth;
        document.body.removeChild(span);
        
        // Set input width with padding
        input.style.width = Math.max(textWidth + 20, 60) + 'px';
    }

    private autoResizeTextarea(textarea: HTMLTextAreaElement): void {
        // Reset height to auto to get the correct scrollHeight
        textarea.style.height = 'auto';
        
        // Set height based on scrollHeight
        const newHeight = Math.max(textarea.scrollHeight, 60);
        textarea.style.height = newHeight + 'px';
        
        // Also adjust width if needed
        const maxWidth = Math.min(window.innerWidth * 0.8, 400);
        if (textarea.scrollWidth > textarea.clientWidth) {
            textarea.style.width = Math.min(textarea.scrollWidth + 20, maxWidth) + 'px';
        }
    }

    private setupCellTouchHandlers(cellEl: HTMLElement, sheetIndex: number, row: number, col: number): void {
        let tapStartTime = 0;
        
        cellEl.addEventListener('touchstart', (e) => {
            tapStartTime = Date.now();
            this.touchStartPos = {
                x: e.touches[0].clientX,
                y: e.touches[0].clientY
            };
            this.isLongPress = false;
            
            // Start long press timer
            this.longPressTimer = window.setTimeout(() => {
                this.isLongPress = true;
                this.handleLongPress(cellEl, sheetIndex, row, col);
                this.triggerHapticFeedback();
            }, this.longPressDelay);
        });

        cellEl.addEventListener('touchmove', (e) => {
            if (!this.touchStartPos) return;
            
            const touch = e.touches[0];
            const deltaX = Math.abs(touch.clientX - this.touchStartPos.x);
            const deltaY = Math.abs(touch.clientY - this.touchStartPos.y);
            
            // Cancel long press if moved too much
            if (deltaX > 10 || deltaY > 10) {
                this.clearLongPressTimer();
            }
        });

        cellEl.addEventListener('touchend', (e) => {
            this.clearLongPressTimer();
            
            if (!this.isLongPress && this.touchStartPos) {
                const tapDuration = Date.now() - tapStartTime;
                
                if (tapDuration < 200) { // Quick tap
                    this.startCellEdit(cellEl, sheetIndex, row, col);
                    this.triggerHapticFeedback('light');
                }
            }
            
            this.touchStartPos = null;
        });

        // Also handle mouse events for desktop compatibility
        cellEl.addEventListener('click', (e) => {
            if (!('ontouchstart' in window)) {
                this.startCellEdit(cellEl, sheetIndex, row, col);
            }
        });
    }

    private setupViewModeTouchHandlers(cellEl: HTMLElement, sheetIndex: number, row: number, col: number): void {
        cellEl.addEventListener('touchstart', (e) => {
            this.touchStartPos = {
                x: e.touches[0].clientX,
                y: e.touches[0].clientY
            };
            
            this.longPressTimer = window.setTimeout(() => {
                this.handleLongPress(cellEl, sheetIndex, row, col);
                this.triggerHapticFeedback();
            }, this.longPressDelay);
        });

        cellEl.addEventListener('touchend', () => {
            this.clearLongPressTimer();
            this.touchStartPos = null;
        });
    }

    private handleLongPress(cellEl: HTMLElement, sheetIndex: number, row: number, col: number): void {
        const sheet = this.sheetData[sheetIndex];
        const cellValue = SheetParser.getCellAt(sheet, row, col);
        const currentValue = cellValue ? SheetParser.getCellValue(cellValue.v) : '';
        
        this.showContextMenu(cellEl, {
            sheetIndex,
            row,
            col,
            value: currentValue
        });
    }

    private showContextMenu(cellEl: HTMLElement, cellInfo: any): void {
        // Remove existing context menu
        const existingMenu = document.querySelector('.sheet-context-menu');
        if (existingMenu) {
            existingMenu.remove();
        }

        const menu = document.createElement('div');
        menu.className = 'sheet-context-menu';
        
        const rect = cellEl.getBoundingClientRect();
        menu.style.position = 'fixed';
        menu.style.left = rect.left + 'px';
        menu.style.top = (rect.bottom + 5) + 'px';
        menu.style.zIndex = '1000';

        const autoSaveEnabled = this.dataManager?.getAutoSaveConfig().enabled ?? true;
        const isLongText = cellInfo.value.length > 50 || cellInfo.value.includes('\n');
        
        const menuItems = [
            { label: 'Edit', action: () => this.toggleEditModeAndEdit(cellEl, cellInfo), shortcut: 'F2' },
            { label: isLongText ? 'Edit as Text' : 'Edit as Multi-line', action: () => this.editCellWithType(cellEl, cellInfo, isLongText ? 'text' : 'multiline') },
            { label: 'Copy', action: () => this.copyCellValue(cellInfo.value), shortcut: 'Ctrl+C' },
            { label: 'Paste', action: () => this.pasteCellValue(cellInfo), shortcut: 'Ctrl+V' },
            { label: 'Clear Cell', action: () => this.clearCell(cellInfo), shortcut: 'Del' },
            { label: 'Delete Row', action: () => this.deleteRow(cellInfo), shortcut: 'Ctrl+Shift+-' },
            { label: 'Delete Column', action: () => this.deleteColumn(cellInfo), shortcut: 'Ctrl+Alt+-' },
            { label: autoSaveEnabled ? 'Disable Auto-save' : 'Enable Auto-save', action: () => this.toggleAutoSave() },
            { label: 'Cancel', action: () => menu.remove(), shortcut: 'Esc' }
        ];

        menuItems.forEach(item => {
            const button = menu.createEl('button', {
                cls: 'sheet-context-menu-item'
            });
            
            const labelSpan = button.createEl('span', { text: item.label });
            if (item.shortcut) {
                const shortcutSpan = button.createEl('span', { 
                    text: item.shortcut,
                    cls: 'sheet-context-menu-shortcut'
                });
            }
            
            button.addEventListener('click', () => {
                item.action();
                menu.remove();
            });
        });

        document.body.appendChild(menu);

        // Auto-remove after 5 seconds
        setTimeout(() => menu.remove(), 5000);
    }

    private toggleEditModeAndEdit(cellEl: HTMLElement, cellInfo: any): void {
        if (!this.isEditing) {
            this.isEditing = true;
            this.render();
        }
        // Wait for render to complete, then start editing
        setTimeout(() => {
            const newCellEl = this.findCellElement(cellInfo.row, cellInfo.col);
            if (newCellEl) {
                this.startCellEdit(newCellEl, cellInfo.sheetIndex, cellInfo.row, cellInfo.col);
            }
        }, 100);
    }

    private editCellWithType(cellEl: HTMLElement, cellInfo: any, editType: 'text' | 'multiline'): void {
        if (!this.isEditing) {
            this.isEditing = true;
            this.render();
        }
        
        // Wait for render to complete, then start editing with specific type
        setTimeout(() => {
            const newCellEl = this.findCellElement(cellInfo.row, cellInfo.col);
            if (newCellEl) {
                this.startCellEditWithType(newCellEl, cellInfo.sheetIndex, cellInfo.row, cellInfo.col, editType);
            }
        }, 100);
    }

    private startCellEditWithType(cellEl: HTMLElement, sheetIndex: number, row: number, col: number, editType: 'text' | 'multiline'): void {
        if (this.currentEditor) {
            this.finishCellEdit();
        }

        const sheet = this.sheetData[sheetIndex];
        const cell = SheetParser.getCellAt(sheet, row, col);
        const currentValue = (cell?.v.f) ? cell.v.f : (cellEl.textContent || '');
        
        cellEl.innerHTML = '';

        let input: HTMLInputElement | HTMLTextAreaElement;
        
        if (editType === 'multiline') {
            input = cellEl.createEl('textarea', {
                value: currentValue,
                cls: 'sheet-cell-input sheet-cell-textarea'
            });
        } else {
            input = cellEl.createEl('input', {
                type: 'text',
                value: currentValue,
                cls: 'sheet-cell-input'
            });
        }

        if (currentValue.startsWith('=')) {
            input.addClass('sheet-formula-input');
        }

        this.currentEditor = input;
        this.editingCell = { sheetIndex, row, col };

        this.setupMobileInput(input);
        input.focus();
        
        if (input instanceof HTMLInputElement) {
            input.select();
        } else {
            input.setSelectionRange(0, input.value.length);
        }

        input.addEventListener('blur', () => this.finishCellEdit());

        input.addEventListener('input', () => {
            if (input.value.startsWith('=')) {
                input.addClass('sheet-formula-input');
                this.showFormulaHelp(input);
            } else {
                input.removeClass('sheet-formula-input');
                this.hideFormulaHelp();
            }
        });
    }

    private findCellElement(row: number, col: number): HTMLElement | null {
        const tables = this.contentEl.querySelectorAll('.sheet-table');
        for (const table of tables) {
            const cells = table.querySelectorAll('.sheet-cell');
            // This is a simplified approach - in a real implementation,
            // you'd want to track cell positions more precisely
            const cellIndex = row * 10 + col; // Approximate
            if (cells[cellIndex]) {
                return cells[cellIndex] as HTMLElement;
            }
        }
        return null;
    }

    private copyCellValue(value: string): void {
        if (navigator.clipboard) {
            navigator.clipboard.writeText(value);
            this.showToast('Cell value copied to clipboard');
        }
    }

    private copySelectedCell(): void {
        if (!this.selectedCell) return;
        
        const sheet = this.sheetData[this.selectedCell.sheetIndex];
        const cellValue = SheetParser.getCellAt(sheet, this.selectedCell.row, this.selectedCell.col);
        const value = cellValue ? SheetParser.getCellValue(cellValue.v) : '';
        
        this.copyCellValue(value);
    }

    private async pasteToSelectedCell(): Promise<void> {
        if (!this.selectedCell || !navigator.clipboard) return;
        
        try {
            const text = await navigator.clipboard.readText();
            await this.pasteCellValue({ 
                sheetIndex: this.selectedCell.sheetIndex,
                row: this.selectedCell.row,
                col: this.selectedCell.col 
            }, text);
        } catch (error) {
            this.showToast('Failed to paste: clipboard access denied');
        }
    }

    private async pasteCellValue(cellInfo: any, value?: string): Promise<void> {
        if (!this.dataManager) return;
        
        try {
            const textToPaste = value || await navigator.clipboard.readText();
            const sheet = this.sheetData[cellInfo.sheetIndex];
            
            this.dataManager.updateCell(sheet.id, cellInfo.row, cellInfo.col, textToPaste);
            this.updateDisplayText();
            this.render();
            this.showToast('Value pasted');
        } catch (error) {
            this.showToast('Failed to paste');
        }
    }

    private cutSelectedCell(): void {
        if (!this.selectedCell) return;
        
        this.copySelectedCell();
        this.clearSelectedCell();
        this.showToast('Cell cut to clipboard');
    }

    private clearCell(cellInfo: any): void {
        if (this.dataManager) {
            const sheet = this.sheetData[cellInfo.sheetIndex];
            this.dataManager.updateCell(sheet.id, cellInfo.row, cellInfo.col, '');
            this.updateDisplayText();
            this.render();
            this.showToast('Cell cleared');
        }
    }

    private deleteRow(cellInfo: any): void {
        if (!this.dataManager) return;

        if (confirm(`Delete row ${cellInfo.row + 1}? This action cannot be undone.`)) {
            try {
                const sheet = this.sheetData[cellInfo.sheetIndex];
                
                // Clear all cells in the row
                const gridData = SheetParser.createGrid(sheet.celldata);
                const { colMap } = gridData;
                
                colMap.forEach(col => {
                    this.dataManager?.updateCell(sheet.id, cellInfo.row, col, '');
                });
                
                this.updateDisplayText();
                this.render();
                this.showToast(`Row ${cellInfo.row + 1} deleted`);
            } catch (error) {
                this.showToast('Failed to delete row');
            }
        }
    }

    private deleteColumn(cellInfo: any): void {
        if (!this.dataManager) return;

        if (confirm(`Delete column ${this.getColumnName(cellInfo.col)}? This action cannot be undone.`)) {
            try {
                const sheet = this.sheetData[cellInfo.sheetIndex];
                
                // Clear all cells in the column
                const gridData = SheetParser.createGrid(sheet.celldata);
                const { rowMap } = gridData;
                
                rowMap.forEach(row => {
                    this.dataManager?.updateCell(sheet.id, row, cellInfo.col, '');
                });
                
                this.updateDisplayText();
                this.render();
                this.showToast(`Column ${this.getColumnName(cellInfo.col)} deleted`);
            } catch (error) {
                this.showToast('Failed to delete column');
            }
        }
    }

    private getColumnName(colIndex: number): string {
        let result = '';
        let index = colIndex;
        
        while (index >= 0) {
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26) - 1;
        }
        
        return result;
    }

    private showToast(message: string): void {
        const toast = document.createElement('div');
        toast.className = 'sheet-toast';
        toast.textContent = message;
        document.body.appendChild(toast);
        
        setTimeout(() => toast.remove(), 2000);
    }

    private clearLongPressTimer(): void {
        if (this.longPressTimer) {
            clearTimeout(this.longPressTimer);
            this.longPressTimer = null;
        }
    }

    private triggerHapticFeedback(type: 'light' | 'medium' | 'heavy' = 'medium'): void {
        if ('vibrate' in navigator) {
            const patterns = {
                light: [10],
                medium: [20],
                heavy: [30]
            };
            navigator.vibrate(patterns[type]);
        }
    }

    private setupOrientationHandler(): void {
        const handleOrientationChange = () => {
            setTimeout(() => {
                this.render();
                // Adjust layout for orientation
                this.adjustForOrientation();
            }, 200);
        };

        if (screen.orientation) {
            screen.orientation.addEventListener('change', handleOrientationChange);
        } else {
            window.addEventListener('orientationchange', handleOrientationChange);
        }
    }

    private adjustForOrientation(): void {
        const isLandscape = window.innerWidth > window.innerHeight;
        
        if (isLandscape) {
            this.contentEl.addClass('sheet-landscape');
            this.contentEl.removeClass('sheet-portrait');
        } else {
            this.contentEl.addClass('sheet-portrait');
            this.contentEl.removeClass('sheet-landscape');
        }

        // Adjust table container height based on orientation
        const tableContainers = this.contentEl.querySelectorAll('.sheet-table-container');
        tableContainers.forEach((container: HTMLElement) => {
            if (isLandscape) {
                container.style.maxHeight = '80vh';
            } else {
                container.style.maxHeight = '60vh';
            }
        });
    }

    private setupKeyboardHandlers(): void {
        // Global keyboard handler for the view
        this.contentEl.addEventListener('keydown', (e) => {
            if (this.currentEditor) {
                // Handle editing mode keys
                this.handleEditingKeys(e);
            } else if (this.selectedCell) {
                // Handle navigation mode keys
                this.handleNavigationKeys(e);
            }
        });
    }

    private handleNavigationKeys(e: KeyboardEvent): void {
        if (!this.selectedCell) return;

        // Handle keyboard shortcuts with modifiers
        if (e.ctrlKey || e.metaKey) {
            switch (e.key.toLowerCase()) {
                case 'c':
                    e.preventDefault();
                    this.copySelectedCell();
                    break;
                case 'v':
                    e.preventDefault();
                    this.pasteToSelectedCell();
                    break;
                case 'x':
                    e.preventDefault();
                    this.cutSelectedCell();
                    break;
                case 's':
                    e.preventDefault();
                    this.saveFile();
                    break;
                case 'z':
                    e.preventDefault();
                    if (e.shiftKey) {
                        this.redo();
                    } else {
                        this.undo();
                    }
                    break;
                case 'e':
                    e.preventDefault();
                    this.toggleEditMode();
                    break;
                default:
                    return; // Don't prevent default for unhandled shortcuts
            }
            return;
        }

        switch (e.key) {
            case 'ArrowUp':
                e.preventDefault();
                this.moveSelection(-1, 0);
                break;
            case 'ArrowDown':
                e.preventDefault();
                this.moveSelection(1, 0);
                break;
            case 'ArrowLeft':
                e.preventDefault();
                this.moveSelection(0, -1);
                break;
            case 'ArrowRight':
                e.preventDefault();
                this.moveSelection(0, 1);
                break;
            case 'Tab':
                e.preventDefault();
                this.moveSelection(0, e.shiftKey ? -1 : 1);
                break;
            case 'Enter':
                e.preventDefault();
                if (this.isEditing) {
                    this.startCellEditFromSelection();
                } else {
                    this.moveSelection(1, 0);
                }
                break;
            case 'F2':
            case ' ':
                e.preventDefault();
                this.startCellEditFromSelection();
                break;
            case 'Delete':
            case 'Backspace':
                e.preventDefault();
                this.clearSelectedCell();
                break;
            case 'Escape':
                e.preventDefault();
                this.clearSelection();
                break;
        }
    }

    private handleEditingKeys(e: KeyboardEvent): void {
        if (!this.currentEditor) return;

        switch (e.key) {
            case 'Enter':
                if (!e.shiftKey && !this.isMultilineInput()) {
                    e.preventDefault();
                    this.finishCellEdit();
                    this.moveSelection(1, 0);
                }
                break;
            case 'Tab':
                e.preventDefault();
                this.finishCellEdit();
                this.moveSelection(0, e.shiftKey ? -1 : 1);
                break;
            case 'Escape':
                e.preventDefault();
                this.cancelCellEdit();
                break;
        }
    }

    private moveSelection(deltaRow: number, deltaCol: number): void {
        if (!this.selectedCell) return;

        const newRow = Math.max(0, Math.min(this.gridDimensions.rows - 1, 
            this.findVisualRow(this.selectedCell.row) + deltaRow));
        const newCol = Math.max(0, Math.min(this.gridDimensions.cols - 1, 
            this.findVisualCol(this.selectedCell.col) + deltaCol));

        // Convert back to actual coordinates
        const actualRow = this.getActualRow(newRow);
        const actualCol = this.getActualCol(newCol);

        this.selectCell(this.selectedCell.sheetIndex, actualRow, actualCol);
    }

    private findVisualRow(actualRow: number): number {
        // Find the visual position of this actual row
        const gridData = SheetParser.createGrid(this.sheetData[this.selectedCell?.sheetIndex || 0].celldata);
        return gridData.rowMap.indexOf(actualRow);
    }

    private findVisualCol(actualCol: number): number {
        // Find the visual position of this actual column
        const gridData = SheetParser.createGrid(this.sheetData[this.selectedCell?.sheetIndex || 0].celldata);
        return gridData.colMap.indexOf(actualCol);
    }

    private getActualRow(visualRow: number): number {
        const gridData = SheetParser.createGrid(this.sheetData[this.selectedCell?.sheetIndex || 0].celldata);
        return gridData.rowMap[visualRow] || 0;
    }

    private getActualCol(visualCol: number): number {
        const gridData = SheetParser.createGrid(this.sheetData[this.selectedCell?.sheetIndex || 0].celldata);
        return gridData.colMap[visualCol] || 0;
    }

    private selectCell(sheetIndex: number, row: number, col: number): void {
        this.selectedCell = { sheetIndex, row, col };
        this.render();
        
        // Focus the selected cell
        const cellEl = this.findCellElement(row, col);
        if (cellEl) {
            cellEl.focus();
            cellEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    }

    private startCellEditFromSelection(): void {
        if (!this.selectedCell || !this.isEditing) return;

        const cellEl = this.findCellElement(this.selectedCell.row, this.selectedCell.col);
        if (cellEl) {
            this.startCellEdit(cellEl, this.selectedCell.sheetIndex, this.selectedCell.row, this.selectedCell.col);
        }
    }

    private clearSelectedCell(): void {
        if (!this.selectedCell || !this.dataManager) return;

        const sheet = this.sheetData[this.selectedCell.sheetIndex];
        this.dataManager.updateCell(sheet.id, this.selectedCell.row, this.selectedCell.col, '');
        this.updateDisplayText();
        this.render();
        this.showToast('Cell cleared');
    }

    private clearSelection(): void {
        this.selectedCell = null;
        this.render();
    }

    private setupCellInteractionHandlers(cellEl: HTMLElement, sheetIndex: number, row: number, col: number): void {
        // Click handler for selection
        cellEl.addEventListener('click', (e) => {
            e.preventDefault();
            this.selectCell(sheetIndex, row, col);
        });

        // Double-click for editing
        cellEl.addEventListener('dblclick', (e) => {
            e.preventDefault();
            if (this.isEditing) {
                this.startCellEdit(cellEl, sheetIndex, row, col);
            }
        });

        // Focus handler
        cellEl.addEventListener('focus', () => {
            this.selectCell(sheetIndex, row, col);
        });

        // Touch handlers for mobile
        if (this.isEditing) {
            this.setupCellTouchHandlers(cellEl, sheetIndex, row, col);
        } else {
            this.setupViewModeTouchHandlers(cellEl, sheetIndex, row, col);
        }
    }

    private setupSwipeGestures(): void {
        let swipeStartX = 0;
        let swipeStartY = 0;
        let swipeStartTime = 0;

        this.contentEl.addEventListener('touchstart', (e) => {
            if (e.touches.length === 1) {
                swipeStartX = e.touches[0].clientX;
                swipeStartY = e.touches[0].clientY;
                swipeStartTime = Date.now();
            }
        });

        this.contentEl.addEventListener('touchend', (e) => {
            if (e.changedTouches.length === 1) {
                const swipeEndX = e.changedTouches[0].clientX;
                const swipeEndY = e.changedTouches[0].clientY;
                const swipeTime = Date.now() - swipeStartTime;

                const deltaX = swipeEndX - swipeStartX;
                const deltaY = swipeEndY - swipeStartY;

                // Only process quick swipes
                if (swipeTime < 300) {
                    this.handleSwipeGesture(deltaX, deltaY);
                }
            }
        });
    }

    private handleSwipeGesture(deltaX: number, deltaY: number): void {
        const absX = Math.abs(deltaX);
        const absY = Math.abs(deltaY);

        // Must be primarily horizontal or vertical
        if (absX < this.swipeThreshold && absY < this.swipeThreshold) return;

        if (absX > absY) {
            // Horizontal swipe
            if (deltaX > this.swipeThreshold) {
                // Swipe right - show previous sheet or toggle mode
                this.handleSwipeRight();
            } else if (deltaX < -this.swipeThreshold) {
                // Swipe left - show next sheet or undo
                this.handleSwipeLeft();
            }
        } else {
            // Vertical swipe
            if (deltaY > this.swipeThreshold) {
                // Swipe down - minimize or exit edit mode
                this.handleSwipeDown();
            } else if (deltaY < -this.swipeThreshold) {
                // Swipe up - toggle edit mode or save
                this.handleSwipeUp();
            }
        }
    }

    private handleSwipeRight(): void {
        // Toggle to view mode if editing
        if (this.isEditing) {
            this.toggleEditMode();
            this.showToast('Switched to view mode');
            this.triggerHapticFeedback('light');
        }
    }

    private handleSwipeLeft(): void {
        // Undo last action if available
        if (this.dataManager?.canUndo()) {
            this.undo();
            this.showToast('Undid last action');
            this.triggerHapticFeedback('medium');
        }
    }

    private handleSwipeDown(): void {
        // Exit edit mode if editing
        if (this.isEditing) {
            this.finishCellEdit();
            this.toggleEditMode();
            this.showToast('Exited edit mode');
            this.triggerHapticFeedback('light');
        }
    }

    private handleSwipeUp(): void {
        // Toggle edit mode or save if dirty
        if (!this.isEditing) {
            this.toggleEditMode();
            this.showToast('Entered edit mode');
            this.triggerHapticFeedback('light');
        } else if (this.dataManager?.isDirty()) {
            this.saveFile();
            this.showToast('File saved');
            this.triggerHapticFeedback('medium');
        }
    }

    private finishCellEdit(): void {
        if (!this.currentEditor || !this.editingCell || !this.dataManager) return;

        const newValue = this.currentEditor.value;
        const { sheetIndex, row, col } = this.editingCell;
        const sheet = this.sheetData[sheetIndex];

        this.hideFormulaHelp();

        try {
            const success = this.dataManager.updateCell(sheet.id, row, col, newValue);
            if (success) {
                this.updateDisplayText();
            } else {
                this.showToast('Invalid formula or value');
                this.showFallbackForUnsupportedOperation('cell update', newValue);
            }
        } catch (error) {
            this.showToast('Error: ' + (error instanceof Error ? error.message : 'Unknown error'));
            this.showFallbackForUnsupportedOperation('cell update', newValue, error);
        }

        this.currentEditor = null;
        this.editingCell = null;
        this.render();
    }

    private cancelCellEdit(): void {
        this.hideFormulaHelp();
        this.currentEditor = null;
        this.editingCell = null;
        this.render();
    }

    private updateDisplayText(): void {
        const tabEl = (this.leaf as any).tabHeaderInnerTitleEl;
        if (tabEl) {
            tabEl.textContent = this.getDisplayText();
        }
    }

    private undo(): void {
        if (this.dataManager?.undo()) {
            this.updateDisplayText();
            this.render();
        }
    }

    private redo(): void {
        if (this.dataManager?.redo()) {
            this.updateDisplayText();
            this.render();
        }
    }

    addValidationRule(sheetId: string, row: number, col: number, rule: ValidationRule): void {
        this.dataManager?.addValidationRule(sheetId, row, col, rule);
    }

    private formatTimeSince(ms: number | null): string {
        if (!ms) return 'never';
        
        const seconds = Math.floor(ms / 1000);
        if (seconds < 60) return 'just now';
        
        const minutes = Math.floor(seconds / 60);
        if (minutes < 60) return `${minutes}m ago`;
        
        const hours = Math.floor(minutes / 60);
        if (hours < 24) return `${hours}h ago`;
        
        return 'yesterday';
    }

    private async performSave(): Promise<void> {
        if (!this.file || !this.dataManager) return;

        // Show loading state
        this.showLoadingState('Saving...');
        
        try {
            const serialized = SheetParser.serialize(this.dataManager.getSheetData());
            await this.app.vault.modify(this.file, serialized);
            this.dataManager.markClean();
            this.updateDisplayText();
            
            this.hideLoadingState();
            // Re-render to update status indicators
            this.render();
        } catch (error) {
            this.hideLoadingState();
            throw error;
        }
    }

    private async saveFile(): Promise<void> {
        if (!this.file || !this.dataManager?.isDirty()) return;

        const unsavedChanges = this.dataManager.getModifiedCells().length;
        const confirmSave = await this.showSaveConfirmDialog(unsavedChanges);
        
        if (!confirmSave) return;

        this.showLoadingState('Saving file...');
        
        try {
            await this.performSave();
            this.showToast('File saved manually');
        } catch (error) {
            console.error('Failed to save file:', error);
            this.showToast('Save failed: ' + (error instanceof Error ? error.message : 'Unknown error'));
        } finally {
            this.hideLoadingState();
        }
    }

    // Add method to toggle auto-save
    toggleAutoSave(): void {
        if (this.dataManager) {
            const currentState = this.dataManager.getAutoSaveConfig().enabled;
            this.dataManager.enableAutoSave(!currentState);
            this.showToast(`Auto-save ${!currentState ? 'enabled' : 'disabled'}`);
            this.render();
        }
    }

    // Force auto-save (useful for testing or manual triggers)
    async forceAutoSave(): Promise<void> {
        if (this.dataManager) {
            this.showLoadingState('Auto-saving...');
            try {
                await this.dataManager.forceAutoSave();
                this.showToast('Auto-save completed');
            } catch (error) {
                this.showToast('Auto-save failed');
            } finally {
                this.hideLoadingState();
            }
        }
    }

    private showLoadingState(message: string): void {
        // Remove existing loading indicator
        const existing = document.querySelector('.sheet-loading-overlay');
        if (existing) existing.remove();

        const overlay = document.createElement('div');
        overlay.className = 'sheet-loading-overlay';
        overlay.innerHTML = `
            <div class="sheet-loading-content">
                <div class="sheet-loading-spinner"></div>
                <span class="sheet-loading-text">${message}</span>
            </div>
        `;
        
        document.body.appendChild(overlay);
    }

    private hideLoadingState(): void {
        const overlay = document.querySelector('.sheet-loading-overlay');
        if (overlay) {
            overlay.remove();
        }
    }

    private async showSaveConfirmDialog(unsavedChanges: number): Promise<boolean> {
        return new Promise((resolve) => {
            // Remove existing dialog
            const existing = document.querySelector('.sheet-save-dialog');
            if (existing) existing.remove();

            const dialog = document.createElement('div');
            dialog.className = 'sheet-save-dialog';
            
            dialog.innerHTML = `
                <div class="sheet-save-dialog-content">
                    <h3>Save Changes</h3>
                    <p>You have ${unsavedChanges} unsaved change${unsavedChanges !== 1 ? 's' : ''}.</p>
                    <p>Do you want to save these changes?</p>
                    <div class="sheet-save-dialog-buttons">
                        <button class="sheet-dialog-btn sheet-dialog-btn-primary" data-action="save">Save</button>
                        <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-action="cancel">Cancel</button>
                    </div>
                </div>
            `;

            const handleClick = (e: Event) => {
                const target = e.target as HTMLElement;
                const action = target.getAttribute('data-action');
                
                if (action === 'save') {
                    resolve(true);
                } else if (action === 'cancel') {
                    resolve(false);
                }
                
                dialog.remove();
            };

            dialog.addEventListener('click', handleClick);
            
            // Handle escape key
            const handleKeydown = (e: KeyboardEvent) => {
                if (e.key === 'Escape') {
                    resolve(false);
                    dialog.remove();
                    document.removeEventListener('keydown', handleKeydown);
                }
            };
            
            document.addEventListener('keydown', handleKeydown);
            document.body.appendChild(dialog);
            
            // Focus the save button
            const saveBtn = dialog.querySelector('[data-action="save"]') as HTMLElement;
            saveBtn?.focus();
        });
    }

    private showFallbackForUnsupportedOperation(operation: string, value?: string, error?: any): void {
        console.warn(`Unsupported operation: ${operation}`, { value, error });
        
        // Create a simple fallback dialog
        const fallbackDialog = document.createElement('div');
        fallbackDialog.className = 'sheet-fallback-dialog';
        
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        
        fallbackDialog.innerHTML = `
            <div class="sheet-fallback-dialog-content">
                <h3>⚠️ Operation Not Supported</h3>
                <p>The operation "${operation}" could not be completed.</p>
                ${value ? `<p><strong>Value:</strong> ${value}</p>` : ''}
                ${error ? `<p><strong>Error:</strong> ${errorMessage}</p>` : ''}
                <div class="sheet-fallback-suggestions">
                    <h4>Try these alternatives:</h4>
                    <ul>
                        <li>Simplify the formula or value</li>
                        <li>Check for syntax errors</li>
                        <li>Use a different format</li>
                        <li>Save as plain text</li>
                    </ul>
                </div>
                <div class="sheet-fallback-buttons">
                    <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-action="copy-error">Copy Error</button>
                    <button class="sheet-dialog-btn sheet-dialog-btn-primary" data-action="close">Close</button>
                </div>
            </div>
        `;

        const handleClick = (e: Event) => {
            const target = e.target as HTMLElement;
            const action = target.getAttribute('data-action');
            
            if (action === 'copy-error') {
                const errorDetails = `Operation: ${operation}\nValue: ${value || 'N/A'}\nError: ${errorMessage}`;
                if (navigator.clipboard) {
                    navigator.clipboard.writeText(errorDetails);
                    this.showToast('Error details copied to clipboard');
                }
            } else if (action === 'close') {
                fallbackDialog.remove();
            }
        };

        fallbackDialog.addEventListener('click', handleClick);
        
        // Auto-remove after 10 seconds
        setTimeout(() => {
            if (document.body.contains(fallbackDialog)) {
                fallbackDialog.remove();
            }
        }, 10000);
        
        document.body.appendChild(fallbackDialog);
    }

    // Conflict detection and monitoring
    private initializeConflictDetection(): void {
        if (!this.file || !this.dataManager) return;

        // Initialize file tracking
        const fileStats = this.app.vault.adapter.stat(this.file.path);
        fileStats.then(stats => {
            if (this.dataManager) {
                this.dataManager.stateManager.lastFileModified = new Date(stats.mtime);
                this.dataManager.stateManager.lastFileChecksum = this.dataManager.generateFileChecksum(this.lastFileContent);
            }
        });

        // Start periodic conflict checking
        this.startConflictMonitoring();
    }

    private startConflictMonitoring(): void {
        if (this.fileWatcher) {
            clearInterval(this.fileWatcher);
        }

        this.fileWatcher = window.setInterval(() => {
            this.checkForConflicts();
        }, this.conflictCheckInterval);
    }

    private stopConflictDetection(): void {
        if (this.fileWatcher) {
            clearInterval(this.fileWatcher);
            this.fileWatcher = null;
        }
    }

    private async checkForConflicts(): Promise<void> {
        if (!this.file || !this.dataManager) return;

        try {
            // Read current file content
            const currentContent = await this.app.vault.read(this.file);
            const stats = await this.app.vault.adapter.stat(this.file.path);
            const fileModifiedTime = new Date(stats.mtime);

            // Check for conflicts
            const conflict = this.dataManager.detectConflicts(currentContent, fileModifiedTime);
            
            if (conflict) {
                console.warn('Conflict detected:', conflict);
                this.handleConflictDetected(conflict);
            }

        } catch (error) {
            console.error('Error checking for conflicts:', error);
        }
    }

    private handleConflictDetected(conflict: ConflictInfo): void {
        // Update the toolbar to show conflict indicator
        this.render();
        
        // Show notification
        this.showToast(`⚠️ Conflicts detected in ${conflict.conflictedCells.length} cell${conflict.conflictedCells.length !== 1 ? 's' : ''}`);
        
        // Auto-show conflict resolution dialog if user is actively editing
        if (this.isEditing) {
            setTimeout(() => {
                this.showConflictResolutionDialog();
            }, 2000);
        }
    }

    private async showConflictResolutionDialog(): Promise<void> {
        if (!this.dataManager) return;

        const conflict = this.dataManager.getActiveConflict();
        if (!conflict) return;

        // Remove existing dialog
        const existing = document.querySelector('.sheet-conflict-dialog');
        if (existing) existing.remove();

        const dialog = document.createElement('div');
        dialog.className = 'sheet-conflict-dialog';
        
        dialog.innerHTML = `
            <div class="sheet-conflict-dialog-content">
                <h3>🔄 Resolve Conflicts</h3>
                <p>${conflict.conflictedCells.length} cell${conflict.conflictedCells.length !== 1 ? 's have' : ' has'} conflicting changes.</p>
                <div class="sheet-conflict-summary">
                    <p><strong>Conflict Type:</strong> ${this.getConflictTypeDescription(conflict.type)}</p>
                    <p><strong>Detected:</strong> ${conflict.timestamp.toLocaleString()}</p>
                </div>
                <div class="sheet-conflict-cells">
                    ${this.renderConflictedCells(conflict)}
                </div>
                <div class="sheet-conflict-strategies">
                    <h4>Resolution Strategy:</h4>
                    <div class="sheet-strategy-buttons">
                        <button class="sheet-dialog-btn sheet-dialog-btn-primary" data-strategy="keep_local">Keep My Changes</button>
                        <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-strategy="keep_remote">Accept Remote Changes</button>
                        <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-strategy="merge">Smart Merge</button>
                        <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-strategy="manual">Resolve Manually</button>
                    </div>
                </div>
                <div class="sheet-conflict-actions">
                    <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-action="backup">Create Backup</button>
                    <button class="sheet-dialog-btn sheet-dialog-btn-secondary" data-action="cancel">Cancel</button>
                </div>
            </div>
        `;

        const handleClick = async (e: Event) => {
            const target = e.target as HTMLElement;
            const strategy = target.getAttribute('data-strategy');
            const action = target.getAttribute('data-action');
            
            if (strategy) {
                await this.resolveConflictsWithStrategy(conflict, strategy as any);
                dialog.remove();
            } else if (action === 'backup') {
                await this.createBackupBeforeResolving();
                this.showToast('Backup created successfully');
            } else if (action === 'cancel') {
                dialog.remove();
            }
        };

        dialog.addEventListener('click', handleClick);
        document.body.appendChild(dialog);
    }

    private getConflictTypeDescription(type: ConflictInfo['type']): string {
        switch (type) {
            case 'file_modified': return 'File was modified externally';
            case 'concurrent_edit': return 'Concurrent editing detected';
            case 'version_mismatch': return 'Version mismatch detected';
            default: return 'Unknown conflict type';
        }
    }

    private renderConflictedCells(conflict: ConflictInfo): string {
        return conflict.conflictedCells.slice(0, 5).map(cell => {
            const localValue = this.getCellDisplayValue(cell.localValue);
            const remoteValue = this.getCellDisplayValue(cell.remoteValue);
            const cellRef = this.getCellReference(cell.row, cell.col);
            
            return `
                <div class="sheet-conflict-cell">
                    <strong>${cellRef}:</strong>
                    <div class="sheet-conflict-values">
                        <div class="sheet-conflict-value local">
                            <label>Your change:</label>
                            <span>"${localValue}"</span>
                        </div>
                        <div class="sheet-conflict-value remote">
                            <label>Remote change:</label>
                            <span>"${remoteValue}"</span>
                        </div>
                    </div>
                </div>
            `;
        }).join('') + (conflict.conflictedCells.length > 5 ? `<p>...and ${conflict.conflictedCells.length - 5} more</p>` : '');
    }

    private getCellDisplayValue(cellValue: any): string {
        if (!cellValue) return '(empty)';
        if (cellValue.f) return cellValue.f;
        return String(cellValue.v || cellValue.m || '');
    }

    private getCellReference(row: number, col: number): string {
        return `${this.getColumnName(col)}${row + 1}`;
    }

    private async resolveConflictsWithStrategy(conflict: ConflictInfo, strategy: 'keep_local' | 'keep_remote' | 'merge' | 'manual'): Promise<void> {
        if (!this.dataManager) return;

        let resolution: ConflictResolution;

        switch (strategy) {
            case 'keep_local':
                resolution = this.createKeepLocalResolution(conflict);
                break;
            case 'keep_remote':
                resolution = this.createKeepRemoteResolution(conflict);
                break;
            case 'merge':
                resolution = this.createMergeResolution(conflict);
                break;
            case 'manual':
                await this.showManualResolutionDialog(conflict);
                return;
            default:
                return;
        }

        const success = await this.dataManager.resolveConflicts(resolution);
        if (success) {
            this.render();
            this.showToast(`Conflicts resolved using "${strategy}" strategy`);
        } else {
            this.showToast('Failed to resolve conflicts');
        }
    }

    private createKeepLocalResolution(conflict: ConflictInfo): ConflictResolution {
        return {
            strategy: 'keep_local',
            resolvedCells: conflict.conflictedCells.map(cell => ({
                row: cell.row,
                col: cell.col,
                resolvedValue: cell.localValue,
                reasoning: 'Kept local changes'
            })),
            timestamp: new Date()
        };
    }

    private createKeepRemoteResolution(conflict: ConflictInfo): ConflictResolution {
        return {
            strategy: 'keep_remote',
            resolvedCells: conflict.conflictedCells.map(cell => ({
                row: cell.row,
                col: cell.col,
                resolvedValue: cell.remoteValue,
                reasoning: 'Accepted remote changes'
            })),
            timestamp: new Date()
        };
    }

    private createMergeResolution(conflict: ConflictInfo): ConflictResolution {
        return {
            strategy: 'merge',
            resolvedCells: conflict.conflictedCells.map(cell => {
                // Simple merge strategy: concatenate non-empty values
                const localStr = this.getCellDisplayValue(cell.localValue);
                const remoteStr = this.getCellDisplayValue(cell.remoteValue);
                
                let mergedValue: any = cell.localValue;
                let reasoning = 'Kept local value';

                if (localStr === '(empty)' && remoteStr !== '(empty)') {
                    mergedValue = cell.remoteValue;
                    reasoning = 'Used remote value (local was empty)';
                } else if (localStr !== '(empty)' && remoteStr === '(empty)') {
                    mergedValue = cell.localValue;
                    reasoning = 'Used local value (remote was empty)';
                } else if (localStr !== remoteStr) {
                    // Both have values - try to merge intelligently
                    if (localStr.includes(remoteStr) || remoteStr.includes(localStr)) {
                        mergedValue = localStr.length > remoteStr.length ? cell.localValue : cell.remoteValue;
                        reasoning = 'Used longer value';
                    } else {
                        mergedValue = { v: `${localStr} | ${remoteStr}`, m: `${localStr} | ${remoteStr}` };
                        reasoning = 'Concatenated both values';
                    }
                }

                return {
                    row: cell.row,
                    col: cell.col,
                    resolvedValue: mergedValue,
                    reasoning
                };
            }),
            timestamp: new Date()
        };
    }

    private async showManualResolutionDialog(conflict: ConflictInfo): Promise<void> {
        // For now, just show the first conflict dialog again - in a full implementation,
        // this would show a detailed editor for each conflicted cell
        this.showToast('Manual resolution not yet implemented. Please choose an automatic strategy.');
    }

    private async createBackupBeforeResolving(): Promise<void> {
        if (!this.dataManager) return;

        try {
            const backup = this.dataManager.createBackup();
            const backupFileName = `${this.file?.basename || 'sheet'}_backup_${Date.now()}.json`;
            
            // Save backup to a backup folder (in a real implementation, you'd create this folder)
            // For now, we'll just store it in localStorage as a simple backup
            localStorage.setItem(`sheet_backup_${this.file?.path}`, backup);
            
            console.log('Backup created:', backupFileName);
        } catch (error) {
            console.error('Failed to create backup:', error);
            throw error;
        }
    }
}