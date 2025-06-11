import { TextFileView, TFile, WorkspaceLeaf } from 'obsidian';
import { SheetParser, SheetData, DataManager, ValidationRule, AutoSaveState } from './sheetParser';
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
        this.sheetData = SheetParser.parse(data);
        this.dataManager = new DataManager(this.sheetData, 100, {
            enabled: true,
            intervalMs: 30000, // 30 seconds
            debounceMs: 2000,  // 2 seconds after last change
            maxRetries: 3
        });
        
        // Set up auto-save callback
        this.dataManager.setSaveCallback(() => this.performSave());
        
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
            
            if (autoSaveState.isSaving) {
                toolbar.createEl('span', {
                    text: '💾 Saving...',
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
            { label: 'Edit', action: () => this.toggleEditModeAndEdit(cellEl, cellInfo) },
            { label: isLongText ? 'Edit as Text' : 'Edit as Multi-line', action: () => this.editCellWithType(cellEl, cellInfo, isLongText ? 'text' : 'multiline') },
            { label: 'Copy', action: () => this.copyCellValue(cellInfo.value) },
            { label: 'Clear', action: () => this.clearCell(cellInfo) },
            { label: autoSaveEnabled ? 'Disable Auto-save' : 'Enable Auto-save', action: () => this.toggleAutoSave() },
            { label: 'Cancel', action: () => menu.remove() }
        ];

        menuItems.forEach(item => {
            const button = menu.createEl('button', {
                text: item.label,
                cls: 'sheet-context-menu-item'
            });
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

    private clearCell(cellInfo: any): void {
        if (this.dataManager) {
            const sheet = this.sheetData[cellInfo.sheetIndex];
            this.dataManager.updateCell(sheet.id, cellInfo.row, cellInfo.col, '');
            this.updateDisplayText();
            this.render();
            this.showToast('Cell cleared');
        }
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
            }
        } catch (error) {
            this.showToast('Error: ' + (error instanceof Error ? error.message : 'Unknown error'));
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

        const serialized = SheetParser.serialize(this.dataManager.getSheetData());
        await this.app.vault.modify(this.file, serialized);
        this.dataManager.markClean();
        this.updateDisplayText();
        
        // Re-render to update status indicators
        this.render();
    }

    private async saveFile(): Promise<void> {
        if (!this.file || !this.dataManager?.isDirty()) return;

        try {
            await this.performSave();
            this.showToast('File saved manually');
        } catch (error) {
            console.error('Failed to save file:', error);
            this.showToast('Save failed: ' + (error instanceof Error ? error.message : 'Unknown error'));
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
            try {
                await this.dataManager.forceAutoSave();
                this.showToast('Auto-save completed');
            } catch (error) {
                this.showToast('Auto-save failed');
            }
        }
    }
}