import { TextFileView, TFile, WorkspaceLeaf } from 'obsidian';
import { SheetParser, SheetData, DataManager, ValidationRule, AutoSaveState, ConflictInfo, ConflictResolution, CellFormat } from './sheetParser';
import { FormulaEngine } from './formulaEngine';
import { FormatManager } from './formatManager';

export const VIEW_TYPE_SHEET = 'sheet-view';

// Debug configuration
const DEBUG = false; // Set to true during development

function debugLog(message: string, data?: any) {
    if (DEBUG) {
        if (data) {
            console.debug(`[SheetView] ${message}`, data);
        } else {
            console.debug(`[SheetView] ${message}`);
        }
    }
}

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
    private isLargeFile: boolean = false;
    private forceEditingEnabled: boolean = false;
    private copiedFormat: CellFormat | null = null;
    private readonly fileSizeThresholds = {
        mobile: 100000, // 100KB
        tablet: 250000, // 250KB
        desktop: 500000, // 500KB
        maxCells: 1000 // Maximum number of cells for editing
    };

    constructor(leaf: WorkspaceLeaf) {
        super(leaf);
        debugLog('Constructor called');
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
        debugLog('setViewData called');
        this.data = data;
        this.lastFileContent = data;
        
        // Check file size and cell count before proceeding
        this.checkFileSize(data);
        
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
        
        debugLog('Parsed sheet data');
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
        debugLog('Render called');
        this.contentEl.empty();
        this.contentEl.addClass('mobile-sheet-viewer');

        this.renderToolbar();
        this.setupSwipeGestures();

        if (this.sheetData.length === 0) {
            debugLog('No sheet data found');
            this.contentEl.createEl('div', { 
                text: 'No sheet data found',
                cls: 'sheet-error' 
            });
            return;
        }

        this.sheetData.forEach((sheet, index) => {
            debugLog('Rendering sheet', {index, name: sheet.name});
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
        
        if (this.isLargeFile && !this.forceEditingEnabled) {
            editToggle.disabled = true;
            editToggle.title = 'Editing disabled for large files. Click the file size warning for options.';
            editToggle.addClass('sheet-edit-disabled');
        } else {
            editToggle.addEventListener('click', () => this.toggleEditMode());
        }

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

        // Formatting toolbar (only when editing)
        if (this.isEditing && (!this.isLargeFile || this.forceEditingEnabled)) {
            this.renderFormattingToolbar(toolbar);
        }

        // File size warning indicator
        if (this.isLargeFile) {
            const fileSizeWarning = toolbar.createEl('button', {
                text: '⚠️ Large File',
                cls: 'sheet-file-size-warning'
            });
            fileSizeWarning.title = 'File is too large for optimal editing';
            fileSizeWarning.addEventListener('click', () => this.showFileSizeDialog());
        }

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
                const cellValue = SheetParser.getFormattedCellValue(cellData);
                const actualRow = rowMap[r];
                const actualCol = colMap[c];
                
                const td = row.createEl('td', { 
                    cls: 'sheet-cell',
                    text: cellValue
                });

                // Apply cell formatting styles
                if (cellData?.s) {
                    const cellStyle = SheetParser.generateCellStyle(cellData.s);
                    if (cellStyle) {
                        td.setAttribute('style', cellStyle);
                    }
                }

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
                        if (cellState.isModified || cellState.isFormatModified) {
                            td.addClass('sheet-cell-modified');
                        }
                        if (!cellState.isValid) {
                            td.addClass('sheet-cell-invalid');
                            td.title = cellState.validationError || 'Invalid value';
                        }
                    }
                }

                // Add interaction handlers only if editing is enabled
                if (!this.isLargeFile || this.forceEditingEnabled) {
                    this.setupCellInteractionHandlers(td, sheetIndex, actualRow, actualCol);
                } else {
                    // For large files, only allow viewing
                    td.addEventListener('click', () => {
                        this.showToast('Editing disabled for large files');
                    });
                }
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
        
        cellEl.empty();

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

    private showFormulaHelp(input: HTMLInputElement | HTMLTextAreaElement): void {
        // Remove existing help
        this.hideFormulaHelp();
        
        const help = document.createElement('div');
        help.className = 'sheet-formula-help';
        const functionsDiv = help.createDiv({cls: 'formula-functions'});
        const functionsStrong = functionsDiv.createEl('strong');
        functionsStrong.textContent = 'Functions:';
        functionsDiv.appendText(' SUM(A1:A5), COUNT(A1:A5), AVERAGE(A1:A5), MIN(A1:A5), MAX(A1:A5), ');
        functionsDiv.appendText('IF(condition, true_value, false_value), CONCATENATE(text1, text2), ');
        functionsDiv.appendText('LEN(text), UPPER(text), LOWER(text), ROUND(number, digits)');
        
        const examplesDiv = help.createDiv({cls: 'formula-examples'});
        const examplesStrong = examplesDiv.createEl('strong');
        examplesStrong.textContent = 'Examples:';
        examplesDiv.appendText(' =A1+B1, =SUM(A1:A5), =IF(A1>10, "High", "Low")');
        
        const rect = input.getBoundingClientRect();
        help.style.setProperty('--help-top', (rect.bottom + 5) + 'px');
        help.style.setProperty('--help-left', rect.left + 'px');
        
        document.body.appendChild(help);
    }

    private hideFormulaHelp(): void {
        const existing = document.querySelector('.sheet-formula-help');
        if (existing) {
            existing.remove();
        }
    }

    private setupMobileInput(input: HTMLInputElement | HTMLTextAreaElement): void {
        // Device-specific optimizations
        const deviceType = this.detectDeviceType();
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
        const isAndroid = /Android/.test(navigator.userAgent);
        
        // Add device-specific classes
        if (isIOS) {
            input.classList.add('sheet-cell-input-ios');
        } else if (isAndroid) {
            input.classList.add('sheet-cell-input-android');
        } else {
            input.style.fontSize = '14px';
        }
        
        // Handle virtual keyboard with device-specific delays
        input.addEventListener('focus', () => {
            this.handleVirtualKeyboard(true, deviceType);
        });
        
        input.addEventListener('blur', () => {
            this.handleVirtualKeyboard(false, deviceType);
        });

        // Auto-resize input based on content with debouncing
        let resizeTimeout: number | null = null;
        input.addEventListener('input', () => {
            if (resizeTimeout) clearTimeout(resizeTimeout);
            resizeTimeout = window.setTimeout(() => {
                if (input instanceof HTMLTextAreaElement) {
                    this.autoResizeTextarea(input);
                } else {
                    this.autoResizeInput(input);
                }
            }, 100);
        });

        // Initial resize with device-specific delay
        const initialDelay = isIOS ? 200 : isAndroid ? 150 : 50;
        setTimeout(() => {
            if (input instanceof HTMLTextAreaElement) {
                this.autoResizeTextarea(input);
            } else {
                this.autoResizeInput(input);
            }
        }, initialDelay);
        
        // Add touch event handling for better mobile interaction
        this.setupInputTouchHandling(input);
    }
    
    private setupInputTouchHandling(input: HTMLInputElement | HTMLTextAreaElement): void {
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
        const isAndroid = /Android/.test(navigator.userAgent);
        
        // Prevent input from losing focus on touch events
        input.addEventListener('touchstart', (e) => {
            e.stopPropagation();
        });
        
        input.addEventListener('touchend', (e) => {
            e.stopPropagation();
            // iOS-specific: ensure input retains focus
            if (isIOS && input !== document.activeElement) {
                setTimeout(() => input.focus(), 50);
            }
        });
        
        // Handle paste operations with better mobile support
        input.addEventListener('paste', (e) => {
            e.stopPropagation();
            // Allow default paste behavior but ensure proper resizing
            setTimeout(() => {
                if (input instanceof HTMLTextAreaElement) {
                    this.autoResizeTextarea(input);
                } else {
                    this.autoResizeInput(input);
                }
            }, 100);
        });
        
        // Improved selection handling for mobile
        if (isIOS || isAndroid) {
            input.addEventListener('selectionchange', () => {
                // Selection styles are handled by CSS classes
                // No additional JS styling needed
            });
        }
    }

    private handleVirtualKeyboard(isShowing: boolean, deviceType: 'mobile' | 'tablet' | 'desktop' = 'mobile'): void {
        const viewport = document.querySelector('meta[name=viewport]');
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
        const isAndroid = /Android/.test(navigator.userAgent);
        
        if (!viewport) return;

        if (isShowing) {
            // Device-specific viewport adjustments
            if (isIOS) {
                viewport.setAttribute('content', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover');
            } else if (isAndroid) {
                viewport.setAttribute('content', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
            } else {
                viewport.setAttribute('content', 'width=device-width, initial-scale=1.0, maximum-scale=3.0');
            }
            
            // Add keyboard-open class for CSS adjustments
            document.body.classList.add('keyboard-open');
            
            // Scroll to editing cell with device-specific timing
            if (this.currentEditor) {
                const scrollDelay = isIOS ? 500 : isAndroid ? 300 : 100;
                setTimeout(() => {
                    if (this.currentEditor) {
                        this.currentEditor.scrollIntoView({ 
                            behavior: 'smooth', 
                            block: deviceType === 'mobile' ? 'start' : 'center',
                            inline: 'nearest'
                        });
                    }
                }, scrollDelay);
                
                // Additional iOS-specific handling
                if (isIOS) {
                    // Add class to handle keyboard open state
                    document.body.classList.add('sheet-keyboard-open');
                }
            }
        } else {
            // Restore normal viewport
            if (isIOS) {
                viewport.setAttribute('content', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
                // Remove keyboard open class
                document.body.classList.remove('sheet-keyboard-open');
            } else {
                viewport.setAttribute('content', 'width=device-width, initial-scale=1.0');
            }
            
            // Remove keyboard-open class
            document.body.classList.remove('keyboard-open');
        }
    }

    private autoResizeInput(input: HTMLInputElement): void {
        const deviceType = this.detectDeviceType();
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
        
        try {
            // Create temporary span to measure text width
            const span = document.createElement('span');
            span.className = 'sheet-text-measure';
            // Copy computed styles for accurate measurement
            span.style.fontSize = window.getComputedStyle(input).fontSize;
            span.style.fontFamily = window.getComputedStyle(input).fontFamily;
            span.style.fontWeight = window.getComputedStyle(input).fontWeight;
            span.textContent = input.value || input.placeholder || 'Placeholder';
            
            document.body.appendChild(span);
            const textWidth = span.offsetWidth;
            document.body.removeChild(span);
            
            // Device-specific width calculations
            const minWidth = deviceType === 'mobile' ? 80 : 60;
            const maxWidth = deviceType === 'mobile' ? 
                Math.min(window.innerWidth * 0.9, 300) : 
                Math.min(window.innerWidth * 0.7, 400);
            
            const padding = isIOS ? 24 : 20;
            const calculatedWidth = Math.max(textWidth + padding, minWidth);
            
            // Use CSS custom properties for dynamic width
            const finalWidth = Math.min(calculatedWidth, maxWidth);
            input.style.setProperty('--input-width', finalWidth + 'px');
            
            // Ensure input stays within viewport
            if (input.getBoundingClientRect().right > window.innerWidth) {
                const rect = input.getBoundingClientRect();
                const overflow = rect.right - window.innerWidth + 10;
                const adjustedWidth = Math.max(calculatedWidth - overflow, minWidth);
                input.style.setProperty('--input-width', adjustedWidth + 'px');
            }
        } catch (error) {
            // Fallback if measurement fails
            // Auto-resize fallback is not a critical error
            debugLog('Auto-resize failed, using fallback:', error);
            const fallbackWidth = deviceType === 'mobile' ? 120 : 80;
            input.style.setProperty('--input-width', fallbackWidth + 'px');
        }
    }

    private autoResizeTextarea(textarea: HTMLTextAreaElement): void {
        const deviceType = this.detectDeviceType();
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
        
        try {
            // Reset height to auto to get the correct scrollHeight
            textarea.style.height = 'auto';
            
            // Device-specific height calculations
            const minHeight = deviceType === 'mobile' ? 80 : 60;
            const maxHeight = deviceType === 'mobile' ? 
                Math.min(window.innerHeight * 0.4, 250) : 
                Math.min(window.innerHeight * 0.5, 300);
            
            const scrollHeight = textarea.scrollHeight;
            const newHeight = Math.max(Math.min(scrollHeight, maxHeight), minHeight);
            
            textarea.style.setProperty('--textarea-height', newHeight + 'px');
            
            // Adjust width with device-specific constraints
            const maxWidth = deviceType === 'mobile' ? 
                Math.min(window.innerWidth * 0.9, 350) : 
                Math.min(window.innerWidth * 0.8, 450);
                
            const minWidth = deviceType === 'mobile' ? 200 : 150;
            
            if (textarea.scrollWidth > textarea.clientWidth) {
                const padding = isIOS ? 24 : 20;
                const newWidth = Math.min(Math.max(textarea.scrollWidth + padding, minWidth), maxWidth);
                textarea.style.setProperty('--textarea-width', newWidth + 'px');
            }
            
            // Ensure textarea stays within viewport bounds
            const rect = textarea.getBoundingClientRect();
            if (rect.bottom > window.innerHeight) {
                textarea.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            }
            
            if (rect.right > window.innerWidth) {
                const overflow = rect.right - window.innerWidth + 10;
                const currentWidth = parseInt(getComputedStyle(textarea).getPropertyValue('--textarea-width')) || textarea.offsetWidth;
                const adjustedWidth = Math.max(currentWidth - overflow, minWidth);
                textarea.style.setProperty('--textarea-width', adjustedWidth + 'px');
            }
        } catch (error) {
            // Fallback if resize fails
            // Textarea auto-resize fallback is not a critical error
            debugLog('Textarea auto-resize failed, using fallback:', error);
            textarea.style.setProperty('--textarea-height', (deviceType === 'mobile' ? 100 : 80) + 'px');
            textarea.style.setProperty('--textarea-width', (deviceType === 'mobile' ? 250 : 200) + 'px');
        }
    }

    private setupCellTouchHandlers(cellEl: HTMLElement, sheetIndex: number, row: number, col: number): void {
        let tapStartTime = 0;
        let preventClick = false;
        const deviceType = this.detectDeviceType();
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
        
        // Device-specific touch thresholds
        const moveThreshold = deviceType === 'mobile' ? 15 : 10;
        const tapThreshold = isIOS ? 250 : 200;
        
        cellEl.addEventListener('touchstart', (e) => {
            e.preventDefault(); // Prevent default to avoid double events
            tapStartTime = Date.now();
            preventClick = false;
            
            this.touchStartPos = {
                x: e.touches[0].clientX,
                y: e.touches[0].clientY
            };
            this.isLongPress = false;
            
            // Start long press timer with device-specific delay
            const longPressDelay = deviceType === 'mobile' ? 600 : this.longPressDelay;
            this.longPressTimer = window.setTimeout(() => {
                this.isLongPress = true;
                this.handleLongPress(cellEl, sheetIndex, row, col);
                this.triggerHapticFeedback();
                preventClick = true;
            }, longPressDelay);
        }, { passive: false });

        cellEl.addEventListener('touchmove', (e) => {
            if (!this.touchStartPos) return;
            
            const touch = e.touches[0];
            const deltaX = Math.abs(touch.clientX - this.touchStartPos.x);
            const deltaY = Math.abs(touch.clientY - this.touchStartPos.y);
            
            // Cancel long press if moved too much
            if (deltaX > moveThreshold || deltaY > moveThreshold) {
                this.clearLongPressTimer();
                preventClick = true;
            }
        });

        cellEl.addEventListener('touchend', (e) => {
            e.preventDefault(); // Prevent ghost clicks
            this.clearLongPressTimer();
            
            if (!this.isLongPress && !preventClick && this.touchStartPos) {
                const tapDuration = Date.now() - tapStartTime;
                
                if (tapDuration < tapThreshold) { // Quick tap
                    // Use timeout to ensure stable editing initiation
                    setTimeout(() => {
                        this.startCellEdit(cellEl, sheetIndex, row, col);
                        this.triggerHapticFeedback('light');
                    }, 50);
                }
            }
            
            this.touchStartPos = null;
            
            // Reset click prevention after a delay
            setTimeout(() => {
                preventClick = false;
            }, 300);
        }, { passive: false });

        // Handle mouse events for desktop compatibility (but prevent if touch was used)
        cellEl.addEventListener('click', (e) => {
            if (!('ontouchstart' in window) && !preventClick) {
                this.startCellEdit(cellEl, sheetIndex, row, col);
            }
        });
        
        // Additional iOS-specific handling
        if (isIOS) {
            cellEl.addEventListener('touchcancel', () => {
                this.clearLongPressTimer();
                this.touchStartPos = null;
            });
        }
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
            { label: '---', action: () => {}, shortcut: '' }, // Separator
            { label: 'Format Cell', action: () => this.showFormattingDialog(cellInfo), shortcut: 'Ctrl+1' },
            { label: 'Bold', action: () => this.toggleCellFormat(cellInfo, 'bold'), shortcut: 'Ctrl+B' },
            { label: 'Italic', action: () => this.toggleCellFormat(cellInfo, 'italic'), shortcut: 'Ctrl+I' },
            { label: 'Clear Format', action: () => this.clearCellFormat(cellInfo), shortcut: '' },
            { label: '---', action: () => {}, shortcut: '' }, // Separator
            { label: 'Copy', action: () => this.copyCellValue(cellInfo.value), shortcut: 'Ctrl+C' },
            { label: 'Paste', action: () => this.pasteCellValue(cellInfo), shortcut: 'Ctrl+V' },
            { label: 'Copy Format', action: () => this.copyCellFormat(cellInfo), shortcut: '' },
            { label: 'Paste Format', action: () => this.pasteCellFormat(cellInfo), shortcut: '' },
            { label: '---', action: () => {}, shortcut: '' }, // Separator
            { label: 'Clear Cell', action: () => this.clearCell(cellInfo), shortcut: 'Del' },
            { label: 'Delete Row', action: () => this.deleteRow(cellInfo), shortcut: 'Ctrl+Shift+-' },
            { label: 'Delete Column', action: () => this.deleteColumn(cellInfo), shortcut: 'Ctrl+Alt+-' },
            { label: '---', action: () => {}, shortcut: '' }, // Separator
            { label: autoSaveEnabled ? 'Disable Auto-save' : 'Enable Auto-save', action: () => this.toggleAutoSave() },
            { label: 'Cancel', action: () => menu.remove(), shortcut: 'Esc' }
        ];

        menuItems.forEach(item => {
            if (item.label === '---') {
                menu.createEl('div', { cls: 'sheet-context-menu-separator' });
                return;
            }
            
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
        
        cellEl.empty();

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
        for (let i = 0; i < tables.length; i++) {
            const table = tables[i];
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
                case 'b':
                    e.preventDefault();
                    this.applyQuickFormat('bold');
                    break;
                case 'i':
                    e.preventDefault();
                    this.applyQuickFormat('italic');
                    break;
                case 'u':
                    e.preventDefault();
                    this.applyQuickFormat('underline');
                    break;
                case '1':
                    e.preventDefault();
                    if (this.selectedCell) {
                        this.showFormattingDialog({
                            sheetIndex: this.selectedCell.sheetIndex,
                            row: this.selectedCell.row,
                            col: this.selectedCell.col,
                            value: this.getCellDisplayValue(this.selectedCell)
                        });
                    }
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
        
        // Clean up mobile-specific event listeners
        this.cleanupInputEventListeners(this.currentEditor);

        try {
            const success = this.dataManager.updateCell(sheet.id, row, col, newValue);
            if (success) {
                this.updateDisplayText();
                this.triggerHapticFeedback('light');
            } else {
                this.showToast('Invalid formula or value');
                this.showFallbackForUnsupportedOperation('cell update', newValue);
                this.triggerHapticFeedback('heavy');
            }
        } catch (error) {
            this.showToast('Error: ' + (error instanceof Error ? error.message : 'Unknown error'));
            this.showFallbackForUnsupportedOperation('cell update', newValue, error);
            this.triggerHapticFeedback('heavy');
        }

        this.currentEditor = null;
        this.editingCell = null;
        
        // Restore viewport and keyboard state
        this.handleVirtualKeyboard(false);
        
        this.render();
    }
    
    private cleanupInputEventListeners(input: HTMLInputElement | HTMLTextAreaElement): void {
        // Remove mobile-specific event listeners to prevent memory leaks
        const newInput = input.cloneNode(true) as HTMLInputElement | HTMLTextAreaElement;
        input.parentNode?.replaceChild(newInput, input);
    }

    private cancelCellEdit(): void {
        this.hideFormulaHelp();
        
        // Clean up mobile-specific event listeners
        if (this.currentEditor) {
            this.cleanupInputEventListeners(this.currentEditor);
        }
        
        this.currentEditor = null;
        this.editingCell = null;
        
        // Restore viewport and keyboard state
        this.handleVirtualKeyboard(false);
        
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
            // Keep error logging for save failures as they're critical
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

    // Cell formatting methods
    private showFormattingDialog(cellInfo: any): void {
        const sheet = this.sheetData[cellInfo.sheetIndex];
        const currentFormat = this.dataManager?.getCellFormat(sheet.id, cellInfo.row, cellInfo.col) || null;
        
        // Remove existing dialog
        const existing = document.querySelector('.sheet-format-dialog');
        if (existing) existing.remove();

        const dialog = document.createElement('div');
        dialog.className = 'sheet-format-dialog';
        
        // Create dialog content using DOM API
        const content = dialog.createDiv({cls: 'sheet-format-dialog-content'});
        const h3 = content.createEl('h3');
        h3.textContent = 'Format Cell';
        
        // Create tabs
        const tabsDiv = content.createDiv({cls: 'sheet-format-tabs'});
        const quickTab = tabsDiv.createEl('button', {cls: 'sheet-format-tab active', attr: {'data-tab': 'quick'}});
        quickTab.textContent = 'Quick';
        const fontTab = tabsDiv.createEl('button', {cls: 'sheet-format-tab', attr: {'data-tab': 'font'}});
        fontTab.textContent = 'Font';
        const numberTab = tabsDiv.createEl('button', {cls: 'sheet-format-tab', attr: {'data-tab': 'number'}});
        numberTab.textContent = 'Number';
        const borderTab = tabsDiv.createEl('button', {cls: 'sheet-format-tab', attr: {'data-tab': 'border'}});
        borderTab.textContent = 'Border';
        
        // Create content area
        const formatContent = content.createDiv({cls: 'sheet-format-content'});
        
        // Create panels
        const quickPanel = formatContent.createDiv({cls: 'sheet-format-panel active', attr: {'data-panel': 'quick'}});
        this.renderQuickFormatPanelDOM(quickPanel, currentFormat);
        
        const fontPanel = formatContent.createDiv({cls: 'sheet-format-panel', attr: {'data-panel': 'font'}});
        this.renderFontPanelDOM(fontPanel, currentFormat);
        
        const numberPanel = formatContent.createDiv({cls: 'sheet-format-panel', attr: {'data-panel': 'number'}});
        this.renderNumberPanelDOM(numberPanel, currentFormat);
        
        const borderPanel = formatContent.createDiv({cls: 'sheet-format-panel', attr: {'data-panel': 'border'}});
        this.renderBorderPanelDOM(borderPanel, currentFormat);
        
        // Create actions
        const actionsDiv = content.createDiv({cls: 'sheet-format-actions'});
        const cancelBtn = actionsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-action': 'cancel'}});
        cancelBtn.textContent = 'Cancel';
        const applyBtn = actionsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-primary', attr: {'data-action': 'apply'}});
        applyBtn.textContent = 'Apply';

        this.setupFormatDialogHandlers(dialog, cellInfo);
        document.body.appendChild(dialog);
    }

    // New DOM-based render methods
    private renderQuickFormatPanelDOM(panel: HTMLElement, currentFormat: CellFormat | null): void {
        const presets = FormatManager.getMobileQuickFormats();
        const presetsDiv = panel.createDiv({cls: 'sheet-format-presets'});
        
        presets.forEach(preset => {
            const presetBtn = presetsDiv.createEl('button', {cls: 'sheet-format-preset', attr: {'data-preset': preset.id}});
            const previewDiv = presetBtn.createDiv({cls: 'sheet-format-preset-preview', attr: {style: SheetParser.generateCellStyle(preset.format)}});
            previewDiv.textContent = preset.name;
            const nameSpan = presetBtn.createSpan({cls: 'sheet-format-preset-name'});
            nameSpan.textContent = preset.name;
        });
        
        const colorsDiv = panel.createDiv({cls: 'sheet-format-colors'});
        const colorLabel = colorsDiv.createEl('label');
        colorLabel.textContent = 'Colors:';
        
        const textPalette = colorsDiv.createDiv({cls: 'sheet-color-palette'});
        FormatManager.getColorPalette().forEach(color => {
            textPalette.createEl('button', {
                cls: 'sheet-color-btn',
                attr: {
                    'data-color': color,
                    'data-type': 'text',
                    'style': `background-color: ${color}`,
                    'title': `Text: ${color}`
                }
            });
        });
        
        const bgPalette = colorsDiv.createDiv({cls: 'sheet-color-palette'});
        FormatManager.getColorPalette().forEach(color => {
            bgPalette.createEl('button', {
                cls: 'sheet-color-btn',
                attr: {
                    'data-color': color,
                    'data-type': 'background',
                    'style': `background-color: ${color}; border: 2px solid #fff; box-shadow: 0 0 0 1px #ccc`,
                    'title': `Background: ${color}`
                }
            });
        });
    }
    
    private renderFontPanelDOM(panel: HTMLElement, currentFormat: CellFormat | null): void {
        // Font family section
        const fontSection = panel.createDiv({cls: 'sheet-format-section'});
        const fontLabel = fontSection.createEl('label');
        fontLabel.textContent = 'Font Family:';
        const fontSelect = fontSection.createEl('select', {cls: 'sheet-format-select', attr: {'data-property': 'fontFamily'}});
        fontSelect.createEl('option', {value: '', text: 'Default'});
        const fonts = [
            ['Arial, sans-serif', 'Arial'],
            ["'Times New Roman', serif", 'Times New Roman'],
            ["'Courier New', monospace", 'Courier New'],
            ['Helvetica, sans-serif', 'Helvetica']
        ];
        fonts.forEach(([value, text]) => {
            const option = fontSelect.createEl('option', {value, text});
            if (currentFormat?.fontFamily === value) option.selected = true;
        });
        
        // Font size section
        const sizeSection = panel.createDiv({cls: 'sheet-format-section'});
        const sizeLabel = sizeSection.createEl('label');
        sizeLabel.textContent = 'Font Size:';
        sizeSection.createEl('input', {
            type: 'number',
            cls: 'sheet-format-input',
            attr: {
                'data-property': 'fontSize',
                'value': String(currentFormat?.fontSize || 14),
                'min': '8',
                'max': '72'
            }
        });
        
        // Text style section
        const styleSection = panel.createDiv({cls: 'sheet-format-section'});
        const styleLabel = styleSection.createEl('label');
        styleLabel.textContent = 'Text Style:';
        const buttonsDiv = styleSection.createDiv({cls: 'sheet-format-buttons'});
        
        const boldBtn = buttonsDiv.createEl('button', {
            cls: 'sheet-format-toggle sheet-format-bold',
            attr: {'data-property': 'fontWeight', 'data-value': 'bold'}
        });
        boldBtn.textContent = 'B';
        if (currentFormat?.fontWeight === 'bold') boldBtn.classList.add('active');
        
        const italicBtn = buttonsDiv.createEl('button', {
            cls: 'sheet-format-toggle sheet-format-italic',
            attr: {'data-property': 'fontStyle', 'data-value': 'italic'}
        });
        italicBtn.textContent = 'I';
        if (currentFormat?.fontStyle === 'italic') italicBtn.classList.add('active');
        
        const underlineBtn = buttonsDiv.createEl('button', {
            cls: 'sheet-format-toggle sheet-format-underline',
            attr: {'data-property': 'textDecoration', 'data-value': 'underline'}
        });
        underlineBtn.textContent = 'U';
        if (currentFormat?.textDecoration === 'underline') underlineBtn.classList.add('active');
        
        // Text alignment section
        const alignSection = panel.createDiv({cls: 'sheet-format-section'});
        const alignLabel = alignSection.createEl('label');
        alignLabel.textContent = 'Text Alignment:';
        const alignButtons = alignSection.createDiv({cls: 'sheet-format-buttons'});
        
        ['left', 'center', 'right'].forEach(align => {
            const btn = alignButtons.createEl('button', {
                cls: 'sheet-format-toggle',
                attr: {'data-property': 'textAlign', 'data-value': align}
            });
            btn.textContent = align.charAt(0).toUpperCase() + align.slice(1);
            if (currentFormat?.textAlign === align) btn.classList.add('active');
        });
    }
    
    private renderNumberPanelDOM(panel: HTMLElement, currentFormat: CellFormat | null): void {
        // Number type section
        const typeSection = panel.createDiv({cls: 'sheet-format-section'});
        const typeLabel = typeSection.createEl('label');
        typeLabel.textContent = 'Number Type:';
        const typeSelect = typeSection.createEl('select', {cls: 'sheet-format-select', attr: {'data-property': 'numberType'}});
        
        const types = [
            ['general', 'General'],
            ['number', 'Number'],
            ['currency', 'Currency'],
            ['percent', 'Percent'],
            ['date', 'Date'],
            ['text', 'Text']
        ];
        types.forEach(([value, text]) => {
            const option = typeSelect.createEl('option', {value, text});
            if (currentFormat?.numberFormat?.type === value) option.selected = true;
        });
        
        // Decimal places section
        const decimalSection = panel.createDiv({cls: 'sheet-format-section'});
        const decimalLabel = decimalSection.createEl('label');
        decimalLabel.textContent = 'Decimal Places:';
        decimalSection.createEl('input', {
            type: 'number',
            cls: 'sheet-format-input',
            attr: {
                'data-property': 'decimalPlaces',
                'value': String(currentFormat?.numberFormat?.decimalPlaces || 2),
                'min': '0',
                'max': '10'
            }
        });
        
        // Currency symbol section
        const currencySection = panel.createDiv({cls: 'sheet-format-section'});
        const currencyLabel = currencySection.createEl('label');
        currencyLabel.textContent = 'Currency Symbol:';
        currencySection.createEl('input', {
            type: 'text',
            cls: 'sheet-format-input',
            attr: {
                'data-property': 'currencySymbol',
                'value': currentFormat?.numberFormat?.currencySymbol || '$',
                'placeholder': '$'
            }
        });
        
        // Thousands separator section
        const thousandsSection = panel.createDiv({cls: 'sheet-format-section'});
        const thousandsLabel = thousandsSection.createEl('label');
        const checkbox = thousandsLabel.createEl('input', {
            type: 'checkbox',
            attr: {'data-property': 'thousandsSeparator'}
        });
        if (currentFormat?.numberFormat?.thousandsSeparator) checkbox.checked = true;
        thousandsLabel.appendText(' Use thousands separator');
    }
    
    private renderBorderPanelDOM(panel: HTMLElement, currentFormat: CellFormat | null): void {
        // Border style section
        const styleSection = panel.createDiv({cls: 'sheet-format-section'});
        const styleLabel = styleSection.createEl('label');
        styleLabel.textContent = 'Border Style:';
        const styleSelect = styleSection.createEl('select', {cls: 'sheet-format-select', attr: {'data-property': 'borderStyle'}});
        
        const styles = [
            ['none', 'None'],
            ['solid', 'Solid'],
            ['dashed', 'Dashed'],
            ['dotted', 'Dotted']
        ];
        styles.forEach(([value, text]) => {
            styleSelect.createEl('option', {value, text});
        });
        
        // Border width section
        const widthSection = panel.createDiv({cls: 'sheet-format-section'});
        const widthLabel = widthSection.createEl('label');
        widthLabel.textContent = 'Border Width:';
        widthSection.createEl('input', {
            type: 'number',
            cls: 'sheet-format-input',
            attr: {
                'data-property': 'borderWidth',
                'value': '1',
                'min': '1',
                'max': '5'
            }
        });
        
        // Border color section
        const colorSection = panel.createDiv({cls: 'sheet-format-section'});
        const colorLabel = colorSection.createEl('label');
        colorLabel.textContent = 'Border Color:';
        colorSection.createEl('input', {
            type: 'color',
            cls: 'sheet-format-input',
            attr: {
                'data-property': 'borderColor',
                'value': '#000000'
            }
        });
        
        // Apply to section
        const applySection = panel.createDiv({cls: 'sheet-format-section'});
        const applyLabel = applySection.createEl('label');
        applyLabel.textContent = 'Apply to:';
        const applyButtons = applySection.createDiv({cls: 'sheet-format-buttons'});
        
        ['all', 'top', 'right', 'bottom', 'left'].forEach(border => {
            const btn = applyButtons.createEl('button', {
                cls: 'sheet-format-toggle',
                attr: {'data-border': border}
            });
            btn.textContent = border.charAt(0).toUpperCase() + border.slice(1);
        });
    }
    

    private setupFormatDialogHandlers(dialog: HTMLElement, cellInfo: any): void {
        let currentFormat: CellFormat = {};
        
        // Tab switching
        dialog.addEventListener('click', (e) => {
            const target = e.target as HTMLElement;
            
            if (target.classList.contains('sheet-format-tab')) {
                const tabName = target.getAttribute('data-tab');
                if (tabName) {
                    // Switch tabs
                    dialog.querySelectorAll('.sheet-format-tab').forEach(tab => tab.classList.remove('active'));
                    dialog.querySelectorAll('.sheet-format-panel').forEach(panel => panel.classList.remove('active'));
                    
                    target.classList.add('active');
                    const panel = dialog.querySelector(`[data-panel="${tabName}"]`);
                    if (panel) panel.classList.add('active');
                }
            }
            
            // Preset selection
            if (target.classList.contains('sheet-format-preset') || target.closest('.sheet-format-preset')) {
                const presetBtn = target.closest('.sheet-format-preset') as HTMLElement;
                const presetId = presetBtn.getAttribute('data-preset');
                if (presetId) {
                    const preset = FormatManager.getPreset(presetId);
                    if (preset) {
                        currentFormat = { ...preset.format };
                        this.highlightSelectedPreset(dialog, presetBtn);
                    }
                }
            }
            
            // Color selection
            if (target.classList.contains('sheet-color-btn')) {
                const color = target.getAttribute('data-color');
                const type = target.getAttribute('data-type');
                if (color && type) {
                    if (type === 'text') {
                        currentFormat.color = color;
                    } else if (type === 'background') {
                        currentFormat.backgroundColor = color;
                    }
                }
            }
            
            // Format toggles
            if (target.classList.contains('sheet-format-toggle')) {
                const property = target.getAttribute('data-property');
                const value = target.getAttribute('data-value');
                
                if (property && value) {
                    target.classList.toggle('active');
                    if (target.classList.contains('active')) {
                        (currentFormat as any)[property] = value;
                    } else {
                        delete (currentFormat as any)[property];
                    }
                }
            }
            
            // Action buttons
            if (target.getAttribute('data-action') === 'apply') {
                this.applyCellFormat(cellInfo, currentFormat);
                dialog.remove();
            } else if (target.getAttribute('data-action') === 'cancel') {
                dialog.remove();
            }
        });
        
        // Input changes
        dialog.addEventListener('input', (e) => {
            const target = e.target as HTMLInputElement | HTMLSelectElement;
            const property = target.getAttribute('data-property');
            
            if (property) {
                if (target.type === 'checkbox') {
                    const checkbox = target as HTMLInputElement;
                    if (property === 'thousandsSeparator') {
                        if (!currentFormat.numberFormat) currentFormat.numberFormat = { type: 'number' };
                        currentFormat.numberFormat.thousandsSeparator = checkbox.checked;
                    }
                } else {
                    const value = target.value;
                    if (property === 'fontSize') {
                        currentFormat.fontSize = parseInt(value) || 14;
                    } else if (property === 'numberType') {
                        if (!currentFormat.numberFormat) currentFormat.numberFormat = { type: 'general' };
                        currentFormat.numberFormat.type = value as any;
                    } else if (property === 'decimalPlaces') {
                        if (!currentFormat.numberFormat) currentFormat.numberFormat = { type: 'number' };
                        currentFormat.numberFormat.decimalPlaces = parseInt(value) || 2;
                    } else if (property === 'currencySymbol') {
                        if (!currentFormat.numberFormat) currentFormat.numberFormat = { type: 'currency' };
                        currentFormat.numberFormat.currencySymbol = value;
                    } else {
                        (currentFormat as any)[property] = value;
                    }
                }
            }
        });
    }

    private highlightSelectedPreset(dialog: HTMLElement, selectedBtn: HTMLElement): void {
        dialog.querySelectorAll('.sheet-format-preset').forEach(btn => btn.classList.remove('selected'));
        selectedBtn.classList.add('selected');
    }

    private applyCellFormat(cellInfo: any, format: CellFormat): void {
        if (this.dataManager) {
            const sheet = this.sheetData[cellInfo.sheetIndex];
            this.dataManager.formatCell(sheet.id, cellInfo.row, cellInfo.col, format);
            this.render();
            this.showToast('Format applied');
        }
    }

    private toggleCellFormat(cellInfo: any, formatType: 'bold' | 'italic'): void {
        if (!this.dataManager) return;
        
        const sheet = this.sheetData[cellInfo.sheetIndex];
        const currentFormat = this.dataManager.getCellFormat(sheet.id, cellInfo.row, cellInfo.col) || {};
        
        let newFormat: CellFormat = {};
        
        if (formatType === 'bold') {
            newFormat.fontWeight = currentFormat.fontWeight === 'bold' ? 'normal' : 'bold';
        } else if (formatType === 'italic') {
            newFormat.fontStyle = currentFormat.fontStyle === 'italic' ? 'normal' : 'italic';
        }
        
        this.dataManager.formatCell(sheet.id, cellInfo.row, cellInfo.col, newFormat);
        this.render();
        this.showToast(`${formatType} ${newFormat.fontWeight === 'bold' || newFormat.fontStyle === 'italic' ? 'applied' : 'removed'}`);
    }

    private clearCellFormat(cellInfo: any): void {
        if (this.dataManager) {
            const sheet = this.sheetData[cellInfo.sheetIndex];
            this.dataManager.clearCellFormat(sheet.id, cellInfo.row, cellInfo.col);
            this.render();
            this.showToast('Format cleared');
        }
    }

    private copyCellFormat(cellInfo: any): void {
        if (this.dataManager) {
            const sheet = this.sheetData[cellInfo.sheetIndex];
            this.copiedFormat = this.dataManager.getCellFormat(sheet.id, cellInfo.row, cellInfo.col);
            this.showToast('Format copied');
        }
    }

    private pasteCellFormat(cellInfo: any): void {
        if (this.dataManager && this.copiedFormat) {
            const sheet = this.sheetData[cellInfo.sheetIndex];
            this.dataManager.formatCell(sheet.id, cellInfo.row, cellInfo.col, this.copiedFormat);
            this.render();
            this.showToast('Format pasted');
        } else {
            this.showToast('No format to paste');
        }
    }

    private showLoadingState(message: string): void {
        // Remove existing loading indicator
        const existing = document.querySelector('.sheet-loading-overlay');
        if (existing) existing.remove();

        const overlay = document.createElement('div');
        overlay.className = 'sheet-loading-overlay';
        const content = overlay.createDiv({cls: 'sheet-loading-content'});
        content.createDiv({cls: 'sheet-loading-spinner'});
        const textSpan = content.createSpan({cls: 'sheet-loading-text'});
        textSpan.textContent = message;
        
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
            
            const content = dialog.createDiv({cls: 'sheet-save-dialog-content'});
            const h3 = content.createEl('h3');
            h3.textContent = 'Save Changes';
            
            const p1 = content.createEl('p');
            p1.textContent = `You have ${unsavedChanges} unsaved change${unsavedChanges !== 1 ? 's' : ''}.`;
            
            const p2 = content.createEl('p');
            p2.textContent = 'Do you want to save these changes?';
            
            const buttonsDiv = content.createDiv({cls: 'sheet-save-dialog-buttons'});
            const saveBtn = buttonsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-primary', attr: {'data-action': 'save'}});
            saveBtn.textContent = 'Save';
            const cancelBtn = buttonsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-action': 'cancel'}});
            cancelBtn.textContent = 'Cancel';

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
            const saveBtnElement = dialog.querySelector('[data-action="save"]') as HTMLElement;
            saveBtnElement?.focus();
        });
    }

    private showFallbackForUnsupportedOperation(operation: string, value?: string, error?: any): void {
        // Keep warning for unsupported operations as it affects functionality
        console.warn(`Unsupported operation: ${operation}`);
        
        // Create a simple fallback dialog
        const fallbackDialog = document.createElement('div');
        fallbackDialog.className = 'sheet-fallback-dialog';
        
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        
        const content = fallbackDialog.createDiv({cls: 'sheet-fallback-dialog-content'});
        const h3 = content.createEl('h3');
        h3.textContent = '⚠️ Operation Not Supported';
        
        const p1 = content.createEl('p');
        p1.textContent = `The operation "${operation}" could not be completed.`;
        
        if (value) {
            const p2 = content.createEl('p');
            const strong1 = p2.createEl('strong');
            strong1.textContent = 'Value:';
            p2.appendText(` ${value}`);
        }
        
        if (error) {
            const p3 = content.createEl('p');
            const strong2 = p3.createEl('strong');
            strong2.textContent = 'Error:';
            p3.appendText(` ${errorMessage}`);
        }
        
        const suggestionsDiv = content.createDiv({cls: 'sheet-fallback-suggestions'});
        const h4 = suggestionsDiv.createEl('h4');
        h4.textContent = 'Try these alternatives:';
        
        const ul = suggestionsDiv.createEl('ul');
        const li1 = ul.createEl('li');
        li1.textContent = 'Simplify the formula or value';
        const li2 = ul.createEl('li');
        li2.textContent = 'Check for syntax errors';
        const li3 = ul.createEl('li');
        li3.textContent = 'Use a different format';
        const li4 = ul.createEl('li');
        li4.textContent = 'Save as plain text';
        
        const buttonsDiv = content.createDiv({cls: 'sheet-fallback-buttons'});
        const copyBtn = buttonsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-action': 'copy-error'}});
        copyBtn.textContent = 'Copy Error';
        const closeBtn = buttonsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-primary', attr: {'data-action': 'close'}});
        closeBtn.textContent = 'Close';

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
            if (this.dataManager && stats) {
                // Update file modification tracking - would need public methods in DataManager
                // this.dataManager.updateFileModified(new Date(stats.mtime));
                // this.dataManager.updateFileChecksum(this.lastFileContent);
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
            const fileModifiedTime = stats ? new Date(stats.mtime) : new Date();

            // Check for conflicts
            const conflict = this.dataManager.detectConflicts(currentContent, fileModifiedTime);
            
            if (conflict) {
                // Keep conflict warnings as they affect data integrity
                console.warn('Conflict detected');
                this.handleConflictDetected(conflict);
            }

        } catch (error) {
            // Keep error logging for conflict checking failures
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
        
        const content = dialog.createDiv({cls: 'sheet-conflict-dialog-content'});
        const h3 = content.createEl('h3');
        h3.textContent = '🔄 Resolve Conflicts';
        
        const p1 = content.createEl('p');
        p1.textContent = `${conflict.conflictedCells.length} cell${conflict.conflictedCells.length !== 1 ? 's have' : ' has'} conflicting changes.`;
        
        const summaryDiv = content.createDiv({cls: 'sheet-conflict-summary'});
        const typeP = summaryDiv.createEl('p');
        const typeStrong = typeP.createEl('strong');
        typeStrong.textContent = 'Conflict Type:';
        typeP.appendText(` ${this.getConflictTypeDescription(conflict.type)}`);
        
        const timeP = summaryDiv.createEl('p');
        const timeStrong = timeP.createEl('strong');
        timeStrong.textContent = 'Detected:';
        timeP.appendText(` ${conflict.timestamp.toLocaleString()}`);
        
        const cellsDiv = content.createDiv({cls: 'sheet-conflict-cells'});
        this.renderConflictedCellsDOM(cellsDiv, conflict);
        
        const strategiesDiv = content.createDiv({cls: 'sheet-conflict-strategies'});
        const h4 = strategiesDiv.createEl('h4');
        h4.textContent = 'Resolution Strategy:';
        
        const strategyButtons = strategiesDiv.createDiv({cls: 'sheet-strategy-buttons'});
        const keepLocalBtn = strategyButtons.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-primary', attr: {'data-strategy': 'keep_local'}});
        keepLocalBtn.textContent = 'Keep My Changes';
        const keepRemoteBtn = strategyButtons.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-strategy': 'keep_remote'}});
        keepRemoteBtn.textContent = 'Accept Remote Changes';
        const mergeBtn = strategyButtons.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-strategy': 'merge'}});
        mergeBtn.textContent = 'Smart Merge';
        const manualBtn = strategyButtons.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-strategy': 'manual'}});
        manualBtn.textContent = 'Resolve Manually';
        
        const actionsDiv = content.createDiv({cls: 'sheet-conflict-actions'});
        const backupBtn = actionsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-action': 'backup'}});
        backupBtn.textContent = 'Create Backup';
        const cancelBtn = actionsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-action': 'cancel'}});
        cancelBtn.textContent = 'Cancel';

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

    private renderConflictedCellsDOM(cellsDiv: HTMLElement, conflict: ConflictInfo): void {
        conflict.conflictedCells.slice(0, 5).forEach(cell => {
            const localValue = this.getCellValueDisplay(cell.localValue);
            const remoteValue = this.getCellValueDisplay(cell.remoteValue);
            const cellRef = this.getCellReference(cell.row, cell.col);
            
            const cellDiv = cellsDiv.createDiv({cls: 'sheet-conflict-cell'});
            const strong = cellDiv.createEl('strong');
            strong.textContent = `${cellRef}:`;
            
            const valuesDiv = cellDiv.createDiv({cls: 'sheet-conflict-values'});
            
            const localDiv = valuesDiv.createDiv({cls: 'sheet-conflict-value local'});
            const localLabel = localDiv.createEl('label');
            localLabel.textContent = 'Your change:';
            const localSpan = localDiv.createEl('span');
            localSpan.textContent = `"${localValue}"`;
            
            const remoteDiv = valuesDiv.createDiv({cls: 'sheet-conflict-value remote'});
            const remoteLabel = remoteDiv.createEl('label');
            remoteLabel.textContent = 'Remote change:';
            const remoteSpan = remoteDiv.createEl('span');
            remoteSpan.textContent = `"${remoteValue}"`;
        });
        
        if (conflict.conflictedCells.length > 5) {
            const p = cellsDiv.createEl('p');
            p.textContent = `...and ${conflict.conflictedCells.length - 5} more`;
        }
    }
    

    private getCellValueDisplay(cellValue: any): string {
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
                const localStr = this.getCellValueDisplay(cell.localValue);
                const remoteStr = this.getCellValueDisplay(cell.remoteValue);
                
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
            
            debugLog('Backup created:', backupFileName);
        } catch (error) {
            // Keep error logging for backup failures
            console.error('Failed to create backup:', error);
            throw error;
        }
    }

    // File size management
    private checkFileSize(data: string): void {
        const fileSizeBytes = new TextEncoder().encode(data).length;
        const deviceType = this.detectDeviceType();
        const threshold = this.fileSizeThresholds[deviceType];
        
        // Check file size
        const isTooBig = fileSizeBytes > threshold;
        
        // Check cell count
        const totalCells = this.sheetData.reduce((count, sheet) => {
            return count + sheet.celldata.length;
        }, 0);
        const hasTooManyCells = totalCells > this.fileSizeThresholds.maxCells;
        
        this.isLargeFile = isTooBig || hasTooManyCells;
        
        if (this.isLargeFile) {
            // Keep warning for large files as it affects performance
        console.warn(`Large file detected: ${this.formatFileSize(fileSizeBytes)}, ${totalCells} cells`);
        }
    }

    private detectDeviceType(): 'mobile' | 'tablet' | 'desktop' {
        const userAgent = navigator.userAgent.toLowerCase();
        const screenWidth = window.screen.width;
        
        if (/android|iphone|ipod/.test(userAgent) || screenWidth < 768) {
            return 'mobile';
        } else if (/ipad|tablet/.test(userAgent) || (screenWidth >= 768 && screenWidth < 1024)) {
            return 'tablet';
        } else {
            return 'desktop';
        }
    }

    private showFileSizeDialog(): void {
        // Remove existing dialog
        const existing = document.querySelector('.sheet-filesize-dialog');
        if (existing) existing.remove();

        const dialog = document.createElement('div');
        dialog.className = 'sheet-filesize-dialog';
        
        const fileSizeBytes = new TextEncoder().encode(this.data).length;
        const totalCells = this.sheetData.reduce((count, sheet) => count + sheet.celldata.length, 0);
        const deviceType = this.detectDeviceType();
        const threshold = this.fileSizeThresholds[deviceType];
        
        const content = dialog.createDiv({cls: 'sheet-filesize-dialog-content'});
        const h3 = content.createEl('h3');
        h3.textContent = '📊 File Size Information';
        
        const statsDiv = content.createDiv({cls: 'sheet-filesize-stats'});
        const sizeP = statsDiv.createEl('p');
        const sizeStrong = sizeP.createEl('strong');
        sizeStrong.textContent = 'File Size:';
        sizeP.appendText(` ${this.formatFileSize(fileSizeBytes)}`);
        
        const cellP = statsDiv.createEl('p');
        const cellStrong = cellP.createEl('strong');
        cellStrong.textContent = 'Cell Count:';
        cellP.appendText(` ${totalCells.toLocaleString()}`);
        
        const deviceP = statsDiv.createEl('p');
        const deviceStrong = deviceP.createEl('strong');
        deviceStrong.textContent = 'Device Type:';
        deviceP.appendText(` ${deviceType}`);
        
        const limitP = statsDiv.createEl('p');
        const limitStrong = limitP.createEl('strong');
        limitStrong.textContent = 'Size Limit:';
        limitP.appendText(` ${this.formatFileSize(threshold)}`);
        
        const warningDiv = content.createDiv({cls: 'sheet-filesize-warning'});
        const warningP1 = warningDiv.createEl('p');
        warningP1.appendText('⚠️ ');
        const warningStrong = warningP1.createEl('strong');
        warningStrong.textContent = 'Performance Warning';
        
        const warningP2 = warningDiv.createEl('p');
        warningP2.textContent = 'This file exceeds the recommended size for optimal performance on your device.';
        
        const ul = warningDiv.createEl('ul');
        const li1 = ul.createEl('li');
        li1.textContent = 'Editing has been disabled to prevent browser freezing';
        const li2 = ul.createEl('li');
        li2.textContent = 'You can still view and scroll through the data';
        const li3 = ul.createEl('li');
        li3.textContent = 'Consider splitting large files into smaller sheets';
        
        const actionsDiv = content.createDiv({cls: 'sheet-filesize-actions'});
        const enableBtn = actionsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-secondary', attr: {'data-action': 'force-enable'}});
        enableBtn.textContent = '🔓 Enable Editing (Risky)';
        const closeBtn = actionsDiv.createEl('button', {cls: 'sheet-dialog-btn sheet-dialog-btn-primary', attr: {'data-action': 'close'}});
        closeBtn.textContent = 'Close';

        const handleClick = (e: Event) => {
            const target = e.target as HTMLElement;
            const action = target.getAttribute('data-action');
            
            if (action === 'force-enable') {
                this.forceEditingEnabled = true;
                this.showToast('⚠️ Editing enabled - performance may be affected');
                this.render(); // Re-render to enable editing controls
                dialog.remove();
            } else if (action === 'close') {
                dialog.remove();
            }
        };

        dialog.addEventListener('click', handleClick);
        document.body.appendChild(dialog);
    }

    private formatFileSize(bytes: number): string {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' KB';
        return Math.round(bytes / (1024 * 1024)) + ' MB';
    }

    private renderFormattingToolbar(toolbar: HTMLElement): void {
        const formatSeparator = toolbar.createEl('div', { cls: 'sheet-toolbar-separator' });
        
        // Bold button
        const boldBtn = toolbar.createEl('button', {
            text: 'B',
            cls: 'sheet-format-btn sheet-format-bold',
            title: 'Bold (Ctrl+B)'
        });
        boldBtn.addEventListener('click', () => this.applyQuickFormat('bold'));
        
        // Italic button
        const italicBtn = toolbar.createEl('button', {
            text: 'I',
            cls: 'sheet-format-btn sheet-format-italic',
            title: 'Italic (Ctrl+I)'
        });
        italicBtn.addEventListener('click', () => this.applyQuickFormat('italic'));
        
        // Underline button
        const underlineBtn = toolbar.createEl('button', {
            text: 'U',
            cls: 'sheet-format-btn sheet-format-underline',
            title: 'Underline (Ctrl+U)'
        });
        underlineBtn.addEventListener('click', () => this.applyQuickFormat('underline'));
        
        // Alignment buttons
        const alignLeftBtn = toolbar.createEl('button', {
            text: '⬅',
            cls: 'sheet-format-btn sheet-format-align-left',
            title: 'Align Left'
        });
        alignLeftBtn.addEventListener('click', () => this.applyQuickFormat('align-left'));
        
        const alignCenterBtn = toolbar.createEl('button', {
            text: '⬌',
            cls: 'sheet-format-btn sheet-format-align-center',
            title: 'Align Center'
        });
        alignCenterBtn.addEventListener('click', () => this.applyQuickFormat('align-center'));
        
        const alignRightBtn = toolbar.createEl('button', {
            text: '➡',
            cls: 'sheet-format-btn sheet-format-align-right',
            title: 'Align Right'
        });
        alignRightBtn.addEventListener('click', () => this.applyQuickFormat('align-right'));
        
        // Currency format
        const currencyBtn = toolbar.createEl('button', {
            text: '$',
            cls: 'sheet-format-btn sheet-format-currency',
            title: 'Currency Format'
        });
        currencyBtn.addEventListener('click', () => this.applyQuickFormat('currency'));
        
        // Percentage format
        const percentBtn = toolbar.createEl('button', {
            text: '%',
            cls: 'sheet-format-btn sheet-format-percent',
            title: 'Percentage Format'
        });
        percentBtn.addEventListener('click', () => this.applyQuickFormat('percentage'));
        
        // Format dialog button
        const formatBtn = toolbar.createEl('button', {
            text: '🎨',
            cls: 'sheet-format-btn sheet-format-dialog',
            title: 'More Formatting Options'
        });
        formatBtn.addEventListener('click', () => {
            if (this.selectedCell) {
                this.showFormattingDialog({
                    sheetIndex: this.selectedCell.sheetIndex,
                    row: this.selectedCell.row,
                    col: this.selectedCell.col,
                    value: this.getCellDisplayValue(this.selectedCell)
                });
            } else {
                this.showToast('Select a cell to format');
            }
        });
    }

    private applyQuickFormat(formatType: string): void {
        if (!this.selectedCell || !this.dataManager) {
            this.showToast('Select a cell to format');
            return;
        }
        
        const sheet = this.sheetData[this.selectedCell.sheetIndex];
        const currentFormat = this.dataManager.getCellFormat(sheet.id, this.selectedCell.row, this.selectedCell.col) || {};
        let newFormat: CellFormat = {};
        
        switch (formatType) {
            case 'bold':
                newFormat = FormatManager.createBoldFormat();
                if (currentFormat.fontWeight === 'bold') {
                    newFormat.fontWeight = 'normal';
                }
                break;
            case 'italic':
                newFormat = FormatManager.createItalicFormat();
                if (currentFormat.fontStyle === 'italic') {
                    newFormat.fontStyle = 'normal';
                }
                break;
            case 'underline':
                newFormat = FormatManager.createUnderlineFormat();
                if (currentFormat.textDecoration === 'underline') {
                    newFormat.textDecoration = 'none';
                }
                break;
            case 'align-left':
                newFormat = FormatManager.createAlignmentFormat('left');
                break;
            case 'align-center':
                newFormat = FormatManager.createAlignmentFormat('center');
                break;
            case 'align-right':
                newFormat = FormatManager.createAlignmentFormat('right');
                break;
            case 'currency':
                newFormat = FormatManager.createCurrencyFormat();
                break;
            case 'percentage':
                newFormat = FormatManager.createPercentageFormat();
                break;
        }
        
        this.dataManager.formatCell(sheet.id, this.selectedCell.row, this.selectedCell.col, newFormat);
        this.render();
        this.showToast(`${formatType} format applied`);
    }

    private getCellDisplayValue(cellInfo: { sheetIndex: number; row: number; col: number }): string {
        const sheet = this.sheetData[cellInfo.sheetIndex];
        const cell = SheetParser.getCellAt(sheet, cellInfo.row, cellInfo.col);
        return cell ? SheetParser.getFormattedCellValue(cell.v) : '';
    }
}