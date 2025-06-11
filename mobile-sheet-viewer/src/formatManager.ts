import { CellFormat, NumberFormat, FormatPreset } from './sheetParser';

export class FormatManager {
    private static presets: FormatPreset[] = [
        // Text formatting presets
        {
            id: 'bold',
            name: 'Bold',
            description: 'Bold text',
            category: 'text',
            format: { fontWeight: 'bold' }
        },
        {
            id: 'italic',
            name: 'Italic',
            description: 'Italic text',
            category: 'text',
            format: { fontStyle: 'italic' }
        },
        {
            id: 'underline',
            name: 'Underline',
            description: 'Underlined text',
            category: 'text',
            format: { textDecoration: 'underline' }
        },
        {
            id: 'center',
            name: 'Center',
            description: 'Center aligned text',
            category: 'text',
            format: { textAlign: 'center' }
        },
        
        // Number formatting presets
        {
            id: 'currency',
            name: 'Currency',
            description: 'Currency format ($1,234.56)',
            category: 'number',
            format: {
                numberFormat: {
                    type: 'currency',
                    currencySymbol: '$',
                    decimalPlaces: 2,
                    thousandsSeparator: true
                }
            }
        },
        {
            id: 'percentage',
            name: 'Percentage',
            description: 'Percentage format (12.34%)',
            category: 'number',
            format: {
                numberFormat: {
                    type: 'percentage',
                    decimalPlaces: 2
                }
            }
        },
        {
            id: 'number',
            name: 'Number',
            description: 'Number format with thousands separator',
            category: 'number',
            format: {
                numberFormat: {
                    type: 'number',
                    decimalPlaces: 2,
                    thousandsSeparator: true
                }
            }
        },
        
        // Date formatting presets
        {
            id: 'date_short',
            name: 'Short Date',
            description: 'MM/dd/yyyy',
            category: 'date',
            format: {
                numberFormat: {
                    type: 'date',
                    dateFormat: 'MM/dd/yyyy'
                }
            }
        },
        {
            id: 'date_long',
            name: 'Long Date',
            description: 'EEEE, MMMM dd, yyyy',
            category: 'date',
            format: {
                numberFormat: {
                    type: 'date',
                    dateFormat: 'EEEE, MMMM dd, yyyy'
                }
            }
        },
        {
            id: 'time',
            name: 'Time',
            description: 'HH:mm:ss',
            category: 'date',
            format: {
                numberFormat: {
                    type: 'time',
                    timeFormat: 'HH:mm:ss'
                }
            }
        },
        
        // Highlight presets
        {
            id: 'highlight_yellow',
            name: 'Yellow Highlight',
            description: 'Yellow background',
            category: 'highlight',
            format: { backgroundColor: '#ffeb3b' }
        },
        {
            id: 'highlight_green',
            name: 'Green Highlight',
            description: 'Green background',
            category: 'highlight',
            format: { backgroundColor: '#4caf50', color: '#ffffff' }
        },
        {
            id: 'highlight_red',
            name: 'Red Highlight',
            description: 'Red background',
            category: 'highlight',
            format: { backgroundColor: '#f44336', color: '#ffffff' }
        },
        {
            id: 'header',
            name: 'Header Style',
            description: 'Bold header with gray background',
            category: 'highlight',
            format: {
                fontWeight: 'bold',
                backgroundColor: '#e0e0e0',
                textAlign: 'center',
                border: {
                    all: { width: 1, style: 'solid', color: '#bdbdbd' }
                }
            }
        }
    ];

    static getPresets(): FormatPreset[] {
        return [...this.presets];
    }

    static getPresetsByCategory(category: string): FormatPreset[] {
        return this.presets.filter(preset => preset.category === category);
    }

    static getPreset(id: string): FormatPreset | null {
        return this.presets.find(preset => preset.id === id) || null;
    }

    static addCustomPreset(preset: FormatPreset): void {
        this.presets.push(preset);
    }

    static removePreset(id: string): boolean {
        const index = this.presets.findIndex(preset => preset.id === id);
        if (index !== -1) {
            this.presets.splice(index, 1);
            return true;
        }
        return false;
    }

    // Quick format functions for common operations
    static createBoldFormat(): CellFormat {
        return { fontWeight: 'bold' };
    }

    static createItalicFormat(): CellFormat {
        return { fontStyle: 'italic' };
    }

    static createUnderlineFormat(): CellFormat {
        return { textDecoration: 'underline' };
    }

    static createColorFormat(color: string): CellFormat {
        return { color };
    }

    static createBackgroundFormat(backgroundColor: string): CellFormat {
        return { backgroundColor };
    }

    static createAlignmentFormat(align: 'left' | 'center' | 'right'): CellFormat {
        return { textAlign: align };
    }

    static createFontSizeFormat(fontSize: number): CellFormat {
        return { fontSize };
    }

    static createBorderFormat(width: number = 1, style: 'solid' | 'dashed' | 'dotted' = 'solid', color: string = '#000000'): CellFormat {
        return {
            border: {
                all: { width, style, color }
            }
        };
    }

    static createCurrencyFormat(symbol: string = '$', decimalPlaces: number = 2): CellFormat {
        return {
            numberFormat: {
                type: 'currency',
                currencySymbol: symbol,
                decimalPlaces,
                thousandsSeparator: true
            }
        };
    }

    static createPercentageFormat(decimalPlaces: number = 2): CellFormat {
        return {
            numberFormat: {
                type: 'percentage',
                decimalPlaces
            }
        };
    }

    static createDateFormat(dateFormat: string = 'MM/dd/yyyy'): CellFormat {
        return {
            numberFormat: {
                type: 'date',
                dateFormat
            }
        };
    }

    static createTimeFormat(timeFormat: string = 'HH:mm:ss'): CellFormat {
        return {
            numberFormat: {
                type: 'time',
                timeFormat
            }
        };
    }

    // Utility functions for format detection and suggestions
    static detectFormat(value: string): CellFormat | null {
        // Auto-detect format based on value patterns
        
        // Currency detection
        if (/^\$?\d{1,3}(,\d{3})*(\.\d{2})?$/.test(value.trim())) {
            return this.createCurrencyFormat();
        }
        
        // Percentage detection
        if (/^\d+(\.\d+)?%$/.test(value.trim())) {
            return this.createPercentageFormat();
        }
        
        // Date detection (MM/dd/yyyy, dd/MM/yyyy, yyyy-MM-dd)
        if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value.trim()) || 
            /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
            return this.createDateFormat();
        }
        
        // Time detection (HH:mm:ss, HH:mm)
        if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(value.trim())) {
            return this.createTimeFormat();
        }
        
        // Number with thousands separators
        if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(value.trim())) {
            return {
                numberFormat: {
                    type: 'number',
                    thousandsSeparator: true,
                    decimalPlaces: 2
                }
            };
        }
        
        return null;
    }

    static suggestFormats(value: string): FormatPreset[] {
        const suggestions: FormatPreset[] = [];
        
        // Always suggest basic text formatting
        suggestions.push(
            ...this.getPresetsByCategory('text').slice(0, 4)
        );
        
        // Add specific suggestions based on value
        if (!isNaN(Number(value))) {
            suggestions.push(...this.getPresetsByCategory('number'));
        }
        
        if (this.looksLikeDate(value)) {
            suggestions.push(...this.getPresetsByCategory('date'));
        }
        
        // Add highlight options
        suggestions.push(...this.getPresetsByCategory('highlight').slice(0, 3));
        
        return suggestions;
    }

    private static looksLikeDate(value: string): boolean {
        const trimmed = value.trim();
        return /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed) ||
               /^\d{4}-\d{2}-\d{2}$/.test(trimmed) ||
               /^\d{1,2}-\d{1,2}-\d{4}$/.test(trimmed) ||
               !isNaN(Date.parse(trimmed));
    }

    // Color palette for quick access
    static getColorPalette(): string[] {
        return [
            '#000000', '#424242', '#757575', '#bdbdbd', '#ffffff',
            '#f44336', '#e91e63', '#9c27b0', '#673ab7', '#3f51b5',
            '#2196f3', '#03a9f4', '#00bcd4', '#009688', '#4caf50',
            '#8bc34a', '#cddc39', '#ffeb3b', '#ffc107', '#ff9800',
            '#ff5722', '#795548', '#607d8b'
        ];
    }

    // Mobile-optimized format suggestions
    static getMobileQuickFormats(): FormatPreset[] {
        return [
            this.getPreset('bold')!,
            this.getPreset('italic')!,
            this.getPreset('center')!,
            this.getPreset('currency')!,
            this.getPreset('percentage')!,
            this.getPreset('highlight_yellow')!,
            this.getPreset('highlight_green')!,
            this.getPreset('highlight_red')!
        ].filter(Boolean);
    }
}