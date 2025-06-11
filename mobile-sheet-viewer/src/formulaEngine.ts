export interface CellReference {
    sheet?: string;
    row: number;
    col: number;
}

export interface CellRange {
    start: CellReference;
    end: CellReference;
}

export interface FormulaResult {
    value: number | string | boolean;
    error?: string;
    dependencies: CellReference[];
}

export interface FormulaContext {
    getCellValue: (ref: CellReference) => any;
    getRangeValues: (range: CellRange) => any[];
    currentCell: CellReference;
}

export class FormulaEngine {
    private static readonly FUNCTIONS = new Map<string, Function>([
        ['SUM', FormulaEngine.sum],
        ['COUNT', FormulaEngine.count],
        ['AVERAGE', FormulaEngine.average],
        ['MIN', FormulaEngine.min],
        ['MAX', FormulaEngine.max],
        ['IF', FormulaEngine.if],
        ['CONCATENATE', FormulaEngine.concatenate],
        ['LEN', FormulaEngine.len],
        ['UPPER', FormulaEngine.upper],
        ['LOWER', FormulaEngine.lower],
        ['ROUND', FormulaEngine.round],
        ['ABS', FormulaEngine.abs],
        ['NOW', FormulaEngine.now],
        ['TODAY', FormulaEngine.today]
    ]);

    static parseFormula(formula: string, context: FormulaContext): FormulaResult {
        try {
            if (!formula.startsWith('=')) {
                return { value: formula, dependencies: [] };
            }

            const expression = formula.substring(1); // Remove '='
            const dependencies: CellReference[] = [];
            const result = this.evaluateExpression(expression, context, dependencies);

            return {
                value: result,
                dependencies
            };
        } catch (error) {
            return {
                value: '#ERROR!',
                error: error instanceof Error ? error.message : 'Unknown error',
                dependencies: []
            };
        }
    }

    private static evaluateExpression(expr: string, context: FormulaContext, dependencies: CellReference[]): any {
        // Remove whitespace
        expr = expr.trim();

        // Handle cell references (A1, B2, etc.)
        const cellRefMatch = expr.match(/^([A-Z]+)([0-9]+)$/);
        if (cellRefMatch) {
            const ref = this.parseCellReference(expr);
            dependencies.push(ref);
            return context.getCellValue(ref);
        }

        // Handle numbers
        if (/^-?\d*\.?\d+$/.test(expr)) {
            return parseFloat(expr);
        }

        // Handle strings
        if ((expr.startsWith('"') && expr.endsWith('"')) || (expr.startsWith("'") && expr.endsWith("'"))) {
            return expr.slice(1, -1);
        }

        // Handle function calls
        const funcMatch = expr.match(/^([A-Z_]+)\((.*)\)$/);
        if (funcMatch) {
            const [, funcName, argsStr] = funcMatch;
            const func = this.FUNCTIONS.get(funcName);
            
            if (!func) {
                throw new Error(`Unknown function: ${funcName}`);
            }

            const args = this.parseArguments(argsStr, context, dependencies);
            return func.apply(null, args);
        }

        // Handle basic arithmetic
        return this.evaluateArithmetic(expr, context, dependencies);
    }

    private static parseArguments(argsStr: string, context: FormulaContext, dependencies: CellReference[]): any[] {
        if (!argsStr.trim()) return [];

        const args: any[] = [];
        let current = '';
        let depth = 0;
        let inString = false;
        let stringChar = '';

        for (let i = 0; i < argsStr.length; i++) {
            const char = argsStr[i];

            if (!inString && (char === '"' || char === "'")) {
                inString = true;
                stringChar = char;
                current += char;
            } else if (inString && char === stringChar) {
                inString = false;
                current += char;
            } else if (!inString && char === '(') {
                depth++;
                current += char;
            } else if (!inString && char === ')') {
                depth--;
                current += char;
            } else if (!inString && char === ',' && depth === 0) {
                args.push(this.evaluateExpression(current.trim(), context, dependencies));
                current = '';
            } else {
                current += char;
            }
        }

        if (current.trim()) {
            args.push(this.evaluateExpression(current.trim(), context, dependencies));
        }

        return args;
    }

    private static evaluateArithmetic(expr: string, context: FormulaContext, dependencies: CellReference[]): number {
        // Replace cell references with values
        expr = expr.replace(/([A-Z]+[0-9]+)/g, (match) => {
            const ref = this.parseCellReference(match);
            dependencies.push(ref);
            const value = context.getCellValue(ref);
            return String(this.toNumber(value));
        });

        // Safe arithmetic evaluation without eval
        return this.safeEvaluateArithmetic(expr);
    }

    private static safeEvaluateArithmetic(expr: string): number {
        // Remove whitespace
        expr = expr.replace(/\s/g, '');
        
        // Basic expression parser for +, -, *, /
        try {
            // Handle parentheses first (recursive)
            while (expr.includes('(')) {
                const start = expr.lastIndexOf('(');
                const end = expr.indexOf(')', start);
                if (end === -1) throw new Error('Mismatched parentheses');
                
                const innerExpr = expr.substring(start + 1, end);
                const result = this.safeEvaluateArithmetic(innerExpr);
                expr = expr.substring(0, start) + result + expr.substring(end + 1);
            }

            // Handle multiplication and division (left to right)
            expr = this.processOperations(expr, ['*', '/']);
            
            // Handle addition and subtraction (left to right)
            expr = this.processOperations(expr, ['+', '-']);
            
            const result = parseFloat(expr);
            if (isNaN(result)) throw new Error('Invalid expression');
            
            return result;
        } catch (error) {
            throw new Error('Invalid arithmetic expression');
        }
    }

    private static processOperations(expr: string, operators: string[]): string {
        for (const op of operators) {
            let index = 0;
            while ((index = expr.indexOf(op, index)) !== -1) {
                // Skip if it's a negative sign at the beginning or after another operator
                if (op === '-' && (index === 0 || '+-*/('.includes(expr[index - 1]))) {
                    index++;
                    continue;
                }

                // Find left operand
                let leftStart = index - 1;
                while (leftStart >= 0 && /[\d.]/.test(expr[leftStart])) {
                    leftStart--;
                }
                leftStart++;

                // Find right operand
                let rightEnd = index + 1;
                if (expr[rightEnd] === '-') rightEnd++; // Handle negative numbers
                while (rightEnd < expr.length && /[\d.]/.test(expr[rightEnd])) {
                    rightEnd++;
                }

                const left = parseFloat(expr.substring(leftStart, index));
                const right = parseFloat(expr.substring(index + 1, rightEnd));

                let result: number;
                switch (op) {
                    case '+': result = left + right; break;
                    case '-': result = left - right; break;
                    case '*': result = left * right; break;
                    case '/': 
                        if (right === 0) throw new Error('Division by zero');
                        result = left / right; 
                        break;
                    default: throw new Error('Unknown operator');
                }

                expr = expr.substring(0, leftStart) + result + expr.substring(rightEnd);
                index = leftStart + String(result).length;
            }
        }
        return expr;
    }

    static parseCellReference(ref: string): CellReference {
        const match = ref.match(/^([A-Z]+)([0-9]+)$/);
        if (!match) {
            throw new Error(`Invalid cell reference: ${ref}`);
        }

        const [, colStr, rowStr] = match;
        return {
            row: parseInt(rowStr) - 1, // Convert to 0-based
            col: this.columnToIndex(colStr)
        };
    }

    static parseCellRange(rangeStr: string): CellRange {
        const parts = rangeStr.split(':');
        if (parts.length !== 2) {
            throw new Error(`Invalid range: ${rangeStr}`);
        }

        return {
            start: this.parseCellReference(parts[0]),
            end: this.parseCellReference(parts[1])
        };
    }

    private static columnToIndex(col: string): number {
        let result = 0;
        for (let i = 0; i < col.length; i++) {
            result = result * 26 + (col.charCodeAt(i) - 65 + 1);
        }
        return result - 1; // Convert to 0-based
    }

    static indexToColumn(index: number): string {
        let result = '';
        index++; // Convert to 1-based
        while (index > 0) {
            index--;
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26);
        }
        return result;
    }

    private static toNumber(value: any): number {
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            const num = parseFloat(value);
            return isNaN(num) ? 0 : num;
        }
        return 0;
    }

    private static toArray(value: any): any[] {
        return Array.isArray(value) ? value : [value];
    }

    // Built-in functions
    private static sum(...args: any[]): number {
        let total = 0;
        for (const arg of args) {
            if (Array.isArray(arg)) {
                total += arg.reduce((sum, val) => sum + FormulaEngine.toNumber(val), 0);
            } else {
                total += FormulaEngine.toNumber(arg);
            }
        }
        return total;
    }

    private static count(...args: any[]): number {
        let count = 0;
        for (const arg of args) {
            if (Array.isArray(arg)) {
                count += arg.filter(val => val !== null && val !== undefined && val !== '').length;
            } else if (arg !== null && arg !== undefined && arg !== '') {
                count++;
            }
        }
        return count;
    }

    private static average(...args: any[]): number {
        const values: number[] = [];
        for (const arg of args) {
            if (Array.isArray(arg)) {
                values.push(...arg.map(val => FormulaEngine.toNumber(val)));
            } else {
                values.push(FormulaEngine.toNumber(arg));
            }
        }
        return values.length > 0 ? values.reduce((sum, val) => sum + val, 0) / values.length : 0;
    }

    private static min(...args: any[]): number {
        const values: number[] = [];
        for (const arg of args) {
            if (Array.isArray(arg)) {
                values.push(...arg.map(val => FormulaEngine.toNumber(val)));
            } else {
                values.push(FormulaEngine.toNumber(arg));
            }
        }
        return values.length > 0 ? Math.min(...values) : 0;
    }

    private static max(...args: any[]): number {
        const values: number[] = [];
        for (const arg of args) {
            if (Array.isArray(arg)) {
                values.push(...arg.map(val => FormulaEngine.toNumber(val)));
            } else {
                values.push(FormulaEngine.toNumber(arg));
            }
        }
        return values.length > 0 ? Math.max(...values) : 0;
    }

    private static if(condition: any, trueValue: any, falseValue: any): any {
        const cond = FormulaEngine.toNumber(condition);
        return cond !== 0 ? trueValue : falseValue;
    }

    private static concatenate(...args: any[]): string {
        return args.map(arg => String(arg)).join('');
    }

    private static len(text: any): number {
        return String(text).length;
    }

    private static upper(text: any): string {
        return String(text).toUpperCase();
    }

    private static lower(text: any): string {
        return String(text).toLowerCase();
    }

    private static round(number: any, digits: number = 0): number {
        const num = FormulaEngine.toNumber(number);
        const factor = Math.pow(10, digits);
        return Math.round(num * factor) / factor;
    }

    private static abs(number: any): number {
        return Math.abs(FormulaEngine.toNumber(number));
    }

    private static now(): number {
        return Date.now();
    }

    private static today(): string {
        return new Date().toISOString().split('T')[0];
    }
}