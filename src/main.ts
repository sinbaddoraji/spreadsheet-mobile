import { Plugin, TFile, WorkspaceLeaf, ViewState } from 'obsidian';
import { SheetView, VIEW_TYPE_SHEET } from './sheetView';

// Debug configuration
const DEBUG = false; // Set to true during development

function debugLog(message: string, data?: any) {
    if (DEBUG) {
        if (data) {
            console.debug(`[Mobile Sheet Viewer] ${message}`, data);
        } else {
            console.debug(`[Mobile Sheet Viewer] ${message}`);
        }
    }
}

export default class MobileSheetViewer extends Plugin {
    async onload() {
        debugLog('Starting to load plugin');
        
        this.registerView(
            VIEW_TYPE_SHEET,
            (leaf) => {
                debugLog('Creating new SheetView');
                return new SheetView(leaf);
            }
        );

        this.registerExtensions(['sheet'], VIEW_TYPE_SHEET);
        debugLog('Registered extension .sheet with view type', VIEW_TYPE_SHEET);
        
        debugLog('Plugin loaded successfully');
    }

    async onunload() {
        debugLog('Plugin unloaded');
    }
}