import { Plugin, TFile, WorkspaceLeaf, ViewState } from 'obsidian';
import { SheetView, VIEW_TYPE_SHEET } from './sheetView';

export default class MobileSheetViewer extends Plugin {
    async onload() {
        console.log('Mobile Sheet Viewer: Starting to load plugin');
        
        this.registerView(
            VIEW_TYPE_SHEET,
            (leaf) => {
                console.log('Mobile Sheet Viewer: Creating new SheetView');
                return new SheetView(leaf);
            }
        );

        this.registerExtensions(['sheet'], VIEW_TYPE_SHEET);
        console.log('Mobile Sheet Viewer: Registered extension .sheet with view type', VIEW_TYPE_SHEET);
        
        console.log('Mobile Sheet Viewer: Plugin loaded successfully');
    }

    async onunload() {
        console.log('Mobile Sheet Viewer unloaded');
    }
}