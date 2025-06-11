# Mobile-First Obsidian Spreadsheet Plugin Plan

## Overview
Create a new Obsidian plugin that provides spreadsheet functionality with a mobile-first approach, ensuring it works seamlessly on iPhone, iPad, and desktop platforms.

## Key Problems with Current Plugin
1. **FortuneSheet is too heavy** - 60KB+ CSS files, large JavaScript bundle
2. **Memory constraints** - iOS WebView has strict memory limits
3. **Complex dependencies** - React + FortuneSheet creates a large dependency tree
4. **No progressive enhancement** - All-or-nothing loading approach

## Proposed Solution

### 1. Lightweight Spreadsheet Library Options
Instead of FortuneSheet, consider these mobile-friendly alternatives:

#### Option A: **Luckysheet-lite** (Custom Build)
- Strip down Luckysheet to core features only
- Remove unused modules (charts, pivot tables, etc.)
- Estimated size: ~100KB total

#### Option B: **Handsontable CE** (Community Edition)
- More mature, better mobile support
- Modular architecture
- Touch-optimized
- Estimated size: ~200KB with minimal features

#### Option C: **Custom Implementation**
- Build a minimal spreadsheet using native HTML tables
- Use ContentEditable for cell editing
- Implement only essential features
- Estimated size: ~50KB total
- Full control over mobile optimizations

#### Option D: **x-spreadsheet**
- Lightweight (~100KB)
- No dependencies
- Touch support built-in
- Canvas-based rendering (efficient)

### 2. Architecture Design

```
obsidian-mobile-sheets/
├── src/
│   ├── main.ts           # Plugin entry point
│   ├── view.ts           # View handler (no React)
│   ├── spreadsheet/
│   │   ├── core.ts       # Core spreadsheet logic
│   │   ├── renderer.ts   # Rendering engine
│   │   ├── mobile.ts     # Mobile-specific code
│   │   └── formulas.ts   # Basic formula support
│   ├── utils/
│   │   ├── storage.ts    # File handling
│   │   └── compat.ts     # .sheet file compatibility
│   └── styles/
│       ├── base.css      # Minimal base styles
│       └── mobile.css    # Mobile optimizations
├── manifest.json
├── package.json
└── esbuild.config.mjs
```

### 3. Feature Prioritization

#### Phase 1: MVP (Mobile-First)
- [ ] Basic grid display (read-only)
- [ ] Cell selection and navigation
- [ ] Text input and editing
- [ ] Save/load .sheet files
- [ ] Mobile touch gestures

#### Phase 2: Essential Features
- [ ] Basic formulas (SUM, AVERAGE, etc.)
- [ ] Copy/paste functionality
- [ ] Row/column resize
- [ ] Cell formatting (bold, italic)
- [ ] Undo/redo

#### Phase 3: Advanced Features
- [ ] Cell merging
- [ ] Sorting and filtering
- [ ] More formula functions
- [ ] Import/export CSV
- [ ] Cell borders and colors

### 4. Technical Approach

#### Mobile-First Development
1. **Start with iPhone constraints**
   - Max 50MB memory usage
   - Touch-first interactions
   - Responsive from 320px width

2. **Progressive Enhancement**
   - Core features work everywhere
   - Enhanced features for desktop
   - Graceful degradation

3. **Performance Optimizations**
   - Virtual scrolling for large sheets
   - Lazy loading of features
   - Debounced saves
   - Minimal DOM manipulation

#### Data Format
Maintain compatibility with existing .sheet files:
```json
[{
  "name": "Sheet1",
  "data": [
    [{"v": "A1"}, {"v": "B1"}],
    [{"v": "A2"}, {"v": "B2"}]
  ],
  "config": {
    "merge": {},
    "rowlen": {},
    "collen": {}
  }
}]
```

### 5. Implementation Timeline

#### Week 1: Foundation
- Set up project structure
- Implement basic grid renderer
- Mobile touch handling
- File loading/saving

#### Week 2: Editing
- Cell editing functionality
- Keyboard navigation
- Basic formatting
- Auto-save implementation

#### Week 3: Formulas
- Formula parser
- Basic functions
- Cell references
- Dependency tracking

#### Week 4: Polish
- Performance optimization
- Bug fixes
- Documentation
- Testing on various devices

### 6. Technology Stack

#### Core Dependencies
- **No React** - Use vanilla TypeScript
- **No heavy frameworks** - Minimal dependencies
- **Build tool**: esbuild (fast, efficient)
- **Styles**: PostCSS with autoprefixer

#### Optional Libraries (evaluate size/benefit)
- `formula-parser`: For formula evaluation (30KB)
- `hammerjs`: For touch gestures (20KB)
- `localforage`: For better mobile storage (25KB)

### 7. Success Criteria
- [ ] Loads in < 1 second on iPhone
- [ ] Uses < 50MB memory with 1000 cells
- [ ] Smooth scrolling at 60fps
- [ ] Works offline
- [ ] Compatible with existing .sheet files
- [ ] No "Failed to load" errors on mobile

### 8. Testing Strategy
- Test on iPhone SE (smallest screen)
- Test on iPad (tablet layout)
- Test with large spreadsheets (10,000+ cells)
- Memory profiling on iOS
- Battery usage testing

### 9. Migration Path
1. New plugin with different ID: `obsidian-mobile-sheets`
2. Import tool for existing .sheet files
3. Gradual feature parity
4. Eventually replace original plugin

### 10. Next Steps
1. Evaluate spreadsheet library options
2. Create minimal prototype
3. Test on iPhone to validate approach
4. Choose final architecture
5. Begin implementation

## Decision Point
Before proceeding, we need to choose:
1. **Library**: x-spreadsheet vs custom implementation
2. **Scope**: Start minimal or include formulas in MVP?
3. **Compatibility**: 100% .sheet compatibility or new format?

The key is to start small, validate on mobile first, then expand features based on what actually works on iPhone.