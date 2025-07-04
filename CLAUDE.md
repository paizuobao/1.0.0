# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a web-based intelligent seating arrangement system (智能座位安排系统) designed for teachers to manage classroom seating. It's a client-side only application built with vanilla HTML, CSS, and JavaScript.

## Architecture

### Core Components

- **SeatingApp Class** (`app.js`): Single main class that manages the entire application state and functionality
- **HTML Structure** (`index.html`): Three-column layout with student management (left), seating visualization (center), and controls (right)  
- **CSS Styling** (`style.css`): Responsive design with mobile support and modern UI components

### Key Data Structures

- `students[]`: Array of student objects with properties like name, height, gender, vision needs
- `seats[]`: Grid-based array representing classroom layout, each seat can hold a student reference
- `history[]`: Undo/redo system for seat arrangement changes
- `constraints[]`: Rules for intelligent seating algorithms

### Data Persistence

All data is stored in browser LocalStorage under two keys:
- `seatingData`: Current students, seats, layout configuration
- `savedLayouts`: Multiple named seating arrangements for backup/restore

## Running the Application

**Development/Testing:**
```bash
# Simply open in any modern web browser
open index.html
# Or serve via local server for CORS compliance if needed
python -m http.server 8000
```

**Production Deployment:**
- Live site: https://paizuobao.github.io/1.0.0/
- GitHub repository: https://github.com/paizuobao/1.0.0.git
- Deployed via GitHub Pages automatically from main branch

**No build process required** - this is a vanilla JavaScript application.

## Deployment Instructions

### GitHub Pages Setup
1. Push code to GitHub repository: `https://github.com/paizuobao/1.0.0.git`
2. Go to repository Settings > Pages
3. Select "Deploy from a branch" and choose "main" branch
4. Select "/ (root)" as the folder
5. Save - the site will be available at `https://paizuobao.github.io/1.0.0/`

### Deployment Requirements
- All file paths are relative (✓ already configured)
- External libraries loaded via CDN (✓ SheetJS and html2canvas)
- No server-side processing required (✓ client-side only)
- Static hosting compatible (✓ GitHub Pages ready)

### Updating the Live Site
1. Make changes to local files
2. Commit and push to main branch:
   ```bash
   git add .
   git commit -m "Update: [description of changes]"
   git push origin main
   ```
3. GitHub Pages will automatically rebuild and deploy (usually takes 1-2 minutes)

## Key Functionality Areas

### Student Management
- CRUD operations for student records
- Drag-and-drop from student list to seats
- Search/filter capabilities
- Student attributes affect seating algorithms

### Seating Algorithms  
- `randomSeatArrangement()`: Core algorithm with vision-impaired priority (front 1/3 of rows)
- `assignStudentToSeat()`: Manual assignment with displacement logic
- Front row detection: `seat.row < Math.ceil(this.rows / 3)` identifies priority seating
- Constraint system for future algorithm extensions

### Layout Management
- Dynamic grid generation based on rows/cols configuration
- Canvas-based export to PNG images
- Multiple layout save/restore functionality

### UI State Management
- History system with undo/redo (Command pattern)
- Modal dialogs for forms
- Real-time statistics and occupancy tracking

## Extending the System

### Adding New Seating Algorithms
Implement new methods in SeatingApp class following the pattern:
1. Add history checkpoint with `addToHistory()`
2. Clear existing assignments 
3. Apply algorithm logic
4. Call `saveData()`, `renderClassroom()`, `renderStudentList()`, `updateStats()`

### Adding Student Attributes
1. Update student form in HTML modal
2. Modify `saveStudent()` method to capture new fields
3. Update student display in `renderStudentList()`
4. Incorporate into seating algorithms as needed

### Layout Customization
The CSS uses CSS Grid for the classroom layout. Key implementation details:

**Grid Generation:**
- Dynamic grid template: `grid-template-columns: repeat(${this.cols}, 1fr)`
- Dynamic grid template: `grid-template-rows: repeat(${this.rows}, 1fr)`
- Uniform gap spacing applied to all seats via CSS Grid `gap` property

**Responsive Design:**
- Desktop: 0.1rem gap, 65px × 45px seats
- Tablet (≤768px): 0.08rem gap, 50px × 35px seats  
- Mobile (≤480px): 0.05rem gap, 45px × 30px seats

**Seat Rendering Logic:**
- Each seat element is positioned automatically by CSS Grid
- Seat numbers calculated as: `${this.rows - seat.row}-${seat.col + 1}` (converts internal to display coordinates)
- Empty seats show position number, occupied seats show student name with gender-based colors
- Coordinate system annotations added throughout codebase for clarity

### Excel Import Feature
The system supports importing student lists from Excel files:

**Supported Excel Formats:**
- `.xlsx` and `.xls` files
- First row must contain headers
- Required column: "姓名" (Name) - various naming patterns supported
- Optional columns: 学号 (ID), 性别 (Gender), 身高 (Height), 视力 (Vision), 备注 (Notes)

**Column Recognition:**
- Name columns: 姓名, 名字, name
- ID columns: 学号, 编号, id, number  
- Gender columns: 性别, gender
- Height columns: 身高, height
- Vision columns: 视力, 近视, 眼镜, vision
- Notes columns: 备注, 说明, note, remark

**Excel Import Workflow:**
1. `importExcelFile()` - Triggers file selection dialog
2. `parseExcelData()` - Reads and processes Excel file using SheetJS
3. `validateExcelData()` - Validates data and identifies errors
4. `showExcelPreviewModal()` - Shows preview with valid/error data tabs
5. `confirmExcelImport()` - Imports data with duplicate handling options

**Template Download:**
`downloadExcelTemplate()` generates a sample Excel file with proper column headers and example data.

## Recent Architecture Changes

### UI Layout Evolution
The system uses a two-column layout:
- **Left Panel**: Student management with search/filter capabilities and Excel import
- **Center Panel**: Classroom seating visualization with optimized spacing and coordinate system

### Coordinate System and Seat Layout
**Internal vs Display Coordinate Systems:**
- **Internal Coordinates (0-based)**: `row: 0` = first row (near podium), `col: 0` = leftmost column
- **Display Coordinates (1-based)**: Row 1 = front row (near podium), Column 1 = leftmost from teacher perspective
- **Conversion Formula**: `${this.rows - seat.row}-${seat.col + 1}` transforms internal to display coordinates
- **Classroom Perspective**: First row first column (1-1) is front-left, matching traditional classroom layout

**Physical Layout:**
- **Position Display**: Seat numbers appear in top-left corner of each seat
- **Grid Spacing**: Proportional layout with percentage-based padding (4% 4% 0.5% 4%) and row-gap (0.3%)
- **Seat Dimensions**: 100px width × 50px height (desktop), responsive for mobile

### Settings Modal System
Seating configuration is handled through a translucent overlay modal triggered by the "排座设置" button, containing:
- Classroom layout controls (rows/columns)
- Seating rules configuration
- Constraint management with enhanced input placeholder

### Toolbar Integration
The main toolbar includes an undo button positioned between "Random Seat" and "Clear Seats" for immediate access to history functionality.

### Modal System Architecture
- **Student Modal**: Add/edit student information with form validation
- **Excel Preview Modal**: Large modal for importing and previewing Excel data
- **Settings Modal**: Translucent overlay for seating configuration
- All modals use event delegation for dynamic content and proper cleanup

## Dependencies

**External Libraries:**
- SheetJS (XLSX) v0.18.5: Excel file processing (loaded via CDN)
- html2canvas v1.4.1: Screenshot and image export functionality (loaded via CDN)

**Browser Requirements:**
- Modern browser with ES6+ support
- LocalStorage API
- HTML5 drag-and-drop API
- Canvas API (for image export)

## Critical Implementation Details

### Coordinate System Implementation
**Key Principle**: Internal coordinates are 0-based, display coordinates are 1-based classroom perspective
- All internal logic uses 0-based coordinates where `row: 0` = front row, `col: 0` = leftmost
- All UI display uses 1-based coordinates converted via `${this.rows - seat.row}-${seat.col + 1}`
- Coordinate conversion is annotated throughout codebase for maintainability
- Front row detection: `seat.row < Math.ceil(this.rows / 3)` for vision-impaired priority

### History System
The undo system uses a Command pattern with:
- `addToHistory()`: Records state before operations
- `undo()`: Restores previous state
- History is persisted to LocalStorage and restored on page load
- Only seat arrangement operations are tracked (no student CRUD in history)

### Event Handling Patterns
- Direct event binding for main UI elements
- Event delegation for dynamically created content (student list items, modal buttons)
- Careful handling of modal button events to prevent binding issues

### Drag-and-Drop System
**Student Assignment:**
- Students can be dragged from the left panel student list to any seat
- Drag events use `dataTransfer.setData()` with student UUID
- Drop targets handle seat assignment via `assignStudentToSeat()`
- Displacement logic: if target seat is occupied, displaced student moves to dragged student's original seat

**Visual Feedback:**
- Dragging items get `.dragging` class (opacity: 0.5, rotation)
- Drop targets get `.drag-over` class during hover
- Seat removal buttons (×) appear on occupied seats

### State Management Flow
Critical update sequence for any seating change:
1. `addToHistory()` - Record current state
2. Perform operation
3. `saveData()` - Persist to LocalStorage  
4. `renderClassroom()` - Update seat visualization
5. `renderStudentList()` - Update student list
6. `updateStats()` - Update statistics
7. `updateHistoryButtons()` - Update undo button state
8. `applyCurrentFilter()` - Maintain filter state

### Visual Features and Styling

**Gender-Based Color Coding:**
- Male students: Blue text color (#2563eb)
- Female students: Pink text color (#ec4899)
- Applied via CSS classes `.male` and `.female` on seat elements

**Dynamic Font Sizing:**
- Font size automatically adjusts based on name length
- Classes: `.name-4`, `.name-5`, `.name-6`, `.name-7`, `.name-8-plus`
- Prevents text overflow in seats with longer names

**Coordinate Display Toggle:**
- Checkbox in classroom header controls seat number visibility
- Setting persisted in localStorage as `showCoordinates`
- Applied via `.hide-coordinates` class on classroom grid

**Podium Design:**
- Rectangular shape with rounded corners (8px border-radius)
- Brown gradient background with drop shadow effects
- Positioned below seat grid using CSS Grid

### Image Export System

**Export Functionality:**
- Primary method: html2canvas for high-quality screenshots with DOM cloning
- Fallback method: Pure canvas drawing with manual coordinate-based rendering
- Export configuration: complete classroom layout matching web display exactly

**Export Features:**
- Exports complete classroom layout (seats + podium + background) matching web display
- Maintains gender color coding in exported images (blue for boys, pink for girls)
- Preserves all visual styling including gradients, borders, and 3D podium effects
- Seat numbers preserved using display coordinate system
- Dynamic font sizing based on student name length
- Automatic fallback to canvas method if html2canvas fails
- Both methods produce identical visual output

**Image Export Workflow:**
1. `exportLayout()` - Main export function using html2canvas with complete styling
2. Clone DOM elements with exact web styling in `onclone` callback  
3. Apply all visual elements including classroom background, seat styling, and podium
4. Fallback to `exportLayoutCanvas()` if html2canvas fails
5. Canvas fallback draws complete layout with `drawRoundedRect()` and `drawPodium()` helpers
6. Generate PNG download with timestamp filename