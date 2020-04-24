# Chapter 1: Seven Significant Shortcuts

## Section 1: Enter data or formulas in nonadjacent cells simultaneously
* Ctrl+Enter = Enter data in active (typing) cell into all additionally selected cells
* Alt+'=' = autosum hotkey

## Section 2: Enter data or formulas in nonadjacent cells simultaneously
* Double click the smart copy square (bottom right) to copy a value or formula all the way down a column
  * This does not work with rows sadly
* Ctrl+'.' =
  * Jump to the end of a single column
  * Jump to next clockwise corner

## Section 3: Instantly enter today's date or time
* Ctrl+';' = Enter in today's date into the cell
* Ctrl+Shift+';' = Enter the current time into the cell
* TODAY() = function for selecting the current day
* NOW() = Function to display the current day and time

## Section 4: Convert formulas to values with a simple drag
* PROPER() = Function to capitalize only the first letter and letters following punctuation and spaces
* TRIM() = Remove whitespace
* Remove formulas and leave only values
  1. Select the area
  2. Right click and hold
  3. Drag away from the current position of the data set to the target
     1. This can be the exact same position so long as it is first moved away
  4. Let go of right click
  5. Select "Copy Here as Values Only"

## Section 5: Display all worksheet formulas
* Ctrl+'`' = Double width of all columns and show formulas instead of results if one is present
* Select all formula cells
  * Go to Home tab
  * To the far right, use the "Find & Select" tool
  * Select the "Formulas" option
  * It is suggested to use the fill tool to change the colors of the formula colors

## Section 6: Create charts from keystroke shortcuts
* Alt+F1 = Create default chart (clustered bar chart) over selected data
* F11 = Create chart on separate sheet

## Section 7: Zoom in and out quickly
* Mouse and scroll wheel can zoom in and out
* Ctrl+Alt+'+' = Zoom in
* Ctrl+Alt+'-' = Zoom out

# Chapter 2: Ribbon and Quick Access Toolbar Tips

## Section 1: Access ribbon commands with Alt key sequences
* Pressing the Alt Key shows the accessibility menu. This also can be used as a speed boost
* Alt+'H'+'K' = Give selected cells the comma (number) format.
* Alt+'A'+'SA' = Sort Ascending
* Alt+'A'+'SA' = Sort Descending

## Section 2: Expand the ribbon and full-screen views
* Ctrl+F1 = Minimize Ribbon to view more rows
  * Tabs still unfurl breifly if you click them
* Ctrl+Shift+F1 = Fully Minimize Ribbon to view EVEN MORE rows
  * Click the very top bar to briefly show the entire ribbon

# Chapter 3: Display Shortcuts

## Section 1: Create split screens and frozen tables in a flash
* View / Freeze Top Row = keeps top row in place
  * Row must be visible
  * Can be combined with Column
* View / Freeze Top Column = keeps top column in place
  * Column must be visible
  * Can be combined with Row
* Split
  * Highlighting the cell against the farthest left column will create a horizonital split with identical sheets in the top and bottom of the viewing pane
  * Highlighting a cell in the top row will create a veritcal split with identical sheets in the left and right of the viewing pane
  * Highlighting a cell in the middle of the sheet will create a quad-split allowing view of the same sheet 4 times in 4 quadrants

## Section 2: Restore missing labels and hide repeating labels
* Looking at data set with sparse catagories set in one column, you can cascade this down by using the "Find & Select" Tool
  * Use Go To Special
  * Blanks
  * Hit Ok
  * Highlighted should be all blanks in a column
  * The top cell allows you to enter a formula
  * Enter =A2 and hit enter (or top row)
  * This effect should cascade down the sparse column
* Now we can make a conditional format rule
  * Set the rule to use formula
  * Enter =A2=A1 where if the row is equal to the row above, then we do something
  * Set it to format the text to white
  * Now the data is still there even though we can't see it or print it

## Section 3: Customize the display of status
* Look toward the bottom of the screen to the left of the zoom slider
  * You will see some stats on selected cells but not all
  * Right click, and select options you want
    * Recommend checking all from Avg to Sum
* With this you can see stats without writing or generating formulas

# Chapter 4: Navigation and Selection Shortcuts

## Section 1: Navigate between workbooks and worksheets efficiently
* Bottom left, right click to view all sheets in a list
* Ctrl+PgUp = go to next sheet (Number pad 9)
* Ctrl+PgDown = go to previous sheet (Number pad 3)
* Ctrl+'N' = Create new workbook (.xlsx file)
* Switching between them
  * View / Switch Windows
  * Show all open sheets
* Ctrl+Tab = switch between last viewed workbook
* Ctrl+Shift+Tab = switch between next open workbook
* Any button can be added to quick access toolbar

## Section 2: Navigate within worksheets efficienly
* Ctrl+End = jump to lower right portion of the sheet (Number pad 1)
* Ctrl+Home = jump to upper left of sheet (Number pad 7)
  * Freezing rows/cols ignore the default home and go to next non-frozen cell furtherest up and left
* Double click edge of cell (any of the four) = zip to furthest non-black cell in the direction of the click
* Ctrl+Arrow-Keys = zip to furthest non-black cell in the arrow key direction
* Shift+Double-click edge of cell (any of the four) = zip to furthest non-black cell in the direction of the click
* Ctrl+Shift+Arrow-Keys = zip to furthest non-black cell in the arrow key direction
* F5 = type in the cell to go straight to that cell

## Section 3: Open, close, save, and create new workbooks with keystrokes
* Ctrl+O / Ctrl+F12 = Open file
* Ctrl+W / Ctrl+F4 = Close file
* Ctrl+S / Shift+F12 = Save file
* F12 = Save As
* Ctrl+N = New file
* Alt+F4 = Close Excel and all workbooks

## Section 4: Select and entire row, column, region, or worksheet
* Ctrl+'A' / Ctrl+'*' / Shift+Ctrl+'8' / Shift+Ctrl+Space = Select contiguous region
  * 8 is the asterisk key if num pad is not present/desired
* Ctrl+'A' then 'A' = Select entire sheet
  * If a table is present, hit A again to get entire sheet
* Shift+Space = highlight all cells in a row
* Ctrl+Space = highlight all cells in a column

## Section 5: Select noncontiguous ranges, and visible cells only
* Alt+';' = Select visible cells only
  * Alternative GUI option is to use "Find & Select" / Go To Special / Visible Cells only
  * Hidden Columns and Rows will not show up

# Chapter 5: Data Entry and Editing Shortcuts

## Section 1: Accelerate data entry
* Alt+Enter = go to newline
* Excel will automatically suggest options when one option has been used before
* Right click and Select "Pick from Drop-down list" will allow you to pick from previously used data in a column
  * Only usable with text
* Alt+Down-Arrow = Open drop-down list
* Select an area 

## Section 2: Enter dates and date series efficiently
* Excel will auto fill months when it senses abbreviations
  * This does not apply to "Sept", only 3 letters
* Excel can also do this with full spellings
  * May counts as full spellings of months
* You can do this right to forwards in time or left to go backwards
* Same with down which goes forwards in time and up which goes backwards
* All caps spellings work
* Days of the week work
* In order to stop it from auto incrementing, hold Ctrl
* Remember that Ctrl+';' puts the current day in the selected cell
* Right click and select the area to extend into. This will open up a dialog to Copy days, weekdays, months, etc.
* Illegal dates will float right

## Section 3: Enter times and time series efficiently
* Typing times with colons (7:00, 5:45) will create a Data formatted cell
* Type 'a' or 'p' will force it to take the hour
  * '5:00 a' = 05:00:00 AM
  * '5:00 p' = 17:00:00 PM
* Whole hours can be shortened
  * '5 a' = 05:00:00 AM
  * '5 p' = 17:00:00 PM
* Whole hours can be shortened
* Behind the scenes it stores fractions of days
  * Formatting a decimal to Time will make approximate it's time of day
* Highlight a set of 2 times 17 mins apart
  * Smart copy with add 17 mins all the way down the column
* Ctrl+'1' = Format Cells
* Ctrl+Shift+2 = Set Format to Time

## Section 4: Use Custom lists for data entry and list-based sorting
* Setup a custom order of values for easy sorting
  * Highlight data for list (Optional)
  * Go to File
  * Options
  * From the left hit Advanced
  * Scroll nearly all the way down to General
  * The last row of this section, hit "Edit Custom Lists..."
  * This opens up all available list
  * If the data is highlighted Click Import
  * If not you can:
    1. Use the cell select tool near the bottom
    2. Type the values in the "List Entries" Box
  * Hit Ok
* Now click in the column you want to sort
* Select the Data Tab
* Click the bigger "Sort" button
* Make sure the column and Sort on are fine
* For Order, click the drop down and hit "Custom List..."
* Selct your newly made custom list and hit Ok
* List is sortered customly!

## Section 5: Enhance editing tools
* Double-click a word or select a section of text
* This brings up the mini tool bar for quick text mods
* Home Key goes to the begining of a line
* End Key goes to the end of a line
* Ctrl+Home-Key goes to the begining of a cell
* Ctrl+End-Key goes to the end of a cell

# Chapter 6 Drag and Drop Techniques

## Section 1: Accelerate copy and move tasks within cells and worksheets
* Using LMB, click the border and drag the data
  * Formulas come with it and use the new location of data
  * Hold Alt to change sheets by dragging over the sheet in the sheets bar at the bottom
  * Hold Ctrl to make a Copy
* Copy Sheet    
  * Right click sheet and hit Make Copy
  * Hold Ctrl Key and drag sheet to the left of the original sheet
* Make new workbook out of data

## Section 2: Drag and insert cells with the Shift Key
* Select column then Shift+LMB-and-Hold-edge = Move column
* Select column then Shift+RMB-and-Hold-edge = Move column with options
* Select group of cels then Shift+RMB-and-Hold-edge = Move column group

## Section 3: Display the Paste Special options
* Ctrl+Alt+V = Paste Special
  * Transpose Paste
* Also Findable through the Home tab, Click the Past drop down then "Paste Special"
  * Transpose is also preset if you'd like that more

## Section 4: Accelerate copy, move paste actions with the right mouse button
* Using the RMB on the lower right square and dragging to fit an area brings up the smart-fill options menu
  * Select "Fill Formatting Only" to Paste the formatting of the first cell into the rest of the copied tiem
* Using the RMB on any edge of a selection and dragging it elsewhere brings up the move-special options menu.
  * Select "Fill Formatting Only" to Paste the formatting of the pre-selected area to the newly desired one
* You can also right click and drag the selected area to another area and use "Shift right and copy"

# Chapter 7: Formula Shortcuts

## Section 1: Create formulas rapidly
* Alt+'=' = Autosum hotkey
  * Then press the column (ex 'J' or 'A') that will sum all cells in the selected column
* sumif() = Sums all rows X column if the row also has Y value equal to something
  * EX: `=SUMIF(E:E,M13,J:J)`
  * FORM: `=SUMIF(range,criteria,[sum_range])`
    * for range if == critera place in sum
* averageif() = Average all rows X column if the row also has Y value equal to something
  * EX: `=AVERAGEIF(E:E,M13,J:J)`
  * FORM: `=AVERAGEIF(range,criteria,[average_range])`
    * for range if == critera place in sum. Finally average that over all rows that == critera
* Sum can also do multiple rows
  * `=SUM(3:3,13:13)` gets both row 3 and row 13

## Section 2: Select all dependent or precendent cells
* Ctrl+Shift+']' = Select all cells dependent on the currently selected cell
* Trace Dependents = Draw arrows to show the flow of cell that is being used in other cells
  * Formulas Tab, Formula Auditing group
  * Clicking multiple times will show each level down the tree it goes
  * It will also show different sheets (Dashed line to icon of sheet)
* Ctrl+Shift+'[' = Select all cells precendent on the currently selected cell
* Trace Precendent = Draw arrows to show the flow of cell that is being used in selected cell
  * Formulas Tab, Formula Auditing group
  * Clicking multiple times will show each level up the tree it goes

## Section 3: Use AutoSum shortcuts
* Alt+'=' = Autosum hotkey
* Use commas to spearate different ranges/sections
  * Need not be contiguous
* Excel will favor going up and left if the AutoSum is used without highlighting a specific area
* Highlighing a 2+ by 2+ area will do multiple sums at once
  * This favors going down
  * Highlight the area right and down to do both at the same time
  * Also works with average
    * Selectable from the Autosum dropdown

## Section 4: Count the number of unique entries
* Highlight Duplicates can be flipped into Unique
* Filter can look at only unique records
* Array Formula = press Ctrl+Shift+Enter
  * `{=SUM(1/COUNTIF(C3:C743,C3:C743))}`

## Section 5: Use conditional formatting to highlight formula cells
* Find and select only highlights stuff in the moment, not actively adding or removing when a requirement is met or not met
* Create a highlight rule using formula
  * Home tab / Conditional Formatting / New Rule / Use Formula to ... 
  * `=ISFORMULA(A1)`
  * Using A1 is done in the background, but applies to the whole sheet

## Section 6: Preform calculations without formulas
* Paste special can add cell to a group of cells
  1. Select Cell to Add
  2. Select Area to add X to all of group Y
  3. Ctrl+Alt+'V' = Paste Special
  4. Choose 'Add' in the operations
* Can also be used with multiply
  * Ex. Use `1.1` as the cell and choose multiply

# Chapter 8: Operational Shortcuts

## Section 1: Insert, delete, hide, and unhide columns and rows
* Home tab / Cells Group
* Insert Column to the left of another
  * Right click a column, Insert. Inserts a column to the left
* Similar with a row above
* On a group you select to insert
  * Helps not alter a multi table sheet
  * Select a group, RMB, Insert
    * Shift cells down to insert row
    * Shift cells right to insert column
* Delete does the opposite
  * Select cells, RMB, Shift cells up
* Hide also stops printing from seeing hidden cells
* Unhide
  * RMB, Unhide
  * Double LMB
  * Multi-unhide, select all the columns or rows to unhide, RMB, unhide
* Sometimes the column is acutally just too thin and unhide will not work
* If the width is adjust to `<= 0` then the column will be auto hidden

## Section 2: Realign imported text
* Home Tab / Editing Group / Fill / Justify
* Select a group of cells to realign the text into
  * Big long string of text, select 3 (example) cells, Justify will try and width cap it to the width of the 3 cells
  * Afterwards selecting then cells (now multiple) to the right further (blanks of the right) will let justify recalculate the alignment again
* LIMIT OF 255 CHARACTERS
  * Then it removes spaces, does include spaces in the 255 sum

## Section 3: Select and manipulate blank cells
* Home Tab / Editting Group / Find & Select / Go To Special / Blanks
  * Selects only blanks cells
  * If an area is selected it will only select in that area
* You can use formulas to conditionally do operations based on if a cells is blank of not
  * You can do the simple: `=IF(H2="",$M$1,$N$1)*J2`
  * Or the more complicated, but easier to read: `=IF(ISBLANK(H2),$M$1,$N$1)*J2`

## Section 4: Collapse and expand detail with ungroup buttons
* Feature called "Outline"
  * Expand and collapse data
* Data tab / Outline group / Group button / Down arrow / Auto Outline
  * Pops the screen into a more easily controllable set of data
* Ctrl+'8' = Hide Outline Frame
* To clear it go to Ungroup / Clear Outline

## Section 5: Create double-spaced or triple-spaced printouts
* Remember you can hide and unhide rows and columns
* Bottom right near the zoom level bar, click the "Page Break Preview"
  * Slide the right most blue bar to the left to stop print additional rightward tables
  * The dashed lines are assumed breaks
* Quick way of double spacing is to highlight the whole sheet and change the height of a row. This changes all rows to be the height of the changed row.
  * Ctrl+'Z' = Remove the double spacing

# Chapter 9: Formatting Shortcuts

## Section 1: Use keystroke for frequency needed numeric formats
* Ctrl+Shift+'`' = General
* Ctrl+Shift+'1' = Number
  * Commas and Decimal
* Ctrl+Shift+'2' = Time
  * Colon and AM/PM
* Ctrl+Shift+'3' = Date
  * Day month year
* Ctrl+Shift+'4' = Currency
  * Dolloar sign
* Ctrl+Shift+'5' = Percent
  * Percent sign
* Ctrl+Shift+'6' = Scientific
  * `'_.__E+__'` Format

## Section 2: Accentuate data with alignment tools
* Home tab / Alignment group
* Orientation allows you to change the rotation of text
  * Apply borders and the borders will follow the angle as well
* Vertical text looks bad
  * Rotate text up is better
* RMB / Format Cells / Alignment = Custom angled texts
* Merge & Center = Combine Cells together
  * Only works Vertical or Horizontal
  * Crtl Select sets of hori or vert only groups to multi-merge
* Vertical merge and vertical text looks good

## Section 3: Add a color background to every fifth row in a range
* Ctrl+'T' = Apply Table format
* Go to Home tab / Styles group / New Rule / Use Formula ...
  * `=MOD(ROW(),5)=0` = Every 5th row is formatted
  * Stays even when cells change

## Section 4: Use conditional formatting based on comparison criteria
* Go to Home tab / Styles group / New Rule / Use Formula ...
  * cols C and B both have dats in MM/DD/YYYY format
  * `=C1>B1+2` will apply formatting (ex: fill orange) if the C col date is 3 or more days later than B col date for the same row
* For comparing ranks Use
  * `=K1<J1` for when the rank has improved
  * `=K1>J1` for when the rank has dropped
  * `=K1=J1` for when the rank has stayed the same

## Section 5: Use special formats for times over 24 hours
* Ctrl+'1' = Format Cells
* You can format a cell to Time
  * Ctrl+Shift+2
* Formatting an AutoSum as Time will inherently Mod it by 24 thus removing the sum nature of the cell
  * Ctrl+1 to Format the cells / time / The 7th row down is the summed time format
  * To remove seconds go to Custom / Type and remove the ":ss;@"

## Section 6: Add and remove strikethrough borders with keyboard shortcuts
* Ctrl+Shift+7 = Quick apply Border around active cells
* Ctrl+Shift+'-' = Remove all borders on active cells
* Ctrl+5 = Apply or Remove strikethrough

## Section 7: Display values in thousands or millions without formulas
* Under custom Formatting use a trailing comma (`#,##0.0,`) to display values in thousands
* Use 2 for millions, `#,##0.0,,`
* Example used was for U.S. State populations

## Section 8: Format phone and social security numbers
* Formal cells with Ctrl+1 / Special / SSN for social security numbers
  * Exact Format: `000-00-0000`
* Formal cells with Ctrl+1 / Special / Phone number for phone numbers
  * Exact Format: `[<=9999999]###-####;(###) ###-####`
* Formatted Cells are NOT equivalent to the text form
  * 776007286 != 776-00-7286
* Find & Select / Replace can take out these '-', '(', and/or ')'

# Chapter 10: Data Management Tech.md

## Section 1: Clean up spaces with the TRIM Function
* We have leading and trailing spaces/whitespace
* F9 to quick evaluate
* `TRIM()` removes leading or trailing whitespace
  * Ex 1) 
    1. Insert an extra temp column (a new D column), set the first cell in the  column to be `=TRIM(C1)`
    2. Double click the smart box at the bottom right of the active cell
    3. Now RMB and hold, drag over the original column, from the drop down select Copy Here as Values Only
    4. Delete the new temp column
  * Ex 2)
    1. We have a formula `=IF(AND(E2="Full Time",J2>3),1000,0)`
    2. But the E column has trailing whitespace sometimes so the strict comparison of `"Full Time" == "Full Time "` is FALSE
    3. For this we would use TRIM`=IF(AND(TRIM(E2)="Full Time",J2>3),1000,0)`

## Section 2: Split column data using Text to Columns and Flash Fill
* Data tab / Data Tools group / Text to columns
  * Use space delimeted
  * Splits the columns over the space sperateing names
* Data tab / Data Tools group / Flash Fill
  * Flash Fill is a quick way to fill a nearby column with data
  * Excel will assume you are spliting based on common delimiters
* Ctrl-E to use quick Flash Fill

## Section 3: Join data with the TEXTJOIN function and Flash Fill
* Flash Fill can do this for you
  * Type the "last name, first name." Then press enter.
* Or `=C2 & ", " & B2`
* To remove the green triangles in the box go to Fill / Options / Formulas uncheck "Numbers formmatted as text..."
  * Also useful to remove this and more for presentation of a data sheets
* TEXTJOIN
  * Only available in Office 365 which I do not have
  * `=TEXTJOIN("=",TRUE,J2:M2)`
  * `Concatenate(X,"=",Y,...)` will have to do
    * Ex) `=CONCATENATE(J2,"-",K2,"-",L2,"-",M2)`

## Section 4: Insure unique entires with data validation rules
* Hightlight cell rules
  * Does not stop the user from entering poor data such as duplicates
* For this we need data validation
  * Data tab / Data tools group / Data Validation
  * Custom rule
  * `=AND(COUNTIF(B:B,B1)=1,LEN(B1)=4)`
    * Needs to be unique
    * Needs to be a length of strictly 4

## Section 5: Display unique items from large lists
* You can use Remove Duplicates from the Data tab to drop rows that are EXACT duplicates
  * Data tab / Data Tools tab / Remove Duplicates
* To make a copy of the list use the Advanced Filter
  * Data tab / Sort & Filter tab / Advanced (Filter)
  * Toggle to Copy to another location
  * Select the list range
  * Choose a Copy to location
  * Select unique records only
* To see just the unique categories in a list do the same as above but simply select only a single column
* You can also create a Pivot table off of a single column. This makes a new sheet with the pivot table in it. On the right select the column you selected before

# Chapter 11: Charting and Visual Object Tips

## Section 1: Manipulate chart placement and sizing with dragging techniques
* Alt+F1 = Quick chart
* Chart Tools header / Design tab (contextual) / Chart Styles tab / select a different overall style
* Hold Shift while dragging a cornder handle
  * Keeps aspect ratio the same
* Hold Ctrl while dragging a side handle to make an equal change on the opposite 
* Hold Alt on chart move to snap to a cell corner in the top left
* Hold Alt on a handle drag (corner/side) to snap the adjust to a cell
  * Any cell height/width adjusts will adjust the chat too
* Also in the Design tab you'll see the Switch Row/Column tab for showing flipping the axis interpretation of the chart
  * Duplicate the chart if you'd like to see the same chart at the same time

## Section 2: Create chart titles from cell content
* Create chart using Alt+F1
* To add a cell based title:
  1. Click the title
  2. Press '='
  3. Click the cell you want to use as the chart title
  4. Enter
* Note this does not work on older/stranger chart
  * Ex) Treemap charts

## Section 3: Create and manipulate shapes with Shift, Ctrl, and Alt keys
* Hold Shift while dragging a cornder handle
  * Keeps aspect ratio the same
* The Orange circle in some shapes allow to control contextual features
  * Angles on a Hexagon
  * Radius of a Sun
  * Smile or Frown of a Smiley-Face
* Hold Alt to snap to cell
* Use the send forward and backward options to move the order the images appear

## Section 4: Create linked dynamic and linked static images
* Using the Paste dropdown on the Home tab
  * Use "Paste Picture" near the bottom for a static image paste that does NOT update when you change the data (directly or by formula)
  * Use "Paste Linked Picture" near the bottom for a static image paste that does update when you change the data (directly or by formula)
* Use the normal Fill in the Home tab to give the picture a background color
* Useful for monitoring other sheets as formulas change in the active sheet

# Chapter 12: PivotTable Tips

## Section 1: Adjust report layouts to allow totals on top or bottom
* Create a Pivot table in the Insert tab
* Tabular Form does not allow report totals to go to the top or bottom

## Section 2: Revisit relevant source data with the drill down feature
* Create a Pivot table in the Insert tab
* Double click a cell or RMB and select "Show detail" to build a new PT on the specific cell

# Chapter 13: Tiny Tips

## Section 1: Become a more proficient Excel user with these short tips
1. Ctrl+Enter to keep the active cell in place
   1. Keep writing in the current cell
   2. Allows for quick modification
2. Don't capitalize function names
   1. Does it automatically
3. Don't use the right parenthesis
   1. Just press enter
4. Don't collapse data
   1. Move the window
   2. It will collapse when selecting (click, drag, release)
5. Don't LMB when you RMB right after
   1. Does not work for drags
   2. RMB will select it already
6. Shift+F10 will activate the shortcut menu (RMB)
7. RMB on selected cell group, it gives a mini tool bar
   1. Avoid clicking the Home tab
8. Ctrl+Shift+F1 = Go to home screen
   1. Ctrl+F1 = Hide the Ribbon
9. Format entire column not just data
   1.  Format keeps going for any added rows
   2.  Simple modification in the future
10. You don't have to highlight cells to create a chart if surrounded by empty cells

# Chapter 14 Conclusion

## Section 1: Next Steps
* Wrap up
* Thanks for watching
