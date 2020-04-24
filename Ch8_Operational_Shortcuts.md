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
