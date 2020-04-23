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