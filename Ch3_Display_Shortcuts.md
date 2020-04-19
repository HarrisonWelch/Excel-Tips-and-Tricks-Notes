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

