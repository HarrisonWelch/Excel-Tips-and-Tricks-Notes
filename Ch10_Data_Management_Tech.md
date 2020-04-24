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
