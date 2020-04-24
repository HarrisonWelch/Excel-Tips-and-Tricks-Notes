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
