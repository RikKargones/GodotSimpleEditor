# GodotSimpleXlsxEditor
A very simple Excel .xlsx file editor writed in pure GDScript.

Created for task on work - writing data (only values/formulas) in cells. I think it's needs to be public - maybe some unfortunate souls can make use of it)

Code uses a bit changed [GodotXML addon by elenakrittik](https://github.com/elenakrittik/GodotXML) and structurally following [Godot Excel Reader addon by LaoDie1](https://github.com/LaoDie1/godot-excel-reader?ysclid=m6j75bap25275265483).

Most of additinal code is actual edit cell values/formulas and saving xlsx files.

## What SimpleXlsxEditor CAN do:
- **Reading/editing cells inside sheets.**
- **Save edited data into xlsx file (opend original or copy)**

## What SimpleXlsxEditor CAN'T do:
- **Opening and editing xls files.** - Data stracture inside them are very, very different. Maybe sometime I write support for BIFF, but likely not.
- **Creating something bigger then cell** - sheets, styles and file itself.
- **Calculating formula result in cell**

## Differences inside used addons:

### Godot Excel Reader:
- **Really checking file validness** - no crash happening while trying to open xls file, non-xlsx file or corrupted xlsx. Needed data for read/edit are checked on opening.
- **Replaced XML reader code** - original one only reads data from XML files inside xlsx. GodotXML allows edit data.
### GodotXML
- **Seeking and collecting funcs** - as *seek_first_child_by_name()* and *get_childs_by_name()*.
- **XMLDocument keeps XML header for later document save** - (silly one - also trying to know in which UTF-type document should be saved)
- **Fixed: dump_str without *pretty* parameter makes XML file unredable by browser or Excel because of skipped spaces between attributes**.

## How to use:

### Opening/closing

```GDScript
# Creates new ExcelFile object if sucessfully opend file. Otherwise returns null.
var excel_file : ExcelFile = ExcelFile.open_file(file_path)

if excel_file == null: return

# Closing file (without saving):
excel_file.close()

# You can check file validness:
if not ExcelFile.is_file_valid(other_file_path): return

# ExcelFile object itself also can open xlsx. If object have opend xlsx file alredy, closes. Returns OK on sucess.
if not excel_file.open(other_file_path) == OK: return

```

### Cells reading/editing

```GDScript
# Getting sheets array
var sheets_list : Array[ExcelSheet] = excel_file.get_sheets()
# TODO: Make ExcelWorkbook funcs get_sheet_name_list() and get_sheet_by_name() avalible from ExcelFile

if sheets_list.size() == 0: return

var sheet : ExcelSheet = sheets_list[0]

# As in Godot Excel Reader, you can recive cells contents with valid values as dictionaries in dictionaries:
var table : Dictionary = sheet.get_table()

# Table structure are the same: first layer is rows dictionaries, second - cells dictionaries with column keys.
for row in table.keys():
    for column in table[row].keys():
        print(row, ",", column, "-", table[row][column])

# Contents inside can be vary - mostly storing int, float or String value types.

# Also you can get "raw" cells - with exposed formulas.
var table_with_formulas : Dictionary = sheet.get_table(false)

# If you need to know value/formula of a specific cell, you can call this funcs:
var cell_value : Variant = sheet.get_cell_value("A1")
var cell_formula : String = sheet.get_cell_formula("A1")
# Both of them returns empty String if cell is empty/not writed by Excel before.

# There is two funcs for setting cells, both of them returns OK on edit sucess:

# ExcelSheet.set_cell(coord : Vector2i, data : Variant, is_formula : bool = false)
# More robust one - sets one cell at time, reciving Vector2i as coords.
# Columns and rows are x and y here accordingly.
# Also start of table coordinats ("A1") here - Vector2i(1,1).
# Accepts only Strings, int-s and float-s.
# Otherwise returns ERR_INVALID_DATA.
# With "is_formula" set to true, accepts only Strings. Second func acts seemengly.

sheet.set_cell(Vector2i(3,5), 25) # Writes integer number 25 to cell "C5" value
sheet.set_cell(Vector2i(4,6), 67.55) # Writes float number 67.55 to cell "D6" value
sheet.set_cell(Vector2i(5,7), "Hello!") # Writes string "Hello!" to cell "E7" value

# To translate Excel coord to Vector2i, you can use to_coords() func
to_coords("B7") # returns Vector2i(2,7)

# ExcelSheet.set_cells_range(range_str : String, data : Variant, is_formula : bool = false) is more usable func for setting cells.
# Can recive coords as a single cell ("A1") and as an Excel range ("A1:B7").
# Ranges as "B1:A4", "A6:C5" or "F3:A1" will be fixed to normal top-left to right-bottom pairs.
# set_cells_range(), adiitionaly with Strings, int-s and float-s,
# can recive PackedInt32Arrays, PackedInt64Array, PackedFloat32Array, PackedFloat64Array and usual Array.
# (Whoops! Forget to add support for PackedStringsArray...)
# These ones, instead of setting all cells by one value, will fill them element by element until reaching end of array or cells range.
# If usual Array element contains unacceptable by set_cell() value - both value and cell will be skipped.

sheets.set_cells_range("A1:B8", 3) # Sets all cells value to 3 in range "A1:B8"
sheets.set_cells_range("A1:B8", [3, 2.45, True, "Hi!"]) # Result - "A1" - 3, "A2" - 2.45, "A3" - "", "A4" - "Hi!"

# Hint: you can still use ExcelSheet or ExcelWorkbook objects even if xlsx file are closed,
# one differense is what you cannot save edited data.
# ExcelFile will premently load data to ExcelSheets outside of ExcelWorkbook before closing, as if they become edited.
# As any Resource or RefCounted, memory for them will be freed if all varibles, containing them, are set to null.
```

### Saving data

```GDScript
# To save opend file somethere, you need this func. Returns OK on sucess.
excel_file.save_as(path)

# Save to opend file. Seemengly returns OK on sucess.
excel_file.save()
```
