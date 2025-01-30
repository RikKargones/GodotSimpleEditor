# GodotSimpleXlsxEditor
A very simple Excel .xlsx file editor writed in pure GDScript.

Created for task on work - writing data (only values/formulas) in cells. I think it's needs to be public - maybe some unfortunate souls can make use of it)

Code using a bit changed [GodotXML addon by elenakrittik](https://github.com/elenakrittik/GodotXML) and structurally following [Godot Excel Reader addon by LaoDie1](https://github.com/LaoDie1/godot-excel-reader?ysclid=m6j75bap25275265483).

Most of additinal code is actual edit cell values/formulas and saving xlsx files.


## What SimpleXlsxEditor CAN do:
- **Reading/editing cells inside sheets.**
- **Save edited data into xlsx file (opend original or copy)**


## What SimpleXlsxEditor CAN'T do:
- **Opening and editing xls files.** - Data stracture inside them are very, very different. Maybe sometime I write support for BIFF, but likely not.
- **Creating something bigger then cell** - Sheets, styles and file itself.
- **Calculating formula result in cell**


## Differences inside used addons:

### Godot Excel Reader:
- **Really checking file validness** - no crash happening while trying to open xls file, non-xlsx file or corrupted xlsx. Needed data for read/edit are checked on opening.
- **Replaced XML reader code** - original one only reads data from XML files inside xlsx. GodotXML allows edit data.

### GodotXML
- **Seeking and collecting funcs** - as *seek_first_child_by_name()* and *get_childs_by_name()*.
- **XMLDocument keeps XML header for later document save** - (silly one - also trying to know in which UTF-type document should be saved)
- **Fixed: dump_str without *pretty* parameter makes XML file unredable by browser or Excel because of skipped spaces between attributes**.


## Konwn issues:

**ExcelSheet.set_cell()** doesn't know about merged cells.

**ExcelSheet.set_cells_range()** doesn't accepts PackedStringArray's, but technically should.


## How to use:

### Checking file validness:

```GDScript
ExcelFile.is_file_valid(file_path)
```

Returns **true**, if file is valid for ZIPReader and contains these three XML files:
- xl/workbook.xml
- xl/_rels/workbook.xml.rels
- xl/sharedStrings.xml

### Opening xlsx file:

```GDScript
var excel_file : ExcelFile = ExcelFile.open_file(file_path) # Returns null if failed to open file.
```

Alternative:
```GDScript
var excel_file : ExcelFile = ExcelFile.new()

excel_file.open(file_path) # returns OK on sucess
```

### Getting sheet

```GDScript
var sheets_list : Array[ExcelSheet] = excel_file.get_sheets()

if sheets_list.size() == 0: return

var sheet : ExcelSheet = sheets_list[0] # Sucessfuly opend file contains at least ONE sheet.
```

Alternative:
```GDScript
var sheet : ExcelSheet = excel_file._workbook.get_sheet_by_name(needed_name) # returns null if sheet not exist
```
TODO: Make ExcelWorkbook funcs *get_sheet_name_list()* and *get_sheet_by_name()* avalible from ExcelFile

### Getting cells values/formulas
<hr>

**ExcelSheet.get_table(return_formula_results : bool = true) -> Dictionary**

As in Godot Excel Reader, you can recive cells contents with valid values:
```GDScript
var table : Dictionary = sheet.get_table()
```
Table structure are the same: it's a dictionaries with used rows indexes as keys what contains other dictionaries with cell values by column indexes keys.

> [!NOTE]
> Only **used** cells (aka with value/formula) are inside these dictionaries.

Usual use might look like this:
```GDScript
var table : Dictionary = sheet.get_table()

for row in table.keys():
    for column in table[row].keys():
        var cell_value : Variant = table[row][column]
        # Doing something with cell value
```
Contents inside cells can be vary - mostly storing int, float or String value types.

Also you can get "raw" dictionaries - with exposed formulas. Just set flag *return_formula_results* to false:
```GDScript
var table_with_formulas : Dictionary = sheet.get_table(false)
```
Values are still returned if cells doesn't have formula.

<hr>

**ExcelSheet.get_cell_value(excel_pos : String) -> Variant**

Gets cell value by usual Excel cell adress (aka "A1", "C5", "F3" and etc).
```GDScript
var cell_value : Variant = sheet.get_cell_value("A1")
```
Returns empty string if nothing writed inside cell.

<hr>

**ExcelSheet.get_cell_formula(excel_pos : String) -> String**.

Gets cell formula by usual Excel cell adress.
```GDScript
var cell_formula : String = sheet.get_cell_formula("A1")
```
Returns empty string if nothing writed inside cell.

### Converting Excel cell adress to column-row index

**ExcelSheet.to_coord(excel_pos : String) -> Vector2i**

Helper function. Converts usual Excel cell adress to column and row indexes pair.

```GDScript
to_coords("B7") # returns Vector2i(2,7)
```

### Setting cells

There is two functions for setting cells:
<hr>

**ExcelSheet.set_cell(coord : Vector2i, data : Variant, is_formula : bool = false) -> Error**

Sets one cell inside xlsx sheet. Returns **OK** on edit success.

Parameter __*coord*__ is a column-row cell index. You can use *ExcelSheet.to_coords()* helper function to traslate Excel cell adress.

Allowed value type for parameter __*data*__ is int, float and String. In other cases function will return **ERR_INVALID_DATA**.

If __*is_formula*__ are true, sets cell formula instead of value. In this case, only Strings are allowed as data. Also "=" symbol at begining will be trimmed.

```GDScript
sheet.set_cell(Vector2i(3,5), 25)                     # Writes integer number 25 to cell "C5"
sheet.set_cell(Vector2i(4,6), 67.55)                  # Writes float number 67.55 to cell "D6"
sheet.set_cell(Vector2i(5,7), "Hello!")               # Writes string "Hello!" to cell "E7"
sheet.set_cell(to_coord("E12"), "=SUM(C5:D6)", true)  # Writes formula "SUM(C5:D6)" to cell "E12"
```

<hr>

**ExcelSheet.set_cells_range(range_str : String, data : Variant, is_formula : bool = false) -> Error**

Sets one or multiple cells inside xlsx sheet. Returns **OK** on edit success.

Parameter __*range_str*__ can be one cell adress ("A1") or cells range ("A1:B7"). Incorrect ranges as "B1:A4", "A6:C5" or "F3:A1" will be fixed to normal top-left and right-bottom pairs.

Allowed value type for parameter __*data*__ is int, float, String, PackedInt32Array, PackedInt64Array, PackedFloat32Array, PackedFloat64Array and Array. In other cases function will return **ERR_INVALID_DATA**.

In case of passed arrays, values will be assigned element by element in right-to-left and top-to-bottom order, until reaching end of array or cells range. Invalid values from usual Array will be skipped with cell instead of returning error.

If __*is_formula*__ are true, sets cell formula instead of value. In this case, only Strings are allowed as data (PackedStringArray support will be writed a bit later). Also "=" symbol at begining will be trimmed.

```GDScript
sheets.set_cells_range("A1:B8", 3)                         # Sets all cells value to 3 in range "A1:B8"
sheets.set_cells_range("A1:B8", [3, 2.45, true, "Hi!"])    # Result - "A1" - 3, "A2" - 2.45, "A3" - "", "A4" - "Hi!"
```
> [!IMPORTANT]
> After first edit, ExcelSheet will permenatly load data inside yourself and doesn't rely on xlsx data from now.
> 
> Same happens if ExcelFile object are closing, but bounded ExcelSheet still have reference somethere else inside your code.
> 
> In these cases your code will use more RAM then you can expect.

> [!TIP]
> It also means what you can still use *ExcelSheet* or *ExcelWorkbook* objects *<ins>even</ins>* after *ExcelFile.close()* or *ExcelFile.open()* call.
> 
> Only saving functionality will be disabled - editiong or reading will be avalible.


### Saving data

```GDScript
# Save opend file:
excel_file.save() # Returns OK on success. All refs to ExcelSheets and ExcelWorkbook objects are keeped, don't need to refresh them.

# To make copy of opend file somethere:
excel_file.save_as(path) # Returns OK on success. 
```
