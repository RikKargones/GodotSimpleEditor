# GodotSimpleXlsxEditor
A very simple Excel .xlsx file editor writed in pure GDScript.

Created for task on work - writing data (only values/formulas) in cells.

Code uses a bit changed [GodotXML addon by elenakrittik](https://github.com/elenakrittik/GodotXML) and stracturly following [Godot Excel Reader addon by LaoDie1](https://github.com/LaoDie1/godot-excel-reader?ysclid=m6j75bap25275265483).

Most of additinal code is a an actual edit cell values/formulas and saving xlsx files.

## What SimpleXlsxEditor CAN do:
- **Edit cells inside sheets.**
- **Save edited data into xlsx file (opend original or copy)**

## What SimpleXlsxEditor CAN'T do:
- **Opening and editing xls files.** - Data stracture inside them are very, very different. Maybe sometime I write support for BIFF, but likely not.
- **Creating something bigger then cell** - sheets, styles and file itself.
- **Calculating formula result in cell**

## SimpleXlsxEditor differenses inside addons:
### Godot Excel Reader:
- **Really checking file validness** - no crash happening while trying to open xls file, non-xlsx file or corrupted xlsx. Needed for read/edit data are checked on opening.
- **Replaced XML reader code** - original one only reads data from XML files inside xlsx. GodotXML allows edit data.
### GodotXML
- **Seeking and collecting funcs** - as *seek_first_child_by_name()* and *get_childs_by_name()*.
- **XMLDocument keeps XML header for later document save** - (silly one - also trying to know in which UTF-type document should be saved)
- **Fixed: dump_str without *pretty* parameter makes XML file unredable by browser or Excel because of skipped spaces between attributes**.
