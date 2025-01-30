extends FileDepdentRefCounted

class_name ExcelWorkbook

var xml_data	 			: XMLDocument

var sheets 					: Array[ExcelSheet]

var _shared_strings			: XMLDocument
var _shared_strings_arr		: Array
var _edited_shared_strings	: Array

var _rels 					: XMLDocument
var _rid_to_path_map	 	: Dictionary


const xlsx_valid_files 	: PackedStringArray = [
	"xl/workbook.xml",
	"xl/_rels/workbook.xml.rels",
	"xl/sharedStrings.xml",
	]


static func _is_file_valid(file_reader : ZIPReader) -> bool:
	if !is_instance_valid(file_reader): return false
	
	var valid : bool = true
	
	for file in xlsx_valid_files:
		valid = valid && file_reader.file_exists(file)
	
	return valid


func _is_file_open(isflag : FileDepdentRefCounted.IsFlag) -> void:
	is_file_open.emit(isflag)


func _get_shared_string_data() -> Array:
	return _shared_strings_arr.duplicate()
	
	
func _reform_shared_string_arr() -> void:
	_shared_strings_arr.clear()
	
	if !is_instance_valid(_shared_strings) || !is_instance_valid(_shared_strings.root): return
	
	for si in _shared_strings.root.find_children_by_name("si"):
		for child in si.children:
			if child.name != "t": continue
			_shared_strings_arr.append(child.content)
			
	_edited_shared_strings = _shared_strings_arr.duplicate()
	

func _on_file_preclose(excel_file_zip_reader : ZIPReader) -> void:
	for sheet in sheets:
		if sheet.get_reference_count() < 2 && !get_reference_count() < 2: continue
		sheet._read(excel_file_zip_reader, _get_shared_string_data())
		sheet._edited = true


# Frees memory after save, but preserving sheets objects - for avalibility from other parts of code
func _on_save_sucess() -> void:
	for sheet in sheets:
		sheet._edited = false
		sheet._clear_data()
	
	var old_sheets 	: Array[ExcelSheet] = sheets.duplicate()
	var err 		: Error 			= open(_take_zip_reader())
	
	if err != OK:
		printerr("Looks like something was breaked in saving process. Re-opening file returns an error: " + error_string(err))
		return
	
	sheets = old_sheets


func _clear_data() -> void:
	xml_data 		= null
	_shared_strings = null
	_rels 			= null
	
	sheets.clear()
	_shared_strings_arr.clear()
	_edited_shared_strings.clear()
	_rid_to_path_map.clear()


func _is_data_intact() -> bool:
	if !is_instance_valid(_rels) || !is_instance_valid(_shared_strings) || !is_instance_valid(xml_data): return false
	return is_instance_valid(_rels.root) && is_instance_valid(_shared_strings.root) && is_instance_valid(xml_data.root)


func _take_zip_reader() -> ZIPReader:
	var reader : Array = get_dependency_data()
	if reader[0] is ZIPReader: return reader[0]
	return null


func _emit_take_reader(callable : Callable, sheet : ExcelSheet) -> void:
	if !callable.is_valid() || !is_instance_valid(sheet): return
	
	var zip_reader : ZIPReader = _take_zip_reader()
	
	if !is_instance_valid(zip_reader): return
	
	# Edited sheet recives different shared strings arrays to read
	if sheet._edited: callable.call(zip_reader, _edited_shared_strings.duplicate())
	else: callable.call(zip_reader, _shared_strings_arr.duplicate())


func _get_save_data() -> Dictionary:
	if !is_dependency_open() || !_is_data_intact(): return {}
	
	var paths_and_data : Dictionary = {
		"xl/_rels/workbook.xml.rels" 	: _rels.dump_document_buffer(),
		"xl/workbook.xml"				: xml_data.dump_document_buffer(),
	}
	
	var existing_shared_strings 	: Array
	var total_count					: int 	= 0
	
	for sheet in sheets:
		for shared_string_cell in sheet._sheet_data_get_cells_with_shared_strings():
			var value_node : XMLNode = shared_string_cell.seek_first_child_by_name("v")
			
			total_count += 1
			
			if value_node.content in existing_shared_strings: continue
			
			existing_shared_strings.append(value_node.content)
	
	_edited_shared_strings = existing_shared_strings
	
	var zip_reader : ZIPReader = _take_zip_reader()
	
	if !is_instance_valid(zip_reader) || !"xl/sharedStrings.xml" in zip_reader.get_files(): return {}
	
	var shared_strings_copy : XMLDocument = XML.parse_buffer(zip_reader.read_file("xl/sharedStrings.xml"))
	
	shared_strings_copy.root.attributes["count"] 		= total_count
	shared_strings_copy.root.attributes["uniqueCount"] 	= _edited_shared_strings.size()
	
	for si in shared_strings_copy.root.find_children_by_name("si"):
		shared_strings_copy.root.children.erase(si)
		
	for s_str in existing_shared_strings:
		var si_node : XMLNode = XMLNode.new()
		var t_node	: XMLNode = XMLNode.new()
		
		shared_strings_copy.root.children.append(si_node)
		si_node.children.append(t_node)
		
		si_node.name = "si"
		t_node.name = "t"
		t_node.content = s_str
	
	paths_and_data["xl/sharedStrings.xml"] = shared_strings_copy.dump_document_buffer()
	
	# Bringing sheets xml's to savebale state, where cells with shared strings
	# contains position in upadted sharedStrings.xml file.
	# Edited sheets will be converted back, readed - cleared.
	for sheet in sheets:
		sheet._sheet_data_replace_to_shared_strs(existing_shared_strings)
		paths_and_data[sheet._path_in_file] = sheet.xml_data.dump_document_buffer()
		if sheet._edited: sheet._sheet_data_replace_from_shared_strings(existing_shared_strings)
		else: sheet._clear_data()
	
	return paths_and_data


func open(excel_file_zip_reader : ZIPReader) -> Error:
	if !is_instance_valid(excel_file_zip_reader): 	return ERR_FILE_CANT_READ
	if !_is_file_valid(excel_file_zip_reader): 		return ERR_FILE_UNRECOGNIZED
	
	_clear_data()
	
	#Extracting rid-to-path info
	_rels = XML.parse_buffer(excel_file_zip_reader.read_file("xl/_rels/workbook.xml.rels"))
	
	if !is_instance_valid(_rels.root): return ERR_FILE_CORRUPT
	
	for child in _rels.root.children:
		if !child.attributes.has_all(["Id", "Target"]): continue
		_rid_to_path_map[child.attributes["Id"]] = child.attributes["Target"]
	
	
	#Getting shared strings - sheets uses them as ref to text in cells
	_shared_strings = XML.parse_buffer(excel_file_zip_reader.read_file("xl/sharedStrings.xml"))
	
	if !is_instance_valid(_shared_strings.root):
		_clear_data()
		return ERR_FILE_CORRUPT
	
	_reform_shared_string_arr()
	
	#Extracting info about sheets. If none sheets found - file defently corrupted.
	xml_data = XML.parse_buffer(excel_file_zip_reader.read_file("xl/workbook.xml"))
	
	if !is_instance_valid(xml_data.root):
		_clear_data()
		return ERR_FILE_CORRUPT
	
	var sheets_root : XMLNode = xml_data.root.seek_first_child_by_name("sheets")
	
	if !is_instance_valid(sheets_root):
		_clear_data()
		return ERR_FILE_CORRUPT
	
	# Creating sheets, but their data whould't kept long in memory - permanent load will be only on edit demand
	for node in sheets_root.children:
		if node.name != "sheet" || !node.attributes.has_all(["name", "r:id"]): continue
		
		if !node.attributes["r:id"] in _rid_to_path_map.keys() || !excel_file_zip_reader.file_exists("xl/" + _rid_to_path_map[node.attributes["r:id"]]):
			_clear_data()
			return ERR_FILE_CORRUPT
		
		var new_sheet : ExcelSheet = ExcelSheet.new()
		
		new_sheet._rid 				= str(node.attributes["r:id"]).to_int()
		new_sheet._name 			= node.attributes["name"]
		new_sheet._path_in_file 	= "xl/" + _rid_to_path_map[node.attributes["r:id"]]
		
		var err : Error = new_sheet._read(excel_file_zip_reader, _get_shared_string_data(), true)
		
		if err != OK:
			_clear_data()
			return ERR_FILE_CORRUPT
		
		new_sheet.is_file_open.connect(_is_file_open)
		new_sheet.needs_reading.connect(_emit_take_reader.bind(new_sheet._read, new_sheet))
		
		sheets.append(new_sheet)
	
	return OK
	

func get_sheet_name_list() -> PackedStringArray:
	return PackedStringArray(sheets.map(func(item): return item._name))


func get_sheet_by_name(sheet_name : String) -> ExcelSheet:
	for sheet in sheets:
		if sheet._name == sheet_name: return sheet
	return null
