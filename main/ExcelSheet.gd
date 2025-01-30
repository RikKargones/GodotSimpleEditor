extends FileDepdentRefCounted

class_name ExcelSheet

var _symbols_reg			: RegEx			= RegEx.new()
var _cell_reg				: RegEx			= RegEx.new()
var _range_reg				: RegEx			= RegEx.new()

var _was_buffered			: bool			= false
var _buffered_formula_flag	: bool			= false
var _edited					: bool			= false

var _rid					: int			= -1
var _name 					: String
var _path_in_file 			: String
var _buffered_table			: Dictionary

var xml_data				: XMLDocument
var _sheet_data_node		: XMLNode
var _sheet_rows_nodes		: Dictionary

signal needs_reading


func _init() -> void:
	_symbols_reg.compile("[A-Z]+")
	_cell_reg.compile("([A-Z]+)([0-9]+)")
	_range_reg.compile("([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)")


func _clear_buffer() -> void:
	_was_buffered = false
	_buffered_table.clear()
	

func _clear_data(keep_buffer : bool = false) -> void:
	xml_data 			= null
	_sheet_data_node 	= null
	
	_sheet_rows_nodes.clear()
	if !keep_buffer: _clear_buffer()
	
	
func _read(excel_file_zip_reader : ZIPReader, global_shared_strs : Array, autoclose : bool = false) -> Error:
	_clear_data()
	
	if !is_instance_valid(excel_file_zip_reader): return ERR_FILE_CANT_READ
	
	if !excel_file_zip_reader.file_exists(_path_in_file): return ERR_FILE_NOT_FOUND
	
	xml_data = XML.parse_buffer(excel_file_zip_reader.read_file(_path_in_file))
	
	if !is_instance_valid(xml_data.root):
		xml_data = null
		return ERR_FILE_CORRUPT
	
	_sheet_data_node = xml_data.root.seek_first_child_by_name("sheetData")
	
	if !is_instance_valid(_sheet_data_node): return ERR_FILE_CORRUPT
	
	for row in _sheet_data_node.find_children_by_name("row"):
		if !row.attributes.has("r") || !row.attributes["r"].is_valid_int(): continue
		_sheet_rows_nodes[row.attributes["r"].to_int()] = row
		
	_sheet_data_replace_from_shared_strings(global_shared_strs)
	
	if autoclose: _clear_data()
	
	return OK


func _sheet_data_replace_from_shared_strings(global_shared_strs : Array) -> Error:
	for cell in _sheet_data_get_cells_with_shared_strings():
		var value_node : XMLNode = cell.seek_first_child_by_name("v")
		
		if !value_node.content.is_valid_int(): return ERR_INVALID_DATA
		
		var shared_idx : int = value_node.content.to_int()
		
		if shared_idx < 0 || shared_idx >= global_shared_strs.size(): return ERR_DOES_NOT_EXIST
		
		value_node.content = global_shared_strs[shared_idx]
		
	return OK


func _sheet_data_replace_to_shared_strs(global_shared_strs : Array) -> Error:
	for cell in _sheet_data_get_cells_with_shared_strings():
		var value_node 	: XMLNode 	= cell.seek_first_child_by_name("v")
		var value_pos	: int		= global_shared_strs.find(value_node.content)
		
		if value_pos == -1: return ERR_DOES_NOT_EXIST
		
		value_node.content = str(value_pos)
	
	return OK


func _needs_reading() -> bool:
	return !is_instance_valid(xml_data) || !is_instance_valid(xml_data.root) || !is_instance_valid(_sheet_data_node)

#Checking need to read data from file, if it is, try loading.
#After returns false if need to read file still exist.
func _is_data_redable() -> bool:
	if _needs_reading():
		needs_reading.emit()
		return !_needs_reading()
	else:
		return true


func _get_cell_coord_from_xml(cell : XMLNode) -> Vector2i:
	if !is_instance_valid(cell) || !cell.attributes.has("r"): return Vector2i(0,0)
	
	return to_coords(cell.attributes["r"])


func _get_cell_formula_from_xml(cell : XMLNode) -> Variant:
	if !is_instance_valid(cell): return null
	
	var formula_node : XMLNode = cell.seek_first_child_by_name("f")
	
	if !is_instance_valid(formula_node): return null
	
	return str(formula_node.content)
	

func _get_cell_value_from_xml(cell : XMLNode) -> Variant:
	if !is_instance_valid(cell): return null
	
	var value_node : XMLNode = cell.seek_first_child_by_name("v")
	
	if !is_instance_valid(value_node): return null
	
	if cell.attributes.has("t") && cell.attributes["t"] == "s": return value_node.content
	
	var json 	: JSON 		= JSON.new()
	var value 	: String 	= value_node.content
	
	if json.parse(value) == OK: return json.data
	return value


func _sheet_data_has_row(row : int) -> bool:
	return _sheet_rows_nodes.has(row)


func _sheet_data_get_cells_from_row(row : int) -> Array[XMLNode]:
	if !_is_data_redable() || !_sheet_data_has_row(row): return []
	
	return _sheet_rows_nodes[row].find_children_by_name("c")


func _sheet_data_get_cells_with_shared_strings() -> Array[XMLNode]:
	var shared_string_cells : Array[XMLNode]
	
	for row in _sheet_rows_nodes.keys():
		for cell in _sheet_data_get_cells_from_row(row):
			if !cell.attributes.has("t") || cell.attributes["t"] != "s": continue
			
			var value_node : XMLNode = cell.seek_first_child_by_name("v")
			
			if !is_instance_valid(value_node): continue
			
			shared_string_cells.append(cell)
			
	return shared_string_cells
	

func _sheet_data_get_cell(coord : Vector2i) -> XMLNode:
	if coord.x < 1 || coord.y < 1: return null
	
	var cells : Array[XMLNode] = _sheet_data_get_cells_from_row(coord.y)
	
	for cell in cells:
		if coord == _get_cell_coord_from_xml(cell): return cell
		
	return null


#Writes new empty row if this one not exist
func _sheet_data_write_row(row : int) -> Error:
	if !_is_data_redable(): return ERR_FILE_CANT_READ
	
	if _sheet_data_has_row(row): return OK
	
	var rows_keys		: Array = _sheet_rows_nodes.keys()
	var insert_pos 		: int 	= 0
	
	rows_keys.sort()
	
	for idx in rows_keys:
		if _sheet_rows_nodes.keys()[idx] > row: break
		insert_pos += 1
	
	var new_row : XMLNode = XMLNode.new()
	
	new_row.name					= "row"
	new_row.attributes["r"] 		= str(row)
	new_row.attributes["spans"] 	= "0"
	
	_sheet_data_node.children.insert(insert_pos, new_row)
	_sheet_rows_nodes[row] = new_row
	
	return OK


func _sheet_data_update_row_spans(row : int) -> void:
	var cells : Array[XMLNode] = _sheet_data_get_cells_from_row(row)
	
	if cells.size() == 0: return
	
	var start 	: int = 0
	var end 	: int = 0
	
	for cell in cells:
		var coord : Vector2i = _get_cell_coord_from_xml(cell)
		if coord.x < 1 || coord.y < 1: continue
		if start == 0 || coord.x < start: start = coord.x
		if end == 0 || coord.x > end: end = coord.x
		
	if start == end: _sheet_rows_nodes[row].attributes["spans"] = str(start)
	else: _sheet_rows_nodes[row].attributes["spans"] = str(start) + ":" + str(end)


#Writes empty cell (with empty string in value) if this one not exist
func _sheet_data_write_cell(coord : Vector2i, add_formula : bool = false) -> Error:
	if coord.x < 1 || coord.y < 1: return ERR_INVALID_PARAMETER
	
	if _sheet_data_write_row(coord.y) != OK: return ERR_FILE_CANT_READ
	
	var cells 		: Array[XMLNode] 	= _sheet_data_get_cells_from_row(coord.y)
	var insert_pos 	: int 				= 0
	
	var target_cell : XMLNode
	
	var has_value	: bool 				= false
	var has_formula	: bool 				= false
	
	for cell in cells:
		var cell_coord : Vector2i = _get_cell_coord_from_xml(cell)
		
		if cell_coord.x == coord.x:
			target_cell 	= cell
			has_value 		= _get_cell_value_from_xml(target_cell) 	!= null
			has_formula 	= _get_cell_formula_from_xml(target_cell) 	!= null
			
			target_cell.standalone = false
			
			if add_formula && has_formula && has_value: return OK
			elif !add_formula && has_value: return OK
			else: break
			
		if cell_coord.x > coord.x: break
		insert_pos += 1
	
	var row_node : XMLNode = _sheet_rows_nodes[coord.y]
	
	if !is_instance_valid(target_cell):
		target_cell						= XMLNode.new()
		target_cell.name 				= "c"
		target_cell.attributes["r"] 	= column_num_to_char(coord.x) + str(coord.y)
		row_node.children.insert(insert_pos, target_cell)
	
	if !has_value:
		var value_node : XMLNode = XMLNode.new()
		value_node.name	= "v"
		value_node.standalone = true
		target_cell.children.append(value_node)
	
	if add_formula && !has_formula:
		var formula_node : XMLNode = XMLNode.new()
		formula_node.name = "f"
		formula_node.standalone = true
		target_cell.children.insert(0, formula_node)
	
	_sheet_data_update_row_spans(coord.y)
	
	return OK	


func get_sheet_name() -> String:
	return _name
	

func _get_sheet_rid() -> int:
	return _rid

	
func get_table(return_formula_results : bool = true) -> Dictionary:
	if _was_buffered && return_formula_results == _buffered_formula_flag:
		return _buffered_table.duplicate(true)
	
	_clear_buffer()
	
	if !_is_data_redable():
		printerr("Looks like Excel book was deleted from memory.\nOpen file again! (This object is useless now...)")
		return {}
	
	for row_idx in _sheet_rows_nodes.keys():
		var row_data 	: Dictionary
		var cells 		: Array[XMLNode] = _sheet_data_get_cells_from_row(row_idx)
		
		for cell in cells:
			var coords : Vector2i = _get_cell_coord_from_xml(cell)
			
			if coords.x < 1 || coords.y < 1: continue
			
			if !return_formula_results:
				var formula_value : Variant = _get_cell_formula_from_xml(cell)
				if formula_value != null:
					row_data[coords.x] = formula_value
					continue
				
			var cell_value : Variant = _get_cell_value_from_xml(cell)
			if cell_value != null: row_data[coords.x] = cell_value
			
		if row_data.size() < 1: continue
		
		_buffered_table[row_idx] = row_data
	
	_was_buffered = true
	_buffered_formula_flag = return_formula_results
	
	if !_edited: _clear_data(true)
	
	return _buffered_table.duplicate(true)


func column_num_to_char(num : int) -> String:
	if num < 1: return ""
	
	var chars : String
	
	while num > 0:
		chars = char(65 + (num - 1) % 26) + chars
		num = int((num - 1) / 26)
		
	return chars


func column_char_to_num(simb : String) -> int:
	if _symbols_reg.search(simb) == null: return 0
	
	var x 				: int = 0
	var column_length 	: int = simb.length()
	
	for i in column_length:
		var num = (simb.unicode_at(i) - 64)
		x += num * int(pow(26, column_length - 1 - i))
		
	return x

	
func to_coords(excel_pos : String) -> Vector2i:
	var result : RegExMatch = _cell_reg.search(excel_pos)
	
	if result == null: return Vector2i(0,0)
	
	var column_str 		: String 	= result.get_string(1)
	var row_str 		: String	= result.get_string(2)
	
	return Vector2i(column_char_to_num(column_str), row_str.to_int())
	
	
func get_cell_value(excel_pos : String) -> Variant:
	var value : Variant = _get_cell_value_from_xml(_sheet_data_get_cell(to_coords(excel_pos)))
	
	if value != null: return value
	return ""


func get_cell_formula(excel_pos : String) -> String:
	var value : String = _get_cell_formula_from_xml(_sheet_data_get_cell(to_coords(excel_pos)))
	
	if value != null: return value
	return ""


func set_cell(coord : Vector2i, data, is_formula : bool = false) -> Error:
	var err : Error = _sheet_data_write_cell(coord, is_formula)
	
	if err != OK: return err
	
	var cell : XMLNode = _sheet_data_get_cell(coord)
	
	if !is_instance_valid(cell): return ERR_BUG
	
	var value_node 		: XMLNode 	= cell.seek_first_child_by_name('v')
	var formula_node 	: XMLNode 	= cell.seek_first_child_by_name("f")
	
	if is_formula:
		if !data is String: return ERR_INVALID_DATA
		_edited = true
		formula_node.content = data.trim_prefix("=").strip_edges()
		formula_node.standalone = formula_node.content.strip_edges() == ""
		#value_node.content = ""
		#value_node.standalone = true
		_clear_buffer()
		return OK
	
	var has_formula : bool = is_instance_valid(formula_node) && formula_node.content.strip_edges() != ""
	
	if has_formula:
		cell.attributes["t"] = "str"
	
	match typeof(data):
		TYPE_STRING:
			_edited = true
			if !has_formula: cell.attributes["t"] = "s"
			value_node.content = data
			value_node.standalone = value_node.content.strip_edges() == ""
			
		TYPE_INT, TYPE_FLOAT:
			_edited = true
			if !has_formula: cell.attributes.erase("t")
			value_node.content = str(data)
			value_node.standalone = value_node.content.strip_edges() == ""
		_:
			return ERR_INVALID_DATA
	
	_clear_buffer()
	
	return OK
	
	
func set_cells_range(range_str : String, data, is_formula : bool = false) -> Error:
	var res 		: RegExMatch 	= _cell_reg.search(range_str)
	var res_exp 	: RegExMatch 	= _range_reg.search(range_str)
	
	var coord_one 	: Vector2i 		= Vector2(0,0)
	var coord_two	: Vector2i 		= Vector2(0,0)
	
	if res_exp != null:	
		coord_one = to_coords(res_exp.get_string(1) + res_exp.get_string(2))
		coord_two = to_coords(res_exp.get_string(3) + res_exp.get_string(4))
	elif res != null:
		coord_one = to_coords(res.get_string(1) + res.get_string(2))
		coord_two = coord_one
	else:
		return ERR_INVALID_PARAMETER
		
	if coord_one.x < 0 || coord_one.y < 0: return ERR_INVALID_PARAMETER
	
	if coord_two == Vector2i(0,0) || coord_one == coord_two:
		return set_cell(coord_one, data, is_formula)
	
	if (coord_two.x < 0 || coord_two.y < 0): return ERR_INVALID_PARAMETER
	
	#Fixing to standart cell range format: first cell - top left, end cell - bottom right
	if coord_one.x > coord_two.x:
		var swap : int = coord_one.x
		coord_one.x = coord_two.x
		coord_two.x = swap
	
	if coord_one.y > coord_two.y:
		var swap : int = coord_one.y
		coord_one.y = coord_two.y
		coord_two.y = swap
	
	var size_c 	: int 		= coord_two.x - coord_one.x + 1
	var size_r 	: int 		= coord_two.y - coord_one.y + 1
	
	var err 	: Error
	
	match typeof(data):
		TYPE_STRING, TYPE_INT, TYPE_FLOAT:
			for row in size_r:
				for column in size_c:
					err = set_cell(coord_one + Vector2i(column, row), data, is_formula)
					if err != OK && err != ERR_INVALID_PARAMETER: return err
			
		TYPE_ARRAY, TYPE_PACKED_FLOAT32_ARRAY, TYPE_PACKED_FLOAT64_ARRAY, TYPE_PACKED_INT32_ARRAY, TYPE_PACKED_INT64_ARRAY:
			for row in size_r:
				for column in size_c:
					var idx : int = column + row * size_c
					if idx >= data.size(): return OK
					err = set_cell(coord_one + Vector2i(column, row), data[idx], is_formula)
					if err != OK && err != ERR_INVALID_PARAMETER && err != ERR_INVALID_DATA: return err
		_:
			return ERR_INVALID_DATA
	
	return OK
	
	
