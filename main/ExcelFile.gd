extends RefCounted

#Excel file handler.
class_name ExcelFile

var _zip_reader : ZIPReader

var _file_path	: String
var _workbook	: ExcelWorkbook

signal closing

func _notification(what: int) -> void:
	if what == NOTIFICATION_PREDELETE:
		if is_instance_valid(_workbook):
			_workbook._on_file_preclose(_zip_reader)
	
		if is_instance_valid(_zip_reader):
			_zip_reader.close()


func _is_open(isflag : FileDepdentRefCounted.IsFlag) -> void:
	isflag.flag = (is_instance_valid(_zip_reader) && _zip_reader.get_files().size() != 0)
	

func _on_get_data(pass_data : FileDepdentRefCounted.PassableData) -> void:
	pass_data.data_arr.append(_zip_reader)


func _give_file_reader(to_func : Callable) -> void:
	if !to_func.is_valid(): return
	to_func.call(_zip_reader)


static func open_file(path : String) -> ExcelFile:
	var excel_file : ExcelFile = ExcelFile.new()
	
	if excel_file.open(path) == OK: return excel_file
		
	return null


static func is_file_valid(path : String) -> bool:
	var reader : ZIPReader = ZIPReader.new()
	
	if !FileAccess.file_exists(path): return false
	
	if reader.open(path) != OK: return false
	
	var res : bool = ExcelWorkbook._is_file_valid(reader)
	
	reader.close()
	
	return res


func open(path : String) -> Error:
	if is_instance_valid(_workbook): close()
	
	if !FileAccess.file_exists(path): return ERR_FILE_BAD_PATH
	
	if !is_instance_valid(_zip_reader): _zip_reader = ZIPReader.new()
	
	var err : Error = _zip_reader.open(path)
	
	if err != OK: return err
	
	_workbook = ExcelWorkbook.new()
	
	err = _workbook.open(_zip_reader)
	
	if err != OK:
		_workbook = null
	else:
		_file_path = path
		
		_workbook.is_file_open.connect(_is_open)
		_workbook.get_data.connect(_on_get_data)
	
		closing.connect(_workbook._on_file_preclose, CONNECT_ONE_SHOT)
	
	return err
	
	
func get_sheets() -> Array[ExcelSheet]:
	if !is_instance_valid(_workbook): return []
	return _workbook.sheets


func save() -> Error:
	return save_as(_file_path)


func save_as(path : String) -> Error:	
	if !is_instance_valid(_workbook) || !_workbook._is_data_intact(): return ERR_UNCONFIGURED
	
	if path.get_extension() != "xlsx":
		path = path.get_basename() + ".xlsx"
	
	var data : Dictionary = _workbook._get_save_data()
	
	for file in _zip_reader.get_files():
		if file in data.keys(): continue
		data[file] = _zip_reader.read_file(file)
	
	var temp_archive_path 	: String 	= "TempExcel"
	var count 				: int 		= 1
	
	while temp_archive_path + str(count) in DirAccess.get_files_at(path.get_base_dir()):
		count += 1
		
	temp_archive_path = path.get_base_dir() + temp_archive_path + str(count)
	
	var _zip_packer : ZIPPacker = ZIPPacker.new()
	var err 		: Error 	= _zip_packer.open(temp_archive_path, ZIPPacker.APPEND_CREATE)
	
	if err != OK: return err
	
	for buffer_path in data.keys():
		if !buffer_path is String || !data[buffer_path] is PackedByteArray:
			err = ERR_INVALID_DATA
			break
			
		err = _zip_packer.start_file(buffer_path)
		
		if err != OK: break
		
		_zip_packer.write_file(data[buffer_path])
		
		if err != OK: break
	
	_zip_packer.close()
	
	if err != OK:
		DirAccess.remove_absolute(temp_archive_path)
		return err
		
	if path == _file_path:
		_zip_reader.close()
	
	err = DirAccess.copy_absolute(temp_archive_path, path)
	DirAccess.remove_absolute(temp_archive_path)
	
	if err != OK: return err
	
	if path == _file_path:
		err = _zip_reader.open(_file_path)
		if err != OK: return err
		
		for sheet in _workbook.sheets:
			sheet._edited = false
			sheet._clear_data()
		
		_workbook._on_save_sucess()
	
	return OK
	
	
func close() -> void:
	closing.emit(_zip_reader)
	_zip_reader.close()
	_zip_reader = null
	_workbook = null
