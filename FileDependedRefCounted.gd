extends RefCounted

class_name FileDepdentRefCounted

class IsFlag extends RefCounted:
	var flag : bool = false

class PassableData extends RefCounted:
	var data_arr : Array

signal is_file_open	(isflag : IsFlag)
signal get_data (passdata : PassableData)


func is_dependency_open() -> bool:
	var isflag : IsFlag = IsFlag.new()
	is_file_open.emit(isflag)
	return isflag.flag


func get_dependency_data() -> Array:
	var pass_data : PassableData = PassableData.new()
	get_data.emit(pass_data)
	return pass_data.data_arr
