## Represents an XML document.
class_name XMLDocument extends RefCounted

enum ENCODINGS {
	UTF8,
	UTF16,
	UTF32,
}

var file_declaration	: String

## The root XML node.
var root				: XMLNode


func _to_string():
	return "<XMLDocument root=%s>" % str(self.root)


# Dumps as full XML document
func get_file_encoding() -> ENCODINGS:
	if file_declaration.to_upper().contains("UTF-8"): return ENCODINGS.UTF8
	elif file_declaration.to_upper().contains("UTF-16"): return ENCODINGS.UTF16
	elif file_declaration.to_upper().contains("UTF-32"): return ENCODINGS.UTF32
	return ENCODINGS.UTF8
	

func dump_document_string() -> String:
	var declaration_str : String = "<?" + file_declaration + "?>\n"
	if is_instance_valid(root): return declaration_str + root.dump_str()
	elif file_declaration.strip_edges() == "": return ""
	return declaration_str


func dump_document_buffer() -> PackedByteArray:
	match get_file_encoding():
		ENCODINGS.UTF8:
			return dump_document_string().to_utf8_buffer()
		ENCODINGS.UTF16:
			return dump_document_string().to_utf16_buffer()
		ENCODINGS.UTF32:
			return dump_document_string().to_utf32_buffer()
		_:
			return []


func dump_document_file(path : String) -> Error:
	if !is_instance_valid(root): return ERR_FILE_CANT_WRITE
	
	if path.get_extension() != "xml": path = path.trim_suffix(path.get_extension()) + "xml"
	
	var file : FileAccess = FileAccess.open(path, FileAccess.WRITE)
	
	if !is_instance_valid(file): return FileAccess.get_open_error()
	
	file.store_buffer(dump_document_buffer())
	
	file.close()
	
	return OK
