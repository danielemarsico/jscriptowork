load_working_directory = function (){
	
	var path = CURRENT_FOLDER+"/.workspace";
    var workspace = read_all_text_file(path);
	if(!workspace){

		var input_folder = typeof INPUT_FOLDER !== "undefined" ? INPUT_FOLDER : "";
		return CURRENT_FOLDER+input_folder+"\\";

	}else{
		
		return workspace.trim();
	}
	
}

save_working_directory = function(working_directory){
	
	var path = CURRENT_FOLDER+"/.workspace";
	
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try{
		
		var f = fso.CreateTextFile(path, true);
		f.WriteLine(working_directory);
		f.Close();
	}
	catch(exc){
		log(exc.message);
		return null;
		
	}
	
	
	
}

load_properties = function (properties_file){
    
    
    var path = CURRENT_FOLDER+"/"+properties_file;
    var lib = read_all_text_file(path);
    var rows = lib.split("\n");
    log("reading configuration...");
    for(var i= 0;i < rows.length; i++){
        var row = rows[i].trim();
        var new_row = "";
        if(row !== "" && row.charAt(0) !== "#"){
            var eq = row.indexOf("=");
            if(eq !== -1){
                var key   = row.substring(0, eq).trim();
                var value = row.substring(eq + 1)
                    .replace(/\\/g, "\\\\")
                    .replace(/"/g, "\\\"");
                new_row = key+"=\""+value+"\";";
            }
        }
        rows[i] = new_row;
    }
    lib=rows.join('\n');
    log("configuration loaded.");
    eval(lib);
    
}





// Writes text to a file. Pass unicode=true to write UTF-16LE (needed for
// non-ASCII content); defaults to ASCII for backward compatibility.
write_text_to_file = function (text,filepath,unicode){

    var ForWriting = 2;
    var TristateUnicode = -1;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var f = fso.OpenTextFile(filepath, ForWriting, true, unicode ? TristateUnicode : 0);
    try{
        f.Write(text);
    } finally {
        f.Close();
    }

}

stdin   = _script.StdIn;
stdout  = _script.StdOut;


read_line = function (){
    var str= "";

    str += stdin.ReadLine();

    return str;

}

read = function (n){
    return stdin.Read(n);
}

read_all = function(){
    try{

		if (stdin.AtEndOfStream)
			return ("");
		else
			return (stdin.ReadAll());


	}
	catch(exc){
		log("can't read from stdin");
		return null;

	}
}

write_line = function (data){
    
     stdout.WriteLine(data);
    
}

write = function (data){
    
    stdout.Write(data);
    
}

// Returns the full paths of every file directly inside path (not folders,
// despite the name list_folders kept below for backward compatibility).
list_files = function (path){

    var files = [];
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var folder = fso.GetFolder(path);
    var enumerator = new Enumerator(folder.files);
    for (; !enumerator.atEnd(); enumerator.moveNext()){

        files.push(enumerator.item().path);

    }
    return files;
}

// Kept for backward compatibility: despite the name, this lists files (see TODO.md).
list_folders = list_files;

// Returns the full paths of every subfolder directly inside path.
list_subfolders = function (path){

    var folders = [];
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var folder = fso.GetFolder(path);
    var enumerator = new Enumerator(folder.subfolders);
    for (; !enumerator.atEnd(); enumerator.moveNext()){

        folders.push(enumerator.item().path);

    }
    return folders;
}


randomString = function(len, charSet) {
    charSet = charSet || 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var randomString = '';
    for (var i = 0; i < len; i++) {
    	var randomPoz = Math.floor(Math.random() * charSet.length);
    	randomString += charSet.substring(randomPoz,randomPoz+1);
    }
    return randomString;
}


//format date object, currently supports YYYY/MM/DD, YY/MM/DD, YYYYMMDD
format_date = function(d,format){

	var yyyy = ""+d.getFullYear();
	var mm   = ("0"+(d.getMonth()+1)).slice(-2);
	var dd   = ("0"+(d.getDate())).slice(-2);

	if(format == 'YYYY/MM/DD' ){

		return yyyy+"/"+mm+"/"+dd;

	}else if(format == 'YY/MM/DD' ){

		return yyyy.slice(-2)+"/"+mm+"/"+dd;

	}else if(format == 'YYYYMMDD' ){

		return yyyy+mm+dd;

	}else{
		log('format not recognized')
		return d.toString();
	}

}

parse_date = function(ds,format){

	var d = new Date();

	if(format == 'DD/MM/YYYY' ){

		var day   = ds.substring(0,2);
		var month = ds.substring(3,5);
		var year  = ds.substring(6,10);

		d.setFullYear(year);
		d.setMonth(parseInt(month)-1);
		d.setDate(parseInt(day))
		return d;
		//return (""+d.getFullYear())+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+(d.getDate())).slice(-2);

	}else{
		log('format not recognized')
		return ds;
	}



}


// ---------------------------------------------------------------------------
// File-system helpers
// ---------------------------------------------------------------------------

// Reads all text from a file. Returns "" for an empty file, throws on error.
read_text_file = function(path) {
    var ForReading = 1;
    var TristateUseDefault = -2; // auto-detects a Unicode BOM; ASCII files are unaffected
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try {
        var f = fso.OpenTextFile(path, ForReading, false, TristateUseDefault);
        if (f.AtEndOfStream) { f.Close(); return ""; }
        var content = f.ReadAll();
        f.Close();
        return content;
    } catch(exc) {
        throw new Error("read_text_file: " + exc.message);
    }
};

// Returns true if a file exists at path.
file_exists = function(path) {
    return (new ActiveXObject("Scripting.FileSystemObject")).FileExists(path);
};

// Returns true if a folder exists at path.
folder_exists = function(path) {
    return (new ActiveXObject("Scripting.FileSystemObject")).FolderExists(path);
};

// Deletes a file. Throws if the file does not exist.
delete_file = function(path) {
    (new ActiveXObject("Scripting.FileSystemObject")).DeleteFile(path);
};

// Creates a single folder. No-op if it already exists.
create_folder = function(path) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (!fso.FolderExists(path)) { fso.CreateFolder(path); }
};

// Writes an array of integers (0-255) to a binary file using ADODB.Stream.
// Each integer is stored as one raw byte (iso-8859-1, no BOM).
// Note: byte values 128-159 may not round-trip on some locales (Windows-1252 overlap).
write_binary_file = function(path, bytes) {
    var str = "";
    for (var i = 0; i < bytes.length; i++) {
        str += String.fromCharCode(bytes[i] & 0xFF);
    }
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type    = 2;           // adTypeText
    stream.CharSet = "iso-8859-1";
    stream.Open();
    stream.WriteText(str);
    stream.SaveToFile(path, 2);   // adSaveCreateOverWrite
    stream.Close();
};

// Reads a binary file and returns an array of integers (0-255).
read_binary_file = function(path) {
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 1; // adTypeBinary
    stream.Open();
    stream.LoadFromFile(path);
    if (stream.Size === 0) { stream.Close(); return []; }
    var bytes = new VBArray(stream.Read()).toArray();
    stream.Close();
    return bytes;
};

// ---------------------------------------------------------------------------
// http_request(url, method, callback [, body [, headers [, timeout]]])
// callback(responseText, statusCode, responseHeaders)
// body            - optional string to send as request body (for POST/PUT)
// headers         - optional plain object of request headers to set
// timeout         - optional milliseconds; throws if any phase exceeds it
// responseHeaders - raw header block as returned by getAllResponseHeaders()
//
// Still synchronous (open(..., false)) despite the callback-shaped API: true
// async/streaming HTTP is tracked as a separate feature in TODO.md.
http_request = function(url, method, reqListener, body, headers, timeout){
	if(['GET','POST','PUT','DELETE'].indexOf(method) == -1){
		throw 'method not recognized:' + method;
	}
	// ServerXMLHTTP, not plain XMLHTTP: XMLHTTP.6.0 does not implement
	// setTimeouts() at all (it throws "Object doesn't support this property
	// or method"), so a timeout could never actually take effect there.
	var request = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
	request.open(method, url, false);
	if (typeof timeout === "number") {
		request.setTimeouts(timeout, timeout, timeout, timeout);
	}
	if (headers) {
		for (var key in headers) {
			if (Object.prototype.hasOwnProperty.call(headers, key)) {
				request.setRequestHeader(key, headers[key]);
			}
		}
	}
	request.send(body || null);
	reqListener(request.responseText, request.status, request.getAllResponseHeaders());
}