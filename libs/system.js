load_working_directory = function (){
	
	var path = CURRENT_FOLDER+"/.workspace";
    var workspace = read_all_text_file(path);
	if(!workspace){
		
		return CURRENT_FOLDER+INPUT_FOLDER+"\\";
	
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
        var items = rows[i].trim().split("="); 
        var new_row = "";
        if(items.length==2){
            new_row = ""+items[0]+"=\""+items[1]+ "\";";
        }
        rows[i] = new_row;
    }
    lib=rows.join('\n');
    log("configuration loaded.");
    eval(lib);
    
}





write_text_to_file = function (text,filepath){
    
    var ForReading = 1, ForWriting = 2;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var f = fso.OpenTextFile(filepath, ForWriting, true)
    f.Write(text);
    f.Close();
    
    
    
}

var stdin   = _script.StdIn;
var stdout  = _script.StdOut;


read_line = function (){
    var str= "";
    
    str += stdin.ReadLine();
       
    return str;
    
}

read = function (n){
    return stdin.Read(1);
}

read_all = function(){
    try{
		
		if (stdin.AtEndOfStream)
			return ("end of stream");
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

list_folders = function (path){
    
    var folders = [];
    fso = new ActiveXObject("Scripting.FileSystemObject");
    var folder = fso.GetFolder(path);
    var files = new Enumerator(folder.files);
    for (; !files.atEnd(); files.moveNext()){
   
        folders.push(files.item().path);
    
    }
    return folders;
}


function randomString(len, charSet) {
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
	
	
	if(format == 'YYYY/MM/DD' ){
		
		return (""+d.getFullYear())+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+(d.getDate())).slice(-2);
		
	}else{
		log('format not recognized')
		return d.toString();
	}
	
}

parse_date = function(ds,format){
	
	if(format == 'DD/MM/YYYY' ){
		
		var day   = d.substring(0,2);
		var month = d.substring(3,5);
		var year  = d.substring(6,10);
		var d = new Date();
		
		d.setFullYear(year);
		d.setMonth(parseInt(month)-1);
		d.setDate(parseInt(day))
		return d;
		//return (""+d.getFullYear())+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+(d.getDate())).slice(-2);
		
	}else{
		log('format not recognized')
		return d.toString();
	}
	
	
	
}


// ---------------------------------------------------------------------------
// File-system helpers
// ---------------------------------------------------------------------------

// Reads all text from a file. Returns "" for an empty file, throws on error.
read_text_file = function(path) {
    var ForReading = 1;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try {
        var f = fso.OpenTextFile(path, ForReading);
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
// http_request(url, method, callback [, body [, headers]])
// callback(responseText, statusCode)
// body    - optional string to send as request body (for POST/PUT)
// headers - optional plain object of request headers to set
http_request = function(url, method, reqListener, body, headers){
	if(['GET','POST','PUT','DELETE'].indexOf(method) == -1){
		throw 'method not recognized:' + method;
	}
	var request = new ActiveXObject("MSXML2.XMLHTTP.6.0");
	request.open(method, url, false);
	if (headers) {
		for (var key in headers) {
			if (Object.prototype.hasOwnProperty.call(headers, key)) {
				request.setRequestHeader(key, headers[key]);
			}
		}
	}
	request.send(body || null);
	reqListener(request.responseText, request.status);
}