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
	// Shared with the async/binary helpers below, so the accepted method list
	// lives in exactly one place.
	_http_check_method(method);
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

// ---------------------------------------------------------------------------
// sleep(ms)
// ---------------------------------------------------------------------------

// Blocks for ms milliseconds. There is no setTimeout in WSH; this is the only
// way to wait. Not available inside an HTA (no WScript object) - there the
// call is a no-op so shared code does not blow up.
sleep = function(ms) {
    // Deliberately does NOT test `WScript.Sleep` for truthiness first.
    // Property-getting a COM *method* raises "Object doesn't support this
    // property or method" in JScript - the guard would throw before the call
    // it was meant to protect. typeof on the global name is safe, and is all
    // that is needed: WScript exists under WSH and not inside an HTA.
    if (typeof WScript === "undefined") { return; }
    WScript.Sleep(ms);
};

// ---------------------------------------------------------------------------
// Asynchronous and binary HTTP
// ---------------------------------------------------------------------------
//
// http_request() above is synchronous: it blocks until the response arrives,
// so N requests cost the sum of their latencies. http_request_async() sends the
// request and returns immediately, so several can be in flight at once and the
// total cost is the slowest one rather than the sum.
//
//     var a = http_request_async("https://example.com/a", "GET", on_a);
//     var b = http_request_async("https://example.com/b", "GET", on_b);
//     http_wait_all([a, b], 30);      // both callbacks fire here
//
// Options object (all optional), shared by every function in this section:
//     body      string request body (POST/PUT)
//     headers   plain object of request headers
//     timeout   milliseconds, applied to each phase via setTimeouts()
//
// ServerXMLHTTP is used throughout: plain XMLHTTP.6.0 has no setTimeouts() and
// no waitForResponse(), so neither timeouts nor async would work there.

_http_methods = ['GET', 'POST', 'PUT', 'DELETE'];

_http_check_method = function(method) {
    if (_http_methods.indexOf(method) == -1) {
        throw 'method not recognized:' + method;
    }
};

// Creates, opens and configures a request. async decides open()'s third arg.
_http_open = function(url, method, options, async) {
    _http_check_method(method);
    options = options || {};
    var request = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
    request.open(method, url, !!async);
    if (typeof options.timeout === "number") {
        request.setTimeouts(options.timeout, options.timeout, options.timeout, options.timeout);
    }
    if (options.headers) {
        for (var key in options.headers) {
            if (Object.prototype.hasOwnProperty.call(options.headers, key)) {
                request.setRequestHeader(key, options.headers[key]);
            }
        }
    }
    return request;
};

// responseBody is a COM byte array, not a string. Round-tripping it through an
// ADODB.Stream is the same path read_binary_file() uses, so the byte values
// come back the same way here as they do from disk.
_http_response_bytes = function(request) {
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 1;              // adTypeBinary
    stream.Open();
    stream.Write(request.responseBody);
    if (stream.Size === 0) { stream.Close(); return []; }
    stream.Position = 0;
    var bytes = new VBArray(stream.Read()).toArray();
    stream.Close();
    return bytes;
};

// Sends a request and returns immediately, without waiting for the response.
//
// Returns a handle:
//     handle.wait([seconds])  block until the response arrives; fires the
//                             callback exactly once and returns true. Returns
//                             false if it timed out (call wait() again to keep
//                             waiting). With no argument, waits indefinitely.
//     handle.done             true once the callback has fired
//     handle.status           HTTP status, once done
//     handle.text             response text, once done
//     handle.headers          raw response headers, once done
//     handle.abort()          give up on the request
//
// callback(responseText, statusCode, responseHeaders) - the same shape
// http_request() uses.
http_request_async = function(url, method, callback, options) {

    options = options || {};

    var request = _http_open(url, method, options, true);
    request.send(options.body || null);

    var handle = {
        url:     url,
        method:  method,
        done:    false,
        aborted: false,
        status:  null,
        text:    null,
        headers: null
    };

    handle.wait = function(seconds) {
        if (handle.done || handle.aborted) { return handle.done; }

        // waitForResponse takes SECONDS (unlike options.timeout, which is
        // milliseconds, to match setTimeouts). Omitted means wait forever.
        var arrived = (typeof seconds === "number")
            ? request.waitForResponse(seconds)
            : request.waitForResponse();

        if (!arrived) { return false; }

        handle.status  = request.status;
        handle.text    = request.responseText;
        handle.headers = request.getAllResponseHeaders();
        handle.done    = true;
        if (callback) { callback(handle.text, handle.status, handle.headers); }
        return true;
    };

    handle.abort = function() {
        if (handle.done || handle.aborted) { return; }
        handle.aborted = true;
        try { request.abort(); } catch (e) {}
    };

    return handle;
};

// Waits for a batch of handles from http_request_async, firing each callback as
// its response arrives. Returns true only if every handle completed.
//
// The requests are already in flight by the time this is called, so waiting on
// them one after another still overlaps them: total time is the slowest
// request, not the sum. seconds, when given, is the budget for EACH handle.
http_wait_all = function(handles, seconds) {
    var all_done = true;
    for (var i = 0; i < handles.length; i++) {
        if (!handles[i].wait(seconds)) { all_done = false; }
    }
    return all_done;
};

// Synchronous request whose response is handed to the callback as an array of
// byte values (0-255) instead of text - for images, archives, anything that is
// not text. Same convention as read_binary_file/write_binary_file.
//
// callback(byteArray, statusCode, responseHeaders)
http_request_bytes = function(url, method, callback, options) {
    var request = _http_open(url, method, options, false);
    request.send((options && options.body) || null);
    callback(_http_response_bytes(request), request.status, request.getAllResponseHeaders());
};

// Downloads a URL straight to a file, without ever materialising the body as a
// JScript array - the response is streamed from COM to disk, so this is what to
// use for anything large.
//
// Returns the HTTP status code. Nothing is written unless the status is 2xx, so
// an error page never lands on disk pretending to be the file that was asked
// for; check the return value.
http_download_file = function(url, path, options) {
    options = options || {};
    var method  = options.method || 'GET';
    var request = _http_open(url, method, options, false);
    request.send(options.body || null);

    if (request.status < 200 || request.status > 299) { return request.status; }

    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 1;                   // adTypeBinary
    stream.Open();
    try {
        stream.Write(request.responseBody);
        stream.Position = 0;
        stream.SaveToFile(path, 2);    // adSaveCreateOverWrite
    } finally {
        stream.Close();
    }
    return request.status;
};
