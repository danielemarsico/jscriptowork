var _script = WScript; //WScript | CScript
function log( message){
    
    _script.echo(message);
    
}
var APP_FOLDER                 = "bin";
var CURRENT_PATH               = _script.ScriptFullName;
var CURRENT_FOLDER             = CURRENT_PATH.slice(0, CURRENT_PATH.indexOf(_script.ScriptName));
var ROOT_FOLDER                = CURRENT_FOLDER.slice(0, CURRENT_PATH.indexOf(APP_FOLDER));




function read_all_text_file(path){
    var ForReading = 1;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try{

		var f = fso.OpenTextFile(path, ForReading);
		var content = f.AtEndOfStream ? "" : f.ReadAll();
		f.Close();
		return content;
	}
	catch(exc){
		log("cannot load file: " + path);
		return null;

	}
}

function load(modulename){
    
    var path = ROOT_FOLDER+"/libs/"+modulename+".js";
    var lib = read_all_text_file(path);
    eval(lib);
    
    
}



(function(){
	
	var args_count=_script.Arguments.Count();
	log("executing:\t"+_script.Arguments(0));
	var script = read_all_text_file(_script.Arguments(0));
    eval(script);
})();