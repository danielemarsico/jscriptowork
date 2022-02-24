/*

    ****expected folder structure*****

    ./
    ./tools/
    ./tools/in/
    ./tools/libs/
    ./tools/out/
    ./tools/templates/






*/


/************************************************
global variables
*************************************************/
var _script = WScript; //WScript | CScript


/**************************************************
helper functions
**************************************************/

function log( message){
    
    _script.echo(message);
    
}





//global variables
var APP_FOLDER                 = "tools";
var CURRENT_PATH               = _script.ScriptFullName;
var CURRENT_FOLDER             = CURRENT_PATH.slice(0, CURRENT_PATH.indexOf(_script.ScriptName));
var ROOT_FOLDER                = CURRENT_FOLDER.slice(0, CURRENT_PATH.indexOf(APP_FOLDER));

var OUTPUT_FOLDER              = "out";
var INPUT_FOLDER               = "in";

load("core");
load("system");
load("helpers");
load_properties("test.properties");


do_in_word(function(db){
	
	log('hello');
	var tmp = read_line().trim();
	
})





//basic library

function functionName(fun) {
  if(fun){
	  
	var ret = fun.toString();
	ret = ret.substr('function'.length);
	ret = ret.substr(0, ret.indexOf('('));
	if(ret == "") ret = "anonymous function";
	return ret.trim() ;
	  
  }
  else if(fun == ""){
	  
	  return "ROOT";
	  
  }
  else{
	  
	  return "ROOT";
  }
  
}

function read_all_text_file(path){
    var ForReading = 1, ForWriting = 2;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try{
		
		var f = fso.OpenTextFile(path, ForReading);
		if (f.AtEndOfStream)
			return ("");
		else
			return (f.ReadAll());
		
		f.Close();
	}
	catch(exc){
		log("workspace file doesn't exist");
		return null;
		
	}
}

function load(modulename){
    
    var path = CURRENT_FOLDER+"/libs/"+modulename+".js";
    var lib = read_all_text_file(path);
    eval(lib);
    
    
}

