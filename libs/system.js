

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