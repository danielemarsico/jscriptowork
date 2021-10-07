//version 2020-10-29
//@author Daniele Marsico

//file sstem

select_file_from_folder = function (folder, message, pattern){//return selected file path
    
    var file_list           = null; 
    var selected_file_index = 0;
    
    while(true){
                   
        write_line(message);
        read_line();
                    
        file_list = list_folders(folder).filter(function(elem){
                   
            return elem.match(pattern);
       
        });
        if(file_list.length == 0){
            write_line("no files found");
        }else{
            break;
        }
    }
                
    if(file_list.length > 1){
                    
        selected_file_index = -1;
        while( selected_file_index >= file_list.length || selected_file_index < 0 ){
                    
            write_line("select file to process:");
            file_list.forEach(function(value,index){
        
                var path_parts = value.split("\\");
        
                write_line("("+(index+1)+") "+path_parts[path_parts.length-1]);
        
            });
            write(">:");
            selected_file_index = parseInt(read_line().trim(),10);
            if(isNaN(selected_file_index))
                selected_file_index = -1;
            else
                selected_file_index -= 1;
                    
                    
        }
                
    }
    var selected_file = file_list[selected_file_index]
    write_line("selected file: "+selected_file);
    return selected_file;                
    
}



read_choice_from_input = function (message,choices){//format YYYYMMDD
    
    while(true){
        
        write_line(message+" ("+choices.join("|")+") : ");       
        var input = read_line().trim().toUpperCase();
        var index = choices.indexOf(input);
        
        if(index!=-1){ 
            
            return input;
            
        }
        
    }
    
}



read_date_from_input = function (){//format YYYYMMDD
    
    while(true){
        
        write_line("enter delivery date format (YYYYMMDD): ");       
        var input = read_line().trim();
        var result = input.match(/\d{8}/);
        
        if(result!=null){
            
            //formatted_resolution_date = resolution_date = resolution_date.substring(0,4)+resolution_date.substring(5,7)+resolution_date.substring(8,10);
            var year  =  parseInt(input.substring(0,4),10);
            var month =  parseInt(input.substring(4,6),10)-1;
            var day   =  parseInt(input.substring(6,8),10);
            var delivery_date = new Date(year, month, day, 0, 0, 0, 0); 
            
            if(delivery_date){
                
                return delivery_date;
                
            }
            
        }
        
    }
    
}


select_files_from_folder = function (folder, message, pattern){//return file paths matching patterns
    
    var file_list = list_folders(folder).filter(function(elem){
                   
            return elem.match(pattern);
       
    });
    
    return file_list;                
    
}



//launch Excel and and execute to_do function in its context
do_in_excel = function (to_do){
    
    var excel = new ActiveXObject("Excel.Application");
    
    
    excel.DisplayAlerts = false;

    excel.Visible = false;
    
    try{
    
        to_do(excel);
    
    }catch(ex){
        
        log(ex.message);
        
    }
    
    excel.DisplayAlerts = true;
    
    excel.Quit();
   
    
}


//launch Access and execute to_do function in its context
do_in_access = function (to_do,database_filename){
    
    var access = new ActiveXObject("Access.Application");
    access.UserControl = false;
	access.Visible = false;
    
    
	var db = typeof database_filename !== "undefined" ? database_filename: DATABASEPATH;
	
    access.OpenCurrentDataBase(CURRENT_FOLDER+"/"+db);
    
    var db = access.CurrentDb();
  
    try{
        
        to_do(db);
        
        
    }catch(ex){
        
        log(ex.message);
        
    }

    
    
    access.CloseCurrentDatabase();
    access.Quit();
    
    
}




//launch Word and and execute to_do function in its context
do_in_word = function (to_do){
    
    var word = new ActiveXObject("Word.Application");
    
    
    word.DisplayAlerts = false;

    word.Visible = false;
    
    try{
    
        to_do(word);
    
    }catch(ex){
        
        log(ex.message);
        
    }
    
    word.DisplayAlerts = true;
    
    word.Quit();
   
    
}

alphabet = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";


//fill an Excel sheet with data form array of object, labels taken from the first element of the array
fill_sheet = function (sheet,tickets){
    
	
	if(tickets.length == 0){
		
		log("no tickets in the list");
		return;
		
	}
	
	
	
    var ticket = tickets[0];
	
	
    
    var ncolumns = 0;
    for (var k in ticket){
      
        if (ticket.hasOwnProperty(k)){
          
            ncolumns++;                
            
        } 
        
    } 
	
	
    
    
    var area = "B:"+alphabet.charAt(ncolumns);
   
    
    
    sheet.Range("A:A").Copy(sheet.Range(area));
    
    var counter = 0;
    
    for(var key in ticket){
        
        sheet.Cells(1,++counter).Value = key;
        
    }
    for(var i=0; i < tickets.length; i++){
        
        var ticket = tickets[i];
        var counter = 0;
        var row = 2+i;
        for(var key in ticket){
            var value = ticket[key];
            sheet.Cells(row,++counter).Value = value;

        }
    }
    
    
}

/*read elements from excel shhet and returns array of object, keys are taken from first row

if ncolumns is defined takes ncolumns from the sheet. if the label is empty it will use the letter of thecolumnn

if ncolumsn is not defined takes at least 100 columns


max 3000 rows
*/
read_sheet_data = function (sheet,ncolumns,params){
	
	var MAX_ROWS 		= 3000;
	var MAX_COLUMNS		=  100;
    
    //cycling on first row to get labels
    var isempty = false;
    var counter = 1;
    var labels = [];      
    
    
    var start_row = 1;
    if(params && params.start_row){
        
        start_row = params.start_row;
    }
    
    
	var check_empty = function(){return true}	
	
	if(ncolumns){
		
		//log('reading data until:' + ncolumns)
		
		check_empty = function(x){
			
			return (function(sheet,row){
			
				for(var i= 1; i<= x; i++){
				
					var v = sheet.Cells(row,i).Value
                    //log(v)
				
					if(typeof v != 'undefined' && v.toString().trim() != ""  ){
					
						return false;
					
					}
				}
			
				return true;
			});
			
		}(ncolumns)
		
		while(counter <= ncolumns){
			
			var cell = sheet.Cells(start_row,counter);
			var value = cell.Value;
			//log(counter+" : "+value)
			if(!value || value.trim() == "" || typeof(value) == 'undefined'){
				
				value =  alphabet.charAt(counter)
				//log('new value: '+value)
			}
			labels.push(value.trim());
			counter = counter +1;
		}
	}
	else{
		
		
		check_empty = function(sheet,row){
			
			var v = sheet.Cells(row,1).Value;
			
			return !v || v.trim() == "" || typeof(v)== 'undefined';
		}
		
		while(!isempty && counter < MAX_COLUMNS){
			var cell = sheet.Cells(start_row,counter);
			var value = cell.Value;
		   
			if(!value || value.trim() == "" ){
				isempty = true;
				break;
			}else{
			   labels.push(value.trim());
			}
			counter = counter +1;
		}
			
	}
    
    log("n° of columns: "+labels.length);
	log(labels);
                        
    var row_counter = start_row+1;
    
    
    
    
    //getting all problems
    isempty = false;
	
    var problems = [];
            
    while(!isempty && row_counter < MAX_ROWS){
       
        var problem = {};
        //log(row_counter+")");
		
		if(check_empty(sheet,row_counter)) { //end of rows
			
			//log('empty list')
			break;
        }
	
		for(var i=1; i<= labels.length; i++){
            
            var value = sheet.Cells(row_counter,i).Value;
			//log(labels[i-1]+":"+value+" - "+typeof value);
			
			if(typeof value == "number"){
				
				value = value.toString()
				
			}else if(typeof value == "date"){
                
                
               try{
                    
                    var d = new Date("" +value);
                    value = "" + d.getFullYear()+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+d.getDate()).slice(-2);
                    
                }catch(ex){
                    
                    log(ex.message);
                    value="-";
                    
                }
               
                
            }
            else if(typeof value == "undefined" || value==null || !value){
                    value= "";
            
            }
			
			problem[labels[i-1]] = value.trim();
            
        }
        
        problems.push(problem);
        
        row_counter= row_counter +1;
    }
    
    
    return problems;
    
};

//save text file at a filepath replacing excel extension with txt extension in the filename
write_report_to_file = function (report_data,filepath){
    
    var d = new Date();
    var report = "";
    
    if(filepath){
        
        var current_date = "-"+d.getTime();
        
        var tokens = filepath.split("\\");
        filename_radix = (tokens[tokens.length-1]).replace(/\.xls(x)?/,"");
        
        report = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+filename_radix+ current_date + ".txt";
    }else{
        
        var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
    
        report = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+"consistency_analysis_report-"+ current_date + ".txt";
        
    }
    log("saving report to : " + report);
    write_text_to_file(report_data,report);
    
}



get_current_date_as_excel_text = function(){
	
	var d = new Date();
	return  "'"+d.getFullYear()+[d.getMonth()+1,d.getDate(),d.getHours(),d.getMinutes(),d.getSeconds()].map(function(x){return ("0"+x).slice(-2)}).join("");
	
	
}




//execute a saved query in Access
execute_query = function (saved_query,db){
    
    var dbOpenDynamic       = 16;
	var dbOpenDynaset       = 2;
    var dbOpenForwardOnly   = 8;
    var dbOpenSnapshot      = 4;
    var dbOpenTable         = 1;
    
    
    
    var rs = db.OpenRecordset(saved_query,dbOpenDynaset);
    var counter = 0;
    
    var records = [];
    while(!rs.EOF){
        
        var record = {};
        
        for(var  i= 0; i< rs.Fields.Count; i++){
            
			
			var v = rs.Fields(i).Value;
			
			if(v!=null){
				
				v = v.toString().trim();
				
			}else{
				
				v= "";
				
			}
			
			
			
            record[rs.Fields(i).Name]=v;
			
        }
        records.push(record);
        rs.MoveNext();
        
    }
    
	
	
	
    return records;
}