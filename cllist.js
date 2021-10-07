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
	
	//var caller = functionName(arguments.callee.caller);
    
    //_script.echo("<"+caller+">: "+message);
	_script.echo("<DEBUG>: "+message);
    
}





//global variables
var APP_FOLDER                 = "tools";
var CURRENT_PATH               = _script.ScriptFullName;
var CURRENT_FOLDER             = CURRENT_PATH.slice(0, CURRENT_PATH.indexOf(_script.ScriptName));
var ROOT_FOLDER                = CURRENT_FOLDER.slice(0, CURRENT_PATH.indexOf(APP_FOLDER));
var TEMPLATE_FOLDER            = "templates";
var OUTPUT_FOLDER              = "out";
var INPUT_FOLDER               = "in";

load("core");
load("system");
load("helpers");
load_properties("test.properties");


function create_consolidatedlist_workbook_factory(problems,incidents){
    
	
    return function(excel){
        
        var book = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+CONSOLIDATED_LIST_TEMPLATE);
		var new_list = "";
		
        try{
            
            //populating problems
            book.Sheets(1).Activate();
            fill_sheet(book.Sheets(1),problems)
            //populating incidents
            book.Sheets(2).Activate();
            fill_sheet(book.Sheets(2),incidents)
            book.Sheets(1).Activate();
			
			var d = new Date();
        
			//CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
			var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
			new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+CONSOLIDATED_LIST_TEMPLATE.replace("YYMMDD",current_date);
        
			log("saving new consolidated list: " + new_list);
        
			book.SaveCopyAs(new_list);
			
			
        
		}catch(ex){
            
             log(ex.message);
        
		}finally{
			
			try{
				book.Close();
				
			}catch(ex){
				
				log("***ERROR*** - error while trying to close the workbook");
				
			}
			
		}
		
		return new_list;
		
		
            
    }
    
}

function add_notes_factory(tickets, filename,folder){
    
    return function(excel){
        
        var book = excel.Workbooks.Open(folder+filename);
        
        try{
            
            //getting problem/incident key
            var ticket = tickets[0];
            var ticketid_lable = "";
            var pos = 0;
            for (var k in ticket){
                if (pos==0){
                    ticketid_lable = k;
                    break;
                }
            }
            var notes_label = NOTES_COLUMN_NAME;
            
            book.Sheets(1).Activate();
            
            var _notes = read_sheet_data(book.Sheets(1));
            
            for (var i = 0; i < tickets.length; i++){
                
                for(var j =0; j < _notes.length; j++){
                    
                    var id_sx = tickets[i][ticketid_lable];
                    var id_dx = _notes[j][ticketid_lable];
                    
                    if(id_sx==id_dx){
                        tickets[i][notes_label] = _notes[j][notes_label];
                        _notes.splice(j,1);
                        break;                          
                    }
                }
                
            }
            
            
            
        }catch(ex){
            
            log(ex.message);
            
        }
        
        book.Close();
    }
    
   
}


function export_consolidatedlist_to_json(problems,incidents,string_date){
    
    var d = new Date();
    //CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
    if(typeof(string_date) == 'undefined'){
            
        string_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2);
        
    }
    
    
    
    var current_date = string_date+"-"+d.getTime();
    var new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+CONSOLIDATED_LIST_TEMPLATE.replace("YYMMDD",current_date).replace(".xlsx",".json");
    
    var cl = {};
    cl["extraction_date"] = string_date;
    cl["problems"] = problems;
    cl["incidents"] = incidents;
    
    var cl_json = Ext.encode(cl);
    
    write_text_to_file(cl_json,new_list);
    
}

function create_consolidated_list_from_access(){
    
	execute_omgchanges();
	
	//build consolidated list
	var produced_cl = null;
	
    do_in_access(function(db){ 
        
		
        log("loading data from query: "+PROBLEMS_SAVED_QUERY)
        var tmp_problems  = execute_query(PROBLEMS_SAVED_QUERY,db);
        log("data loaded, n° of problems: "+tmp_problems.length);
		
		
		//normalize release note
		var problems = tmp_problems.map(function(elem){
			
			var release_number = elem["Release number"].trim();
			
			//var result = /^R\s?(\d)((?:\.\d)*)\s?(\w*)$/.exec(release_number);
			var result = /^R\s?(\d)((?:\.\d)*)\s?(\w*)$/.exec(release_number);
			
			
			if(result && result.length > 1){
				//log(result);
				
				var rn = "R "+result[1];
				
				for(var i= 2 ; i < result.length -1;i++){
					
					rn = rn + result[i];
					
				}
				
				var lastpart = result[result.length-1].trim()
				if( lastpart != ""){
					
					rn+=" "+lastpart;
					
				}
				
				
				//log(rn);
				
				elem["Release number"] = rn ;
				
				
			}else{
				
				elem["Release number"] =  release_number;
				
			}
			
			
			
			return elem;
		});
		
		
		
		
        
        log("loading data from query: "+INCIDENTS_SAVED_QUERY)
        var incidents = execute_query(INCIDENTS_SAVED_QUERY,db);
        log("data loaded, n° of incidents: "+incidents.length);
        
        log("running consistency analysis...")
        var consistency_analysis_report = run_consistency_analysis(problems);
        log("consistency analysis done.");
       
        write_report_to_file(consistency_analysis_report);
       
        var create_consolidatedlist_workbook = create_consolidatedlist_workbook_factory(problems,incidents);
        var add_notes_to_problems  = add_notes_factory(problems,PROBLEMS_NOTES_FILE,WORKING_DIRECTORY);
        var add_notes_to_incidents = add_notes_factory(incidents,INCIDENTS_NOTES_FILE,WORKING_DIRECTORY);
        
        do_in_excel(function(excel){
            
            log("checking total number of open problems");
            (function(excel){
                
                var book = excel.Workbooks.Open(WORKING_DIRECTORY+PROBLEMS_NOTES_FILE);
        
                try{
            
                    book.Sheets(1).Activate();
                    
                    var excluded_statuses = ["Cancelled","Closed",""];
            
                    var filtered = read_sheet_data(book.Sheets(1)).filter(function(elem){
                        
                        return excluded_statuses.indexOf(elem["Status*"]) == -1 ;
                    });
                    log("n° of tickets in "+PROBLEMS_NOTES_FILE+" :"+filtered.length);
                    log("n° of tickets in "+PROBLEMS_SAVED_QUERY+" :"+problems.length);
                    if(filtered.length != problems.length){
                        
                        log("***********  extraction not valid, tickets number doesn't match     ***********");
                        
                    }
                    
            
                }catch(ex){
            
                    log(ex.message);
            
                }
        
                book.Close();
          
            })(excel);
            
            
            log("adding notes to problems..");
            add_notes_to_problems(excel);
            log("notes added");
			
			//split BOD
			problems = problems.map(function(elem){
				
				//var notes = elem["Notes"].replace(/\n/g, ' ');
				var notes = elem["Notes"];
				//log(notes);
				
				//var result = /^R\s?(\d)((?:\.\d)*)\s?(\w*)$/.exec(release_number);
				
				var result = /BOD(:)?([\s\S]+)/m.exec(notes);
				
				if(result && result.length > 1){
				
					//log(elem[PROBLEM_ID_COLUMN_NAME] +"=> "+ result[result.length-1] );
					elem["BOD"]=result[result.length-1].trim();
				
				}
				
				return elem;
				
				
				
				
			})
            
            log("adding notes to incidents...");
            add_notes_to_incidents(excel);
            log("notes added");
            
            log("exporting data in json format");
            export_consolidatedlist_to_json(problems,incidents);
            log("data exported");
            
            
            log("creating consolidated list...")
            produced_cl = create_consolidatedlist_workbook(excel);
            
            log("process completed.")
            
        });
        
        
        
    });
	
	
	return produced_cl;

}

function create_breakdownperncb_list_from_access(consolidated_list_file){
    
    do_in_access(function(db){ 

        log("loading data from query: "+BREAKDOWN_PROBLEMS_SAVED_QUERY)
        var problems  = execute_query(BREAKDOWN_PROBLEMS_SAVED_QUERY,db);
        log("data loaded, n° of problems: "+problems.length);
        
        log("loading data from query: "+BREAKDOWN_INCIDENTS_SAVED_QUERY)
        var incidents  = execute_query(BREAKDOWN_INCIDENTS_SAVED_QUERY,db);
        log("data loaded, n° of incidents: "+incidents.length);
        
        do_in_excel(function(excel){
            
            if( consolidated_list_file!= null){
                var cl_problems = [];
                try{
                    log("loading problems and incidents from: "+consolidated_list_file)
                    var book = excel.Workbooks.Open(consolidated_list_file);
                    book.Sheets(1).Activate();
                    cl_problems  = read_sheet_data(book.Sheets(1));
                    book.Sheets(2).Activate();
                    cl_incidents  = read_sheet_data(book.Sheets(2));
                    
                    problems = problems.filter(function(elem){
                        
                        var result = cl_problems.find(function(r){
                            
                            return r[PROBLEM_ID_COLUMN_NAME] == elem[PROBLEM_ID_COLUMN_NAME];
                            
                        })
                        if(!result){
                            
                            log(elem[PROBLEM_ID_COLUMN_NAME]+" removed!");
                            return false;
                            
                        }
                        else{
                            //here updates data taking values from the input consolidated list
                            var columns = [ "Problem/Defect",
                                            "Summary*",
                                            "Priority*",
                                            "Environment",
                                            "Submit Date",
                                            "Category",
                                            "Affected Functionality",
                                            "System Entity",
                                            "Status",
                                            "External Reference",
                                            "Resolution Date",
                                            "Release number",
                                            "Related Problems",
                                            "Additional information",
                                            "Updated",
                                            "Notes"];
                                            
                            columns.forEach(function(x){
                                
                                elem[x] = result[x];
                                
                                
                            });
                            
                            
                            
                            return true;
                        }
                        
                        
                    });
                    
                    incidents = incidents.filter(function(elem){
                        
                        var result = cl_incidents.find(function(r){
                            
                            return r[INCIDENT_ID_COLUMN_NAME] == elem[INCIDENT_ID_COLUMN_NAME];
                            
                        })
                        if(!result){
                            
                            log(elem[INCIDENT_ID_COLUMN_NAME]+" removed!");
                            return false;
                        }else{
                            
                            return true;
                        }
                        
                        
                    });
                   
                    book.Close();
                }catch(ex){
            
                    log(ex.message);
            
                }
            }
            
            //loading template
            var book = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+BREAKDOWN_TEMPLATE);
    
            try{
            
                //populating problems
                book.Sheets(2).Activate();
                fill_sheet(book.Sheets(2),problems)
                //populating incidents
                book.Sheets(4).Activate();
                fill_sheet(book.Sheets(4),incidents)
                
                book.Sheets(3).Activate();
                
                var tabella_pivot_incident = book.Sheets(3).PivotTables("tabella_pivot_incident");
                tabella_pivot_incident.RefreshTable();
                
                book.Sheets(1).Activate();
                
                var tabella_pivot_problem = book.Sheets(1).PivotTables("tabella_pivot_problem");
                tabella_pivot_problem.RefreshTable();
            
            
            }catch(ex){
            
                log(ex.message);
            
            }
        
            var d = new Date();
            
            //CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
            var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
            var new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+BREAKDOWN_TEMPLATE.replace("YYMMDD",current_date);
            
            log("saving new breakdown list: " + new_list);
            
            try{
                    book.SaveCopyAs(new_list);
            }catch(ex){
                
                 log(ex.message);
            }
            
            
            book.Close();
            
        
        });
        
        log("process completed.")

    });  
   
}


function create_operday_list_from_access(consolidated_list_file){
    
    do_in_access(function(db){ 

        log("loading data from query: "+OPERDAY_PROBLEMS_SAVED_QUERY)
        var problems  = execute_query(OPERDAY_PROBLEMS_SAVED_QUERY,db);
        log("data loaded, n° of problems: "+problems.length);
        
        do_in_excel(function(excel){
            
            if( consolidated_list_file!= null){
                var cl_problems = [];
                try{
                    log("loading problems from: "+consolidated_list_file)
                    var book = excel.Workbooks.Open(consolidated_list_file);
                    book.Sheets(1).Activate();
                    cl_problems  = read_sheet_data(book.Sheets(1));
                    
                    
                    problems = problems.map(function(elem){
                        
                        var result = cl_problems.find(function(r){
                            
                            return r[PROBLEM_ID_COLUMN_NAME] == elem[PROBLEM_ID_COLUMN_NAME];
                            
                        })
                        if(!result){
                            
                            log(elem[PROBLEM_ID_COLUMN_NAME]+" not found!");
                           
                            
                        }
                        else{
                            //here updates data taking values from the input consolidated list
                            var columns = [ "Problem/Defect",
                                            "Summary*",
                                            "Priority*",
                                            "Environment",
                                            "Submit Date",
                                            "Category",
                                            "Affected Functionality",
                                            "System Entity",
                                            "Status",
                                            "External Reference",
                                            "Resolution Date",
                                            "Release number",
                                            "Related Problems",
                                            "Additional information",
                                            "Updated",
                                            "Notes"];
                                            
                            columns.forEach(function(x){
                                
                                if(x in elem){
                                    
                                    elem[x] = result[x];   
                                
                                }
                                
                            });
                        }
                        
                        return elem;
                        
                    });
                    
                    book.Close();
                }catch(ex){
            
                    log(ex.message);
            
                }
            }
            
            //loading template
            var book = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+OPERDAY_TEMPLATE);
    
            try{
            
                //populating problems
                book.Sheets(1).Activate();
                fill_sheet(book.Sheets(1),problems)
               
            
            
            }catch(ex){
            
                log(ex.message);
            
            }
        
            var d = new Date();
            
            //CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
            var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
            var new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+OPERDAY_TEMPLATE.replace("YYMMDD",current_date);
            
            log("saving new breakdown list: " + new_list);
            
            try{
                    book.SaveCopyAs(new_list);
            }catch(ex){
                
                 log(ex.message);
            }
            
            
            book.Close();
            
        
        });
        
        log("process completed.")

    });  
   
    
    
}

function generate_weekly_report_data(problems, date){
    
    var open_tickets                    = {};
    var pending_tickets                 = {};
    var new_tickets                     = {};
    
    open_tickets.data                   = [];
    pending_tickets.data                = [];    
    new_tickets.data                    = [];  
    
    open_tickets.label = 'Evolution of open tickets';
    pending_tickets.label = 'Evolution of pending tickets';
    new_tickets.label = 'New tickets identified per week';
    
    open_tickets.rows_label = [ 'Open release defects',
                                'Open known production problems',
                                'to be defined',
                                'Total'
                              ];
    pending_tickets.rows_label = [ 'Pending release defects',
                                    'Pending known production problems',
                                    'To be defined',
                                    'Total'
                                ];
    new_tickets.rows_label = [  'New defects identified per week',
                                'New known problems indentified per week',
                                'Tickets not yet categorized indentified per week',
                                'Total'
                                ];
    
    
    var date_label = "'"+date.substr(4,2)+"/"+date.substr(2,2)+"/"+ "20"+date.substr(0,2);    
    
    open_tickets.data.push(date_label);
    pending_tickets.data.push(date_label);
    new_tickets.data.push(date_label); 
    
    
    {
        
        var reduce_fnc_create = function(case_type){
                    
            return function(previousValue, currentValue, currentIndex, array) {
                
                if(case_type == null){
                    return previousValue+1;  
                }else if(case_type == currentValue['Problem/Defect'].trim()){
                     return previousValue+1;
                }else{
                     return previousValue;
                }
            
            }   
            
        }

        
        var defects       = problems.reduce(reduce_fnc_create('Defect'),0);
        var open_problems = problems.reduce(reduce_fnc_create('Problem'),0);
        var to_be_defined = problems.reduce(reduce_fnc_create('To be defined'),0);
        var total         = problems.reduce(reduce_fnc_create(null),0);
        open_tickets.data.push(defects);
        open_tickets.data.push(open_problems);
        open_tickets.data.push(to_be_defined);
        open_tickets.data.push(total);
    
    }
    
    {
        
        var reduce_fnc_create = function(case_type){
                
            return function(previousValue, currentValue, currentIndex, array) {
                
                if( ['Waiting for delivery in UTEST/PROD', 'Pending Documentation Enhancement'].indexOf(currentValue['Status'].trim()) != -1){
                    
                    return previousValue;
                }
                
                
                if(case_type == null){
                    return previousValue+1;  
                }else if(case_type == currentValue['Problem/Defect'].trim()){
                     return previousValue+1;
                }else{
                     return previousValue;
                }
            
            }   
            
        }

        var defects                 = problems.reduce(reduce_fnc_create('Defect'),0);
        var pending_problems        = problems.reduce(reduce_fnc_create('Problem'),0);
        var to_be_defined           = problems.reduce(reduce_fnc_create('To be defined'),0);
        var total                   = problems.reduce(reduce_fnc_create(null),0);
        pending_tickets.data.push(defects);
        pending_tickets.data.push(pending_problems);
        pending_tickets.data.push(to_be_defined);
        pending_tickets.data.push(total);
        
        
    }
    
    {
        var reduce_fnc_create = function(case_type){
                
            return function(previousValue, currentValue, currentIndex, array) {
                
                if( ['*NEW*'].indexOf(currentValue['Updated'].trim()) == -1){
                    
                    return previousValue;
                }
                
                
                if(case_type == null){
                    return previousValue+1;  
                }else if(case_type == currentValue['Problem/Defect'].trim()){
                     return previousValue+1;
                }else{
                     return previousValue;
                }
            
            }   
            
        }

        var defects             = problems.reduce(reduce_fnc_create('Defect'),0);
        var new_problems        = problems.reduce(reduce_fnc_create('Problem'),0);
        var to_be_defined       = problems.reduce(reduce_fnc_create('To be defined'),0);
        var total               = problems.reduce(reduce_fnc_create(null),0);
        new_tickets.data.push(defects);
        new_tickets.data.push(new_problems);
        new_tickets.data.push(to_be_defined);
        new_tickets.data.push(total);
        
        
    }
    
    //grouping per resolution date
    
    var d = new Date();
            
    var current_date = ""+d.getFullYear()+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+(d.getDate())).slice(-2);
                
    log('current date: '+ current_date);
    
	var rd_problems = problems.filter(function(elem){
		
		return (elem['Status'] == 'Work in Progress' || elem['Status'] == 'Pending Client Action Required') && (elem['Release number'] == PREVIOUS_RELEASE || elem["Problem/Defect"] == "Defect" ) ;
		
	}).map(function(x){
		
		
		if(x["Resolution Date"] < current_date){
			
				x["Resolution Date"] = "";
			
		}
		
		return x;
		
	});
	
	log('there are still '+rd_problems.length+" working in progress or pending client action required");
	
	
	var resolution_dates_group = rd_problems.reduce(function(acc,e){
                                         
			var rd = e["Resolution Date"];
			
			if(rd in acc){
				acc[rd].push(e);
			}else{
				acc[rd] = [e];
			}
			return acc
		}, {});
	
	var resolution_dates  =[]

    for(var rd in resolution_dates_group){
		
		resolution_dates.push(rd);
		
	}
	
	resolution_dates = resolution_dates.sort();
	
	if(resolution_dates[0] == ""){
		
		resolution_dates = resolution_dates.slice(1).concat(resolution_dates[0]);
		
	}
	
	resolution_dates = resolution_dates.map(function(rd){
		
		return ["'"+rd,
				resolution_dates_group[rd].length,
				0];
		
	});
	
	
	/*resolution_dates =resolution_dates.map(function(elem){
		
		if(elem==""){
			
			return [elem,""];
		}else{
			
			//return [elem,parse_date(elem,"DD/MM/YYYY")];
			return [elem,""];
		}
		
		
	});*/

    
     
    return   [open_tickets,pending_tickets,new_tickets,resolution_dates];
}

function export_release_note(env){
	
	var pattern = /.+[t2s|T2S][_\s][Rr]elease[_\s][Nn]ote[_\s][vV](\d{2})[\._\s](\d{2})[\._\s](\d{2})[\._\s](\d{2}).*\.docx$/;
	var message = "copy release notes in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
	var release_notes_files = select_files_from_folder(WORKING_DIRECTORY,message,pattern); 
				
	log("importing release notes into access:  "+release_notes_files.length+" ******"); 
					
	var records_cr 			= [];
	var records_problems 	= [];
	var records_defects 	= [];
					
	release_notes_files.forEach(function(release_note_file){
		
		var rn_records_cr			= [] 
		var rn_records_problems 	= [];
		var rn_records_defects 		= [];
		 
		log("importing: "+release_note_file); 
		var results = pattern.exec(release_note_file);
		var version = results.slice(1).join(".");
		log("version RN v"+version);
		
		var eac_dd 		= "";
		var mig1_dd 	= "";
		var mig2_dd 	= "";
		var utest_dd 	= "";
		var prod_dd 	= "";
		
		
		var current_date = get_current_date_as_excel_text();
	
		do_in_word(function(word){
			
			var doc = word.Documents.Open(release_note_file);
			doc.Activate();
			
			//extracting delivery date
			var pars_count = doc.Paragraphs.Count;
			for (var i= 1; i <= pars_count; i++ ){
				
				var par =  doc.Paragraphs(i);
				
				var text = par.Range.Text.slice(0,-1).trim().toUpperCase();
				//log(text);
				
				
				
				var results = /INTEROPERABILITY\s+ENVIRONMENT\s+ON\s+(\d{2})\/(\d{2})\/(\d{4})/.exec(text);
				
				if(results && results.length>1){
					
					eac_dd = results[1]+"/"+results[2]+"/"+results[3];
					
					log("EAC DD: "+ eac_dd);
				}
					
				results = /MIGRATION\s+ENVIRONMENT\s+ON\s+(\d{2})\/(\d{2})\/(\d{4})/.exec(text);
			
				if(results && results.length>1){
				
					mig1_dd = results[1]+"/"+results[2]+"/"+results[3];
				
					log("MIG1 DD: "+ mig1_dd);
				}
					
						
				results = /COMMUNITY\s+ENVIRONMENT\s+ON\s+(\d{2})\/(\d{2})\/(\d{4})/.exec(text);
		
				if(results && results.length>1){
			
					mig2_dd = results[1]+"/"+results[2]+"/"+results[3];
			
					log("MIG2 DD: "+ mig2_dd);
				}
				
				
				results = /PRE\s*PRODUCTION\s+ENVIRONMENT\s+ON\s+(\d{2})\/(\d{2})\/(\d{4})/.exec(text);
				if(results && results.length>1){
			
					utest_dd = results[1]+"/"+results[2]+"/"+results[3];
			
					log("UTEST DD: "+ utest_dd);
				}
				
				
				results = /PRODUCTION\s+ENVIRONMENT\s+ON\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(text);
				if(results && results.length>1){
			
					prod_dd = results[1]+"/"+results[2]+"/"+results[3];
			
					log("PROD DD: "+ prod_dd);
				}
				
				
						
						
					
				
				
			}
			
			var tables_length = doc.Tables.Count;
			log("tables:"+tables_length);
			
			
			//selecting the right table
			var tables = {};
			tables.cr = null;
			tables.problems = null;
			tables.defects = null;
			
			var max= 0;
			
			if(env=="UTEST" || env=="PROD"){
				
				
				
				
				//4CB ticket	Case Type	Customer Trouble case id	L2 reference	Summary	RN	Component	Retest status

				for(var i= 1; i <= tables_length; i++){
					table = doc.Tables(i);
					
					var main_key = table.Cell(1,1).Range.Text.slice(0,-1).trim();
					log(">>>>>>>>>>>>> "+main_key);
					if(main_key == "CR Reference" && tables.cr == null){
						
						tables.cr = table;
						
					}else if( main_key == "4CB ticket" ){
						
						var case_type = table.Cell(2,2).Range.Text.slice(0,-1).trim();
						
						if(case_type == "Problem" ){
							
							if(tables.problems == null) {
								
								tables.problems = table;
							
							}
							
						}else{
							
							if (tables.defects == null){
								
								tables.defects = table;
								
							}
							
							
							
						}
						
						
					}
					
					
				
					
				}
				
				
				var table = tables.cr;
				log("loading crs");
				
				if(table != null){
					
					var cols_length = table.Columns.Count;
					var rows_length = table.Rows.Count;
				
					log("n° of tickets:"+(rows_length-1)) ;
					
					
					//CR Reference	| Title	| Delivery to Interoperability environment|	Link to updated schema/message documentation
					for(var i = 2, counter= 1; i <= rows_length; i++, counter++){
						
						var record = {};
						record["CR Reference"] 									= table.Cell(i,1).Range.Text.slice(0,-1).trim();
						record["Title"] 										= table.Cell(i,2).Range.Text.slice(0,-1).trim();
						record["Delivery to Interoperability environment"] 		= table.Cell(i,3).Range.Text.slice(0,-1).trim();
						record["Link to updated schema/message documentation"] 	= table.Cell(i,4).Range.Text.slice(0,-1).trim();
						
						record["Release Note"] = "V"+version;
						
						record["Filename"]			 = release_note_file;
						record["Export Date"] 	 	 = current_date;
						record["ENV"] 	 			 = env;
						if(env=='UTEST')
							record["UTEST DD"] 	 	 = "'"+utest_dd;
						else if(env=='PROD')
							record["PROD DD"] 	 	 	 = "'"+prod_dd;
						
						
						
						
						rn_records_cr.push(record);
				
					}
					
				}
				
				
				var table = tables.problems;
				log("loading problems");
				if(table != null){
					var cols_length = table.Columns.Count;
					var rows_length = table.Rows.Count;
				
					log("n° of tickets:"+(rows_length-1)) ;
					
					for(var i = 2, counter= 1; i <= rows_length; i++, counter++){
						
						var record = {};
						record["Release Note"] = "V"+version;
						record["4CB ticket"] 				= table.Cell(i,1).Range.Text.slice(0,-1).trim();
						log(counter +") "+record["4CB ticket"]);
						record["Case Type"] 				= table.Cell(i,2).Range.Text.slice(0,-1).trim();
						record["Customer Trouble case id"] 	= table.Cell(i,3).Range.Text.slice(0,-1).trim();
						record["L2 reference"] 				= table.Cell(i,4).Range.Text.slice(0,-1).trim();
						record["Summary"] 					= table.Cell(i,5).Range.Text.slice(0,-1).trim();
						record["RN"] 						= table.Cell(i,6).Range.Text.slice(0,-1).trim();
						record["Component"] 				= table.Cell(i,7).Range.Text.slice(0,-1).trim();
						record["Retest status"] 			= table.Cell(i,8).Range.Text.slice(0,-1).trim();
						record["Filename"]			 = release_note_file;
						record["Export Date"] 	 	 = current_date;
						record["ENV"] 	 			 = env;
						if(env=='UTEST')
							record["UTEST DD"] 	 	 = "'"+utest_dd;
						else if(env=='PROD')
							record["PROD DD"] 	 	 	 = "'"+prod_dd;
						
						rn_records_problems.push(record);
				
					}
					
					
					
				}
				
				var table = tables.defects;
				log("loading defects");
				if(table != null){
					var cols_length = table.Columns.Count;
					var rows_length = table.Rows.Count;
				
					log("n° of tickets:"+(rows_length-1)) ;
					
					for(var i = 2, counter= 1; i <= rows_length; i++, counter++){
						
						var record = {};
						record["Release Note"] = "V"+version;
						record["4CB ticket"] 				= table.Cell(i,1).Range.Text.slice(0,-1).trim();
						log(counter +") "+record["4CB ticket"]);
						record["Case Type"] 				= table.Cell(i,2).Range.Text.slice(0,-1).trim();
						record["Customer Trouble case id"] 	= table.Cell(i,3).Range.Text.slice(0,-1).trim();
						record["L2 reference"] 				= table.Cell(i,4).Range.Text.slice(0,-1).trim();
						record["Summary"] 					= table.Cell(i,5).Range.Text.slice(0,-1).trim();
						record["RN"] 						= table.Cell(i,6).Range.Text.slice(0,-1).trim();
						record["Component"] 				= table.Cell(i,7).Range.Text.slice(0,-1).trim();
						record["Retest status"] 			= table.Cell(i,8).Range.Text.slice(0,-1).trim();
						record["Filename"]			 = release_note_file;
						record["Export Date"] 	 	 = current_date;
						record["ENV"] 	 			 = env;
						if(env=='UTEST')
							record["UTEST DD"] 	 	 = "'"+utest_dd;
						else if(env=='PROD')
							record["PROD DD"] 	 	 	 = "'"+prod_dd;
						
						rn_records_defects.push(record);
				
					}
					
					
					
				}
				
				
				
				
				
				
			}
			else if(env == "EAC"){
				
				for(var i= 1; i <= tables_length; i++){
					table = doc.Tables(i);
					
					var main_key = table.Cell(1,1).Range.Text.slice(0,-1).trim();
					log(">>>>>>>>>>>>> "+main_key);
					if(main_key == "CR Reference" && tables.cr == null){
						
						tables.cr = table;
						
					}else if( main_key == "4CB ticket" ){
						
						var case_type = table.Cell(2,2).Range.Text.slice(0,-1).trim();
						
						if(case_type == "Problem" ){
							
							if(tables.problems == null) {
								
								tables.problems = table;
							
							}
							
						}else{
							
							if (tables.defects == null){
								
								tables.defects = table;
								
							}
							
							
							
						}
						
						
					}
					
					
				
					
				}
				
				
				var table = tables.cr;
				
				if(table != null){
					
					var cols_length = table.Columns.Count;
					var rows_length = table.Rows.Count;
				
					log("n° of tickets:"+(rows_length-1)) ;
					
					
					//CR Reference	| Title	| Delivery to Interoperability environment|	Link to updated schema/message documentation
					for(var i = 2, counter= 1; i <= rows_length; i++, counter++){
						
						var record = {};
						record["CR Reference"] 									= table.Cell(i,1).Range.Text.slice(0,-1).trim();
						record["Title"] 										= table.Cell(i,2).Range.Text.slice(0,-1).trim();
						record["Delivery to Interoperability environment"] 		= table.Cell(i,3).Range.Text.slice(0,-1).trim();
						record["Link to updated schema/message documentation"] 	= table.Cell(i,4).Range.Text.slice(0,-1).trim();
						
						record["Release Note"] = "V"+version;
						
						record["Filename"]			 = release_note_file;
						record["Export Date"] 	 	 = current_date;
						record["ENV"] 	 			 = env;
						record["EAC DD"] 	 		 = "'"+eac_dd;
						record["MIG1 DD"] 	 		 = "'"+mig1_dd;
						record["MIG2 DD"] 	 		 = "'"+mig2_dd;
						
						
						
						
						rn_records_cr.push(record);
				
					}
					
				}
				
				table = tables.problems;
				
				if(table != null){
					
					var cols_length = table.Columns.Count;
					var rows_length = table.Rows.Count;
				
					log("n° of tickets:"+(rows_length-1)) ;
					
					
					//4CB ticket	| Case Type	 | Customer Trouble case id	| L2 reference	| Description

					for(var i = 2, counter= 1; i <= rows_length; i++, counter++){
						
						var record = {};
						record["4CB ticket"] 									= table.Cell(i,1).Range.Text.slice(0,-1).trim();
						record["Case Type"] 									= table.Cell(i,2).Range.Text.slice(0,-1).trim();
						var tmp = table.Cell(i,3).Range.Text;
						if(tmp.length>0){
							
							tmp = tmp.slice(0,-1).trim();
							
						}
						record["Customer Trouble case id"] =  tmp;
						record["L2 reference"] =   table.Cell(i,4).Range.Text.slice(0,-1).trim();
						record["Description"] =   table.Cell(i,5).Range.Text.slice(0,-1).trim();
						record["Release Note"] = "V"+version;
						
						record["Filename"]			 = release_note_file;
						record["Export Date"] 	 	 = current_date;
						record["ENV"] 	 			 = env;
						record["EAC DD"] 	 		 = "'"+eac_dd;
						record["MIG1 DD"] 	 		 = "'"+mig1_dd;
						record["MIG2 DD"] 	 		 = "'"+mig2_dd;
						
						rn_records_problems.push(record);
				
					}
					
				}
				
				table = tables.defects;
				
				if(table != null){
					
					var cols_length = table.Columns.Count;
					var rows_length = table.Rows.Count;
				
					log("n° of tickets:"+(rows_length-1)) ;
					
					
					//4CB ticket	| Case Type	 | Customer Trouble case id	| L2 reference	| Description

					for(var i = 2, counter= 1; i <= rows_length; i++, counter++){
						
						var record = {};
						record["4CB ticket"] 									= table.Cell(i,1).Range.Text.slice(0,-1).trim();
						record["Case Type"] 									= table.Cell(i,2).Range.Text.slice(0,-1).trim();
						var tmp = table.Cell(i,3).Range.Text;
						if(tmp.length>0){
							
							tmp = tmp.slice(0,-1).trim();
							
						}
						record["Customer Trouble case id"] =  tmp;
						record["L2 reference"] =   table.Cell(i,4).Range.Text.slice(0,-1).trim();
						record["Description"] =   table.Cell(i,5).Range.Text.slice(0,-1).trim();
						record["Release Note"] = "V"+version;
						
						record["Filename"]			 = release_note_file;
						record["Export Date"] 	 	 = current_date;
						record["ENV"] 	 			 = env;
						record["EAC DD"] 	 		 = "'"+eac_dd;
						record["MIG1 DD"] 	 		 = "'"+mig1_dd;
						record["MIG2 DD"] 	 		 = "'"+mig2_dd;
						
						rn_records_defects.push(record);
				
					}
					
					
					
				}




		}//end IF EAC
			
			doc.Close();
			
		});
			
		
		records_cr 			= records_cr.concat(rn_records_cr);
		records_problems 	= records_problems.concat(rn_records_problems);
		records_defects 	= records_defects.concat(rn_records_defects);	
	
		
		
	
	
	});
	
	var produced_file = null;
	
	do_in_excel(function(excel){
		var template = (["EAC","MIG1","MIG2"].indexOf(env) != -1) ? TMS_STATUS_OF_DEFECT_IN_RELEASENOTE_EAC_TEMPLATE : TMS_STATUS_OF_DEFECT_IN_RELEASENOTE_UTEST_PROD_TEMPLATE;
		var book = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+template);

		try{

		
			[records_cr, records_problems, records_defects].forEach(function(records,index){
					
				//populating crs
				log("sheet "+index)
				book.Sheets(index+1).Activate();
				fill_sheet(book.Sheets(index+1),records);
				
			});
	

		}catch(ex){
			
			log(ex.message);
			
		}

		var d = new Date();
		
		//CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
		var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
		var new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+template.replace("YYMMDD",current_date);
		
		log("saving exported tickets: " + new_list);
		
		try{
			
			book.SaveCopyAs(new_list);
			produced_file =  new_list;
				
		}catch(ex){
			
			 log(ex.message);
		
		}finally{
			
			try{
				
				book.Close();
				
			}
			catch(ex){
				
				log("***ERROR*** eror closing book file");
				
			}
			
			
		}
		
		
	});
	
	
	return produced_file;
	
}


function import_eac_produced_file_into_access(produced_file){
	
	var rows_cr 	  = null;
	var rows_problems = null;
	var rows_defects  = null;
	
	do_in_excel(function(excel){
            
        log("reading data from " + produced_file);
        var book = null;
		try{
			
			book = excel.Workbooks.Open(produced_file);
			
			
			log(".......")
			
			book.Sheets(1).Activate();
			rows_cr = read_sheet_data(book.Sheets(1));
			log("n° of cr: "+rows_cr.length);
			
			book.Sheets(2).Activate();
			rows_problems = read_sheet_data(book.Sheets(2));
			log("n° of problems: "+rows_problems.length);
			
			book.Sheets(3).Activate();
			rows_defects = read_sheet_data(book.Sheets(3));
			log("n° of defects: "+rows_defects.length);
			
			
			
			
		}
		catch(ex){
	
			log(ex.message);
	
		}finally{
			
			try{
				
				if(book!= null)
					book.Close();
			}
			catch(ex){
				
				log("***ERROR*** error closing the book")
				
			}
			
		}
  
	});
	

	
	var release_notes = rows_cr.concat(rows_problems).concat(rows_defects).map(function(elem){
		//log(elem["Release Note"]);
		return elem["Release Note"];
		
	}).reduce(function(acc,e){
                                         
			
		if(acc.indexOf(e) == -1){
			
			acc.push(e);
			
		}
		return acc
	
	}, []);
	
	log(release_notes);
	
	
	do_in_access(function(db){ 
        
		var in_clause = "("+release_notes.map(function(e){return "'"+e+"'"}).join(",")+")";
        log("deleting tickets for RNs: " + in_clause);
        db.Execute("DELETE FROM RN_EAC WHERE RN_EAC.[Release Note] IN "+in_clause)
		db.Execute("DELETE FROM RN_EAC_CRS WHERE RN_EAC_CRS.[Release Note] IN "+in_clause)
        
		var insert_sql = "INSERT INTO RN_EAC ([Problem ID], [Vendor Ticket Number], [L2], [Release Note], [Summary], [Filename], [Export Date], [ENV],	[EAC DD], [MIG1 DD], [MIG2 DD]) VALUES ";
		
		
		//4CB ticket	Case Type	Customer Trouble case id	L2 reference	Description	Release Note	Filename	Export Date	ENV	EAC DD	MIG1 DD	MIG2 DD

		var fields = ['Problem ID', 'Vendor Ticket Number', 'L2', 'Release Note', 'Summary', 'Filename', 'Export Date', 'ENV',	'EAC DD', 'MIG1 DD', 'MIG2 DD'];
		
		rows_problems.concat(rows_defects).map(function(row){
			
			row['Problem ID']			=	row['4CB ticket'];
			row['Vendor Ticket Number'] =	row['Customer Trouble case id'];
			row['L2']                   = 	row['L2 reference'];
			row['Summary']              = 	row['Description'];
			
			return row;
			
			
		}).forEach(function(row){
			
			var values = fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql + "("+ values + ")";
			
			log(sql);
			db.Execute(sql);
			
			
		});
		
		//CR Reference	Title	Delivery to Interoperability environment	Link to updated schema/message documentation	Release Note	Filename	Export Date	ENV	EAC DD	MIG1 DD	MIG2 DD

		var insert_sql_cr = "INSERT INTO RN_EAC_CRS ([CR Reference], [Title], [Delivery to Interoperability environment], [Link to updated schema/message documentation], [Release Note], [Filename], [Export Date], [ENV],	[EAC DD], [MIG1 DD], [MIG2 DD]) VALUES ";
		var cr_fields = ['CR Reference', 'Title', 'Delivery to Interoperability environment', 'Link to updated schema/message documentation', 'Release Note', 'Filename', 'Export Date', 'ENV', 'EAC DD', 'MIG1 DD', 'MIG2 DD'];
		rows_cr.forEach(function(row){
			
			var values = cr_fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql_cr + "("+ values + ")";
			
			log(sql);
			db.Execute(sql);
			
			
		});
		
		
	},RN_DB_PATH);
	
	
	
}


function import_utest_produced_file_into_access(produced_file){
	
	var rows_cr 	  = null;
	var rows_problems = null;
	var rows_defects  = null;
	
	do_in_excel(function(excel){
            
        log("reading data from " + produced_file);
        var book = null;
		try{
			
			book = excel.Workbooks.Open(produced_file);
			
			
			log(".......")
			
			book.Sheets(1).Activate();
			rows_cr = read_sheet_data(book.Sheets(1));
			log("n° of cr: "+rows_cr.length);
			
			book.Sheets(2).Activate();
			rows_problems = read_sheet_data(book.Sheets(2));
			log("n° of problems: "+rows_problems.length);
			
			book.Sheets(3).Activate();
			rows_defects = read_sheet_data(book.Sheets(3));
			log("n° of defects: "+rows_defects.length);
			
			
			
			
		}
		catch(ex){
	
			log(ex.message);
	
		}finally{
			
			try{
				
				if(book!= null)
					book.Close();
			}
			catch(ex){
				
				log("***ERROR*** error closing the book")
				
			}
			
		}
  
	});
	
	var release_notes = rows_cr.concat(rows_problems).concat(rows_defects).map(function(elem){
		
		return elem["Release Note"];
		
	}).reduce(function(acc,e){
                                         
			
		if(acc.indexOf(e) == -1){
			
			acc.push(e);
			
		}
		return acc
	
	}, []);
	
	log(release_notes);
	
	
	do_in_access(function(db){ 
        
		//Release Note	4CB ticket	Case Type	Customer Trouble case id	L2 reference	Summary	RN	Component	Retest status	Filename	Export Date	ENV	UTEST DD

		var in_clause = "("+release_notes.map(function(e){return "'"+e+"'"}).join(",")+")";
        log("deleting tickets for RNs: " + in_clause);
        db.Execute("DELETE FROM RN_UTEST WHERE RN_UTEST.[Release Note] IN "+in_clause)
        db.Execute("DELETE FROM RN_UTEST_CRS WHERE RN_UTEST_CRS.[Release Note] IN "+in_clause)
		
		var insert_sql = "INSERT INTO RN_UTEST ([Release Note], [4CB ticket], [Case Type], [Customer Trouble case id], [L2 reference], [Summary], [RN],	[Component],[Retest status],[Filename],[Export Date],[ENV],	[UTEST DD]) VALUES ";
		
		var fields = ['Release Note',	'4CB ticket',	'Case Type',	'Customer Trouble case id',	'L2 reference',	'Summary',	'RN',	'Component',	'Retest status',	'Filename',	'Export Date',	'ENV',	'UTEST DD'];
		
		rows_problems.concat(rows_defects).forEach(function(row){
			
			var values = fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql + "("+ values + ")";
			
			//log(sql);
			db.Execute(sql);
			
			
		});
		
		var insert_sql_cr = "INSERT INTO RN_UTEST_CRS ([CR Reference], [Title], [Delivery to Interoperability environment], [Link to updated schema/message documentation], [Release Note], [Filename], [Export Date], [ENV],	[UTEST DD]) VALUES ";
		var cr_fields = ['CR Reference', 'Title', 'Delivery to Interoperability environment', 'Link to updated schema/message documentation', 'Release Note', 'Filename', 'Export Date', 'ENV', 'UTEST DD'];
		rows_cr.forEach(function(row){
			
			var values = cr_fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql_cr + "("+ values + ")";
			
			log(sql);
			db.Execute(sql);
			
			
		});
		
		
		
	},RN_DB_PATH);
	
	
	
}

function import_prod_produced_file_into_access(produced_file){
	
	var rows_cr 	  = null;
	var rows_problems = null;
	var rows_defects  = null;
	
	do_in_excel(function(excel){
            
        log("reading data from " + produced_file);
        var book = null;
		try{
			
			book = excel.Workbooks.Open(produced_file);
			
			
			log(".......")
			
			book.Sheets(1).Activate();
			rows_cr = read_sheet_data(book.Sheets(1));
			log("n° of cr: "+rows_cr.length);
			
			book.Sheets(2).Activate();
			rows_problems = read_sheet_data(book.Sheets(2));
			log("n° of problems: "+rows_problems.length);
			
			book.Sheets(3).Activate();
			rows_defects = read_sheet_data(book.Sheets(3));
			log("n° of defects: "+rows_defects.length);
			
			
			
			
		}
		catch(ex){
	
			log(ex.message);
	
		}finally{
			
			try{
				
				if(book!= null)
					book.Close();
			}
			catch(ex){
				
				log("***ERROR*** error closing the book")
				
			}
			
		}
	
	});
	
	
	var release_notes = rows_cr.concat(rows_problems).concat(rows_defects).map(function(elem){
		
		return elem["Release Note"];
		
	}).reduce(function(acc,e){
                                         
			
		if(acc.indexOf(e) == -1){
			
			acc.push(e);
			
		}
		return acc
	
	}, []);
	
	log(release_notes);
	
	
	do_in_access(function(db){ 
        
		var in_clause = "("+release_notes.map(function(e){return "'"+e+"'"}).join(",")+")";
        log("deleting tickets for RNs: " + in_clause);
        db.Execute("DELETE FROM RN_PROD WHERE RN_PROD.[Release Note] IN "+in_clause)
		db.Execute("DELETE FROM RN_PROD_CRS WHERE RN_PROD_CRS.[Release Note] IN "+in_clause)
        
		var insert_sql = "INSERT INTO RN_PROD ([Release Note], [4CB ticket], [Case Type], [Customer Trouble case id], [L2 reference], [Summary], [RN],	[Component],[Retest status],[Filename],[Export Date],[ENV],	[PROD DD]) VALUES ";
		
		var fields = ['Release Note',	'4CB ticket',	'Case Type',	'Customer Trouble case id',	'L2 reference',	'Summary',	'RN',	'Component',	'Retest status',	'Filename',	'Export Date',	'ENV',	'PROD DD'];
		
		rows_problems.concat(rows_defects).forEach(function(row){
			
			var values = fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql + "("+ values + ")";
			
			//log(sql);
			db.Execute(sql);
			
			
		});
		
		var insert_sql_cr = "INSERT INTO RN_PROD_CRS ([CR Reference], [Title], [Delivery to Interoperability environment], [Link to updated schema/message documentation], [Release Note], [Filename], [Export Date], [ENV],	[PROD DD]) VALUES ";
		var cr_fields = ['CR Reference', 'Title', 'Delivery to Interoperability environment', 'Link to updated schema/message documentation', 'Release Note', 'Filename', 'Export Date', 'ENV', 'PROD DD'];
		rows_cr.forEach(function(row){
			
			var values = cr_fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql_cr + "("+ values + ")";
			
			log(sql);
			db.Execute(sql);
			
			
		});
		
	},RN_DB_PATH);
	
	
	
}

function import_cl_into_access(produced_file){
	
	var rows = null;
	
	var pattern = /.+(?:Consolidated)\s(?:list)\s(?:of)\s(?:open)\s(?:incidents)\s(?:and)\s(?:defects)_(\d{6})(-\d+)?\.xls(x)?$/;

	var r = produced_file.match(pattern)

	var cl_date = r[1]

	log("Consolidated list date: "+cl_date)
	
	do_in_excel(function(excel){
            
        log("reading data from " + produced_file);
        var book = null;
		try{
			
			book = excel.Workbooks.Open(produced_file);
			book.Sheets(1).Activate();
			
			log(".......")
			
			rows = read_sheet_data(book.Sheets(1));
			log("n° of tickets in "+produced_file+" :"+rows.length);
			
			
			
			
		}
		catch(ex){
	
			log(ex.message);
	
		}finally{
			
			try{
				
				if(book!= null)
					book.Close();
			}
			catch(ex){
				
				log("***ERROR*** error closing the book")
				
			}
			
		}
  
	});
	
	
	
	do_in_access(function(db){ 
        
		
        log("deleting tickets fro CL");
        db.Execute("DELETE FROM CL WHERE CL_Date = '" + cl_date+"'");
        //Problem ID*+	Problem/Defect	Summary*	Priority*	Environment	Submit Date	Category	Affected Functionality	System Entity	Status	External Reference	Resolution Date	Release number	Related Problems	Additional information	Updated	Notes

		var insert_sql = "INSERT INTO CL ([Problem ID*+], [Problem/Defect], [Summary*], [Priority*], [Environment], [Submit Date], [Category], [Affected Functionality], [System Entity], [Status], [External Reference], [Resolution Date], [Release number], [Related Problems], [Additional information], [Updated],	 [Notes], [CL_Date]) VALUES ";
		
		var fields = ['Problem ID*+', 'Problem/Defect', 'Summary*', 'Priority*', 'Environment', 'Submit Date', 'Category', 'Affected Functionality', 'System Entity', 'Status', 'External Reference', 'Resolution Date', 'Release number', 'Related Problems', 'Additional information', 'Updated',	 'Notes', 'CL_Date'];
		
		rows.forEach(function(row){
			
			row['CL_Date'] = cl_date;
			
			var values = fields.reduce(function(acc,e){
				
				acc.push("'"+row[e].replace(/'/g, "''")+"'");
				
				return acc;
				
			},[]).join(",");
			
			var sql = insert_sql + "("+ values + ")";
			
			//log(sql);
			db.Execute(sql);
			
			
		});
		
		
		
	},RN_DB_PATH);
	
	
	
	
}


function execute_omgchanges(){
		
	var command = CURRENT_FOLDER.replace(/\\/g, "/")+EXEC_COMMAND+" ";
	
	log(command)
	
	var objShell = new ActiveXObject("Wscript.Shell");
	objShell.Run('"'+command+'"',1,true);
	
}




//human communicaton interface
var WORKING_DIRECTORY = load_working_directory();//CURRENT_FOLDER+INPUT_FOLDER+"\\";
write_line("insert working directory");
write("(default "+WORKING_DIRECTORY+" ):");
var tmp = read_line().trim();
if(tmp != ""){
    
    WORKING_DIRECTORY = tmp.replace(/\\/g, "/")+"/";
}

write_line("selected: "+WORKING_DIRECTORY);
save_working_directory(WORKING_DIRECTORY);


var choice = 0;

while( [1,2,3,4,5,6,7,8,9,10,11].indexOf(choice) == -1){
    
    write_line("Select option to run:");
    write_line("(1) create consolidated list from access");
    write_line("(2) run consistency analysis on file + json export");
    write_line("(3) generate breakdown list per NCB");
    write_line("(4) generate operday list from access");
    write_line("(5) compare consolidated lists (Work in Progress)");
    write_line("(6) export consolidated lists in json format");
    write_line("(7) create weekly indicators (DO NOT USE)");
    write_line("(8) analyze related problems");
	write_line("(9) export release note ");
	write_line("(10) import consolidated list into access ");
	write_line("(11) map categories for UTSU ");
	
    
   
    
    write(">:");
    choice = read_line().trim();
    log(choice);
    choice = parseInt(choice);
    
}
switch(choice){
    
    case 1:     (function(){//create consolidated list from access
                    log("creating consolidated list from access");
                    var produced_file = create_consolidated_list_from_access();
					
					/*
					var imp= read_choice_from_input("can I import the consolidated list "+produced_file+"into Access?",["Y","N"]);
					if(imp == "Y"){
						
						import_cl_into_access(produced_file);
					}*/
               })();
                break;
    
    case 2:     (function(){//run consistency analysis on file + json export
        
                    var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(\d{6})(-\d+)?\.xls(x)?$/;
                    var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                    var consolidated_list_file = select_file_from_folder(WORKING_DIRECTORY,message,pattern); 
                    
                    log("running consistency analysis on: "+consolidated_list_file);    
                    var problems = null;
                    var incidents = null;
                    do_in_excel(function(excel){
                    
                        try{
                            var book = excel.Workbooks.Open(consolidated_list_file);
                            book.Sheets(1).Activate();
                        
                            problems  = read_sheet_data(book.Sheets(1));
                            
                            book.Sheets(2).Activate();
                            
                            incidents = read_sheet_data(book.Sheets(2));
                        
                            book.Close();
                            
                          
                        }catch(ex){

                            log(ex.message);

                        }
                    });
                    log("n° of problems: "+problems.length);

                    log("running consistency analysis...")
                    //var consistency_analysis_report = run_consistency_analysis(problems,linked_problems,tutto_problems);
                    var consistency_analysis_report = run_consistency_analysis(problems);
                    log("consistency analysis done.");

                    log(consistency_analysis_report);
                    write_report_to_file(consistency_analysis_report,consolidated_list_file);
                    
                    log("exporting data in json format");
                    
                    var result = pattern.exec(consolidated_list_file);
                    var date = result[8];
                    export_consolidatedlist_to_json(problems,incidents,date);
                    log("data exported");
        
        
               })();
                break;
    case 3:
                (function(){//generate breakdown list per NCB
                var choice = null;   
                while(choice != 'y' && choice != 'n'){
                    write("do you want merge results with modified consolidated list? (y/n): ");
                    choice = read_line().toLowerCase();
                    log("your choice: "+choice);
                }
                
                var consolidated_list_file = null;
                if(choice == 'y'){
                    
                    var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(?:\d{6})(-\d+)?\.xls(x)?$/;
                    var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                    var folder  = WORKING_DIRECTORY;
                    
                    consolidated_list_file = select_file_from_folder(folder,message,pattern); 
                    
                
                }
                
                log("creating breakdown list");
                create_breakdownperncb_list_from_access(consolidated_list_file);
                   
               })();
                break;
    case 4:     
                (function(){//generate operday list from access
                var choice = null;   
                while(choice != 'y' && choice != 'n'){
                    write("do you want merge results with modified consolidated list? (y/n): ");
                    choice = read_line().toLowerCase();
                    log("your choice: "+choice);
                }
                
                var consolidated_list_file = null;
                if(choice == 'y'){
                    
                    var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(?:\d{6})(-\d+)?\.xls(x)?$/;
                    var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                    var folder  = WORKING_DIRECTORY;
                    
                    consolidated_list_file = select_file_from_folder(folder,message,pattern); 
                    
                
                }
                
                log("creating operday list");
                create_operday_list_from_access(consolidated_list_file);
                   
               })();
                break;
    
    case 5:     (function(){//compare consolidated lists
                
                var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(?:\d{6})(-\d+)?\.xls(x)?$/;
                var previous_message = "copy previous consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                var current_message  = "copy current consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                var folder = WORKING_DIRECTORY;
                
                
                var previous_consolidated_list_file      = select_file_from_folder(folder,previous_message,pattern); 
                var current_consolidated_list_file       = select_file_from_folder(folder,current_message,pattern);
                
                var previous_problems = null, current_problems = null;
                
                do_in_excel(function(excel){
                    
                    try{
                        var book = excel.Workbooks.Open(previous_consolidated_list_file);
                        book.Sheets(1).Activate();
                        previous_problems  = read_sheet_data(book.Sheets(1));
                        book.Close();
                        
                        book = excel.Workbooks.Open(current_consolidated_list_file);
                        book.Sheets(1).Activate();
                        current_problems  = read_sheet_data(book.Sheets(1));
                        book.Close();

                    }catch(ex){

                        log(ex.message);

                    }
                });
                
                log("n° of previous problems: "+previous_problems.length);
                log("n° of current  problems: "+current_problems.length);

                log("running analysis...")
                var analysis_report = run_compare_consolidated_lists(previous_problems,current_problems);
                log("analysis done.");

                log(analysis_report);
                write_report_to_file(analysis_report,current_consolidated_list_file);
                
                
        
            })()
                break;   
               
    case 6:     (function(){//export consolidated lists in json format
        
                    var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(\d{6})(-\d+)?\.xls(x)?$/;
                    var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                    var consolidated_list_files = select_files_from_folder(WORKING_DIRECTORY,message,pattern); 
                    
                    consolidated_list_files.forEach(function(consolidated_list_file){
                        
                        var problems = null;
                        var incidents = null;
                        log(consolidated_list_file);
                        
                        do_in_excel(function(excel){
                    
                            try{
                                var book = excel.Workbooks.Open(consolidated_list_file);
                                book.Sheets(1).Activate();
                        
                                problems  = read_sheet_data(book.Sheets(1));
                            
                                book.Sheets(2).Activate();
                            
                                incidents = read_sheet_data(book.Sheets(2));
                        
                                book.Close();

                            }catch(ex){

                                log(ex.message);

                            }
                        });
                        log("n° of problems: "+problems.length);
                        log("n° of incidents: "+incidents.length);
                        log("exporting data in json format");
                        var result = pattern.exec(consolidated_list_file);
                        var date = result[8];
                        export_consolidatedlist_to_json(problems,incidents,date);
                        log("data exported");
                        
                        
                    });
                })();
                break;
    
    case 7:     (function(){//create weekly indicators
        
		
		log(' **** NOT WORKING use option (10) ****')
		
		
		/*
        var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(\d{6})(-\d+)?\.xls(x)?$/;
        var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
        var consolidated_list_file = select_file_from_folder(WORKING_DIRECTORY,message,pattern); 
                    
        log("generating weekly report data from: "+consolidated_list_file); 
        
        var problems = null;
        
        do_in_excel(function(excel){
                    
            try{
                var book = excel.Workbooks.Open(consolidated_list_file);
                book.Sheets(1).Activate();
                problems  = read_sheet_data(book.Sheets(1));
                book.Close();
                

            }catch(ex){

                log("****** ERROR ***** "+ex.message);

            }
        });
                
        log("n° of problems: "+problems.length);
        
        var result = pattern.exec(consolidated_list_file);
        var date = result[8];

        log("generating weekly report data for: "+date)
        var weekly_report = generate_weekly_report_data(problems,date);
        log("weekly report data generated done.");

       
        
        do_in_excel(function(excel){
            
            var d = new Date();
            
            //CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
            var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
            var new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+WEEKLYREPORT_TEMPLATE.replace("YYMMDD",current_date);
			var new_cl   = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+consolidated_list_file.split("\\").pop();
            
			var book = null;
			var cl_book = null;
			
            try{
                
                book 	= excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+WEEKLYREPORT_TEMPLATE);
				cl_book = excel.Workbooks.Open(consolidated_list_file);
            
                log("saving weekly report data");
                
                var open_tickets     = weekly_report[0];
                var pending_tickets  = weekly_report[1];
                var new_tickets      = weekly_report[2];
                var resolution_dates = weekly_report[3];
                
                
                
                
               				
				
				var resolution_chain_sheet  = cl_book.Sheets.Add();
				resolution_chain_sheet.Name= "Resolution Chain";
				
				
				var graphs_sheet			= cl_book.Sheets.Add();
				graphs_sheet.Name= "Graphs";
				
				//copy chart from template
				book.Sheets(2).Range("A1:I31").Copy(graphs_sheet.Range("A1:I31"));
				
				var alphabet = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";
				
				var current = 2;
				
				resolution_dates.forEach(function(rd){
					log(rd);
					
					graphs_sheet.Cells(1,current).Value = rd[0];
					graphs_sheet.Cells(2,current).Value = rd[1];
					graphs_sheet.Cells(3,current).Value = rd[2];
					current += 1;
					
				});
				
				
				
				var col=alphabet.charAt(current-1);//todo
				log("column: "+col)
				
				
				var chart = graphs_sheet.ChartObjects(1).Chart;
                log("chart name:" + chart.Name)
                chart.SetSourceData(graphs_sheet.Range("B1:"+col+"3"));
				
				var weekly_sheet			= cl_book.Sheets.Add();
				weekly_sheet.Name= "Weekly report";
				
				var start_row = 4;
                var current_row = start_row;
                weekly_sheet.Activate();
                open_tickets_labels = [open_tickets.label].concat(open_tickets.rows_label)
                open_tickets.data.forEach(function(elem,index){
                    
                    current_row = start_row+index;
                    weekly_sheet.Cells(current_row,1).Value = open_tickets_labels[index];
                    weekly_sheet.Cells(current_row,2).Value = elem;
                    
                })
                
                start_row = current_row+4;
                
                pending_tickets_labels = [pending_tickets.label].concat(pending_tickets.rows_label)
                pending_tickets.data.forEach(function(elem,index){
                    current_row = start_row+index;
                    weekly_sheet.Cells(current_row,1).Value = pending_tickets_labels[index];
                    weekly_sheet.Cells(current_row,2).Value = elem;
                    
                })
                
                start_row = current_row+4;
                
                new_tickets_labels = [new_tickets.label].concat(new_tickets.rows_label)
                new_tickets.data.forEach(function(elem,index){
                    current_row = start_row+index;
                    weekly_sheet.Cells(current_row,1).Value = new_tickets_labels[index];
                    weekly_sheet.Cells(current_row,2).Value = elem;
                    
                })
				
				
				
				
				
                
                book.SaveCopyAs(new_list);
                cl_book.SaveCopyAs(new_cl);
				
               
				
            
            }catch(ex){
            
                log("****** ERROR ***** "+ex.message);
            
            }finally{
				
				try{
					
					if(book)
						 book.Close();
					 
					if(cl_book) 
						 cl_book.Close();
					
				}catch(ex){
					
					log("****** ERROR ***** "+ex.message)
				}	
				
			}
            
        });*/
        
    })();
                break;
    case 8:     (function(){//analyze related problems
                    
                    log("checking linked problems");  

                    var report = "";
                    
                    var linked_problems = null, tutto_problems = null;
                    
                    var links_book = null, tutto_book = null;
                    
                    do_in_excel(function(excel){
                    
                        try{
                            
                            //loading linksvar 
                            log("loading links from : "+PROBLEMS_LINKS_FILE);   
                            var links_book = excel.Workbooks.Open(WORKING_DIRECTORY+PROBLEMS_LINKS_FILE);
                            links_book.Sheets(1).Activate();
                            linked_problems  = read_sheet_data(links_book.Sheets(1)).filter(function(t){
                                return t["Problem ID"] != null;
                                
                            });
                            log("file loaded");   
                            
                            
                            log("loading tutto from : "+PROBLEMS_TUTTO_FILE); 
                            var tutto_book = excel.Workbooks.Open(WORKING_DIRECTORY+PROBLEMS_TUTTO_FILE);
                            tutto_book.Sheets(1).Activate();
                            tutto_problems  = read_sheet_data(tutto_book.Sheets(1));
                            log("file loaded"); 
                            

                        }catch(ex){

                            log(ex.message);

                        }
                    });
                    
                    linked_groups =  linked_problems.reduce(function(acc,e){
                                         
                                            var id = e["Problem ID"].trim();
                                            var linked = e["Request ID"].trim();
                                            if(id in acc){
                                                acc[id].push(linked);
                                            }else{
                                                acc[id] = [linked];
                                            }
                                            return acc
                                        }, {});
                    
                    tutto_groups = tutto_problems .reduce(function(acc,e){
                                         
                                            var id = e["Problem ID*+"].trim();
                                            acc[id] = e;
                                            return acc
                                        }, {});
                    
                    
                    var dd = read_date_from_input();
                    
                    //convert in format: 2015/07/16 12:00:00 AM
                    var formatted_dd = (""+dd.getFullYear())+'/'+("0"+(dd.getMonth()+1)).slice(-2)+'/'+("0"+(dd.getDate())).slice(-2)+" 12:00:00 AM";
                    
                    log(dd+" formatted "+formatted_dd);
                    //1.	Check PoP senza EAC DD ? Check EAC DD dei ticket correlati ? se EAC DD= data in input ? ok per RN (chiedere cmq conferma in IAC/EAC telco + aggiornare TMS) 
                    //2.	Check PoP con EAC DD ? check EAC DD dei ticket correlate?  se sono uguali + uguale a data in input ? check stato ticket interno ? se PFE ? ok per RN (verde); se diverso da PFE chiedere in telco (giallo)
                    //3.	Check PFE ? se c’è anche EAC DD = data in input ? ok per RN (verde); se non c’è EAC DD ? to be checked during IAC/EAC telco (giallo)
                    //4.	Check EAC DD= data in input ? se stato PFE ? ok per RN (verde); se diverso da PFE ? to be checked during IAC/EAC telco (giallo)
       
                    
        
                    
                    //1.
                    var pops = tutto_problems.filter(function(elem){
                        
                        return elem["Public Pbm."] != 'No' && elem["Status*"] == 'Pending' && elem["StatusReason_Hidden"] == 'Pending Original Problem' && (elem["EAC Delivery Time"] == "" || elem["EAC Delivery Time"] == null);
                        
                    });
                    
                    //2.
                    var candidate_pops = tutto_problems.filter(function(elem){
                        
                        return elem["Public Pbm."] != 'No' && elem["Status*"] == 'Pending' && elem["StatusReason_Hidden"] == 'Pending Original Problem' && (elem["EAC Delivery Time"] != "" && elem["EAC Delivery Time"] != null);
                        
                    }); 
                    
                    //3.
                    
                    var pfes = tutto_problems.filter(function(elem){
                        
                        return elem["Public Pbm."] != 'No' && elem["Status*"] == 'Pending' && elem["StatusReason_Hidden"] == 'Future Enhancement';
                        
                    }); 
                    
                    var ready_pps = tutto_problems.filter(function(elem){
                        
                        return elem["Public Pbm."] != 'No' && elem["EAC Delivery Time"] == formatted_dd;
                        
                    }); 
                    
                    
                    var result = [];
                    
                    log("1)");
                    pops.forEach(function(pop){
                       
                        var id = pop['Problem ID*+'];
                        
                        if(id in linked_groups){
                            
                                
                                var lps = linked_groups[id].map(function(x){
                                    
                                    if(x in tutto_groups){
                                        
                                        return tutto_groups[x];
                                        
                                    }else{
                                        
                                        return null;
                                    }
                                }).filter(function(item){
                                    
                                    return item != null;
                                    
                                }).forEach(function(x){
                                    
                                    if (x["EAC Delivery Time"] == formatted_dd ){
                                        
                                        log(id+"("+pop["EAC Delivery Time"]+") : "+"\t"+x['Problem ID*+']+"\t("+x["EAC Delivery Time"]+")"+" <GREEN>")
                                        result.push({
                                            
                                            id: id,
                                            obj: pop,
                                            color:"green"
                                            
                                        });
                                    }
                                    
                                });
                            
                            
                        }else{
                            
                            log('ERROR: '+id)
                        }
                        
                    });
                    
                    log("2)");
                    candidate_pops.forEach(function(pop){
                       
                            var id = pop['Problem ID*+'];
                            var dd = pop["EAC Delivery Time"]
                            
                            if(id in linked_groups  ){
                                     
                                
                        
                                var lps = linked_groups[id].map(function(x){
                                    
                                    return x in tutto_groups ? tutto_groups[x] : null;
                                   
                                }).filter(function(item){
                                    
                                    return item != null;
                                    
                                }).forEach(function(x){
                                    
                                    if (x["EAC Delivery Time"] == formatted_dd && dd == formatted_dd ){
                                        
                                        if(x["Status*"] == "Pending" && x["StatusReason_Hidden"] == 'Future Enhancement'){

                                            
                                            log(id+"("+dd+") : "+"\t"+x['Problem ID*+']+"\t("+x["EAC Delivery Time"]+")"+" <GREEN>")
                                            result.push({
                                            
                                                id: id,
                                                obj: pop,
                                                color:"green"
                                                
                                            });
                                        
                                        }else{
                                            
                                            log(id+"("+dd+") : "+"\t"+x['Problem ID*+']+"\t("+x["EAC Delivery Time"]+")"+" <YELLOW>")
                                            result.push({
                                            
                                                id: id,
                                                obj: pop,
                                                color:"yellow"
                                                
                                            });
                                        }
                                        
                                        
                                    }else if(dd == formatted_dd && x["EAC Delivery Time"] != formatted_dd){
                                        
                                        log(id+"("+dd+") : "+"\t"+x['Problem ID*+']+"\t("+x["EAC Delivery Time"]+")"+" <BLANK>")
                                        result.push({
                                            
                                            id: id,
                                            obj: pop,
                                            color:"blank"
                                            
                                        });
                                        
                                    }
                                   
                                    
                                    
                                });
                                
                            }else{
                                
                                log('ERROR: '+id)
                            }
                            
                        
                        })
                    
                    log("3)");                    
                    pfes.forEach(function(pfe){
                       
                        var id = pfe['Problem ID*+'];
                        var dd = pfe["EAC Delivery Time"]
                        
                        if(dd == formatted_dd){
                            
                            log(id+"("+dd+") : "+"\t"+" <GREEN>")
                            result.push({
                                            
                                id: id,
                                obj: pfe,
                                color:"green"
                                
                            });
                        }else{
                            log(id+"("+dd+") : "+"\t"+" <YELLOW>")
                            result.push({
                                            
                                id: id,
                                obj: pfe,
                                color:"yellow"
                                
                            });
                        }
                        
                    });
                    
                    log("4)");
                    ready_pps.forEach(function(x){
                        var id = x['Problem ID*+'];
                        var dd = x["EAC Delivery Time"]
                        
                        if(!(x["Status*"] == 'Pending' && x["StatusReason_Hidden"] == 'Future Enhancement')){
                            
                            log(id+"("+dd+") : "+"\t"+" <YELLOW>")
                            result.push({
                                            
                                id: id,
                                obj: x,
                                color:"yellow"
                                
                            });
                            
                        }
                        
                    });
                    
                    //TODO group tickets for release & color 
                    log(result);
                    var grouped_tickets = result.reduce(function(acc,elem){
        
                        var release_number = elem.obj['PROD Release'];
                        var ticket_type    = elem.obj['Case Type'];

                        if(release_number == "" && ticket_type == 'Defect'){
                            
                            release_number = CURRENT_RELEASE;
                            
                        }else if(release_number == "" && ticket_type == 'Problem'){
                            
                            release_number = 'Not yet assigned to a Release';
                        }
                        else if(release_number == "" && ticket_type == 'To be defined'){
                            
                            release_number = 'ERROR';
                        }
                        
                        if(release_number in acc){
                            
                           acc[release_number].push(elem); 
                            
                        }else{
                            
                           acc[release_number] = [elem];
                            
                        }
                        
                        return acc;
                        
                    },{});
                    
                    
                    for(key in grouped_tickets){
                        
                        report+="Release: "+key+"\n";
                        grouped_tickets[key].forEach(function(x){
                            
                            report+=x.id+"\n";
                            
                        });
                       
                        
                        
                    }
                    
                    //extend analysis to all linked problems
                    write_report_to_file(report,PROBLEMS_TUTTO_FILE)
                    
                    do_in_excel(function(excel){
                    
                        var book = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+RN_TABLE_TEMPLATE);
                        
                        try{
                            var current_book = 1;
                            
                            for(key in grouped_tickets){
                        
                        
                                book.Sheets(current_book).Activate();
                                
                                grouped_tickets[key].forEach(function(x,index){
                            
                                    book.Sheets(current_book).Cells(index+2,1).Value = x.id;
                            
                                });
                       
                                
                                current_book++;
                                book.Sheets.Add;
                            }
                            
                            var d = new Date();
                            var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
                            var new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+RN_TABLE_TEMPLATE.replace("YYMMDD",current_date);
                            book.SaveCopyAs(new_list);
                            book.Close();
                        }catch(ex){
                            log(ex.message);
                        }
                    });
                     
                })();
                break;
    
	case 9:     (function(){//export release note 
					var env= read_choice_from_input("insert environment",["EAC","UTEST","PROD"]);
					var produced_file = export_release_note(env);
					var imp= read_choice_from_input("check the produced file "+produced_file+", can I import it into Access?",["Y","N"]);
					if(imp == "Y"){
						
						if(env == 'EAC')
							import_eac_produced_file_into_access(produced_file);
						else if(env == 'UTEST')
							import_utest_produced_file_into_access(produced_file);
						else if(env == 'PROD')
							import_prod_produced_file_into_access(produced_file);
							
					}
				})();
				break;
				
	case 10:

			(function(){//import consolidated list into access 
        
				var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(\d{6})(-\d+)?\.xls(x)?$/;
				var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
				var consolidated_list_files = select_files_from_folder(WORKING_DIRECTORY,message,pattern); 
				
				consolidated_list_files.forEach(function(consolidated_list_file){
					
					import_cl_into_access(consolidated_list_file);
					
				});
			})();
			break;
	case 11:
			
			(function(){
			    var pattern = /(Consolidated)\s(list)\s(of)\s(open)\s(incidents)\s(and)\s(defects)_(\d{6})(-\d+)?\.xls(x)?$/;
                var message = "copy consolidated list in folder \""+WORKING_DIRECTORY+"\" and press enter to continue ";
                var consolidated_list_file = select_file_from_folder(WORKING_DIRECTORY,message,pattern); 
				
				log("loading problems from: "+consolidated_list_file);    
                var problems = null;
                
				var book = null;
				
                do_in_excel(function(excel){
                    
                        try{
                            book = excel.Workbooks.Open(consolidated_list_file);
                            book.Sheets(1).Activate();
                        
                            problems  = read_sheet_data(book.Sheets(1));
                        
                        }catch(ex){

                            log(ex.message);

                        }finally{
			
							try{
								book.Close();
								
							}catch(ex){
								
								log("***ERROR*** - error while trying to close the workbook");
								
							}
							
						}
                });
                log("n° of problems: "+problems.length);
				
				
				log("loading omg problems");
				
				var omg_problems = null;
				
				var omg_book = null;
				
				do_in_excel(function(excel){
					
					try{
						
						omg_book = excel.Workbooks.Open(WORKING_DIRECTORY+PROBLEMS_NOTES_FILE);
            
						omg_book.Sheets(1).Activate();
                    
						var excluded_statuses = ["Cancelled","Closed",""];
            
						omg_problems = read_sheet_data(omg_book.Sheets(1)).filter(function(elem){
                        
							return excluded_statuses.indexOf(elem["Status*"]) == -1 ;
						});
		    
					}catch(ex){

                            log(ex.message);

                    }finally{
			
						try{
							omg_book.Close();
								
						}catch(ex){
								
							log("***ERROR*** - error while trying to close the workbook");
								
						}
							
					}
          
				});
			
				log("n° of tickets in "+PROBLEMS_NOTES_FILE+" :"+omg_problems.length);
				
				log("adding OMG category to problems...")
				
				problems.forEach(function(elem){
					
					omg_problems.forEach(function(omg_elem){
					
						if(elem["Problem ID*+"] == omg_elem["Problem ID*+"]){
							
							elem["OMG Category"] = omg_elem["Category"];
							return false;
						}
					
					});
					
				});
				
				var categories = null;
				
				var cat_books = null;
				
				do_in_excel(function(excel){
					
					try{
						
						cat_books = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+CATEGORIES_FILE);
            
						cat_books.Sheets(1).Activate();
                    
						categories = read_sheet_data(cat_books.Sheets(1));
		    
					}catch(ex){

                            log(ex.message);

                    }finally{
			
						try{
							cat_books.Close();
								
						}catch(ex){
								
							log("***ERROR*** - error while trying to close the workbook");
								
						}
							
					}
          
				});
				
				problems.forEach(function(elem){
					
					categories.forEach(function(cat){
					
						if(elem["OMG Category"] == cat["Category"]){
							
							elem["service"] = cat["service"];
							return false;
						}
					
					});
					
				});
				
				var items_4_and_5_tickets = problems.map(function(elem){
					
					var items_4_and_5_tkt = {};
					
					items_4_and_5_tkt["Problem ID*+"]		=elem["Problem ID*+"];
					items_4_and_5_tkt["Priority*"]			=elem["Priority*"];
					items_4_and_5_tkt["service"]			=elem["service"];
					items_4_and_5_tkt["Problem/Defect"]		=elem["Problem/Defect"];
					items_4_and_5_tkt["Status"]				=elem["Status"];
					items_4_and_5_tkt["Release number"]		=elem["Release number"];
					
					return items_4_and_5_tkt;
				});
				
				
				do_in_excel(function(excel){
					
					log("writing new categories to file...")
					
					var book = excel.Workbooks.Open(CURRENT_FOLDER+"\\"+TEMPLATE_FOLDER+"\\"+UTSU_CL_TEMPLATE);
					
					var new_list = "";
		
					try{
						
						
						
						//populating problems
						book.Sheets(1).Activate();
						fill_sheet(book.Sheets(1),items_4_and_5_tickets)
            
						//populating table
						
						/*
						C3->H6  Low - 4, Medium - 3, Urgent- 2, Critical - 1  | OPSS	COSE	SLMS	SDAS	INFS	N/A
						
						*/
						
						book.Sheets(2).Activate();
						var sheet = book.Sheets(2);
						
						log("item4...")
						
						var rows = ["4-Low", "3-Medium", "2-Urgent", "1-Critical"];
						var cols = ["OPSS","COSE","SLMS","SDAS","INFS","N/A"];
						
						var ROWS = rows.length, COLS= cols.length;
						
						var item4_reduce_create =  function(row,col){
							
							return function(previousValue, currentValue, currentIndex, array) {
                
								if(currentValue["Priority*"] == row && currentValue["service"] == col){
									 return previousValue+1;
								}else{
									 return previousValue;
								}
							
							};  
							
						};
						
						
						for(var i = 0; i < ROWS; i++){
							
							for(var j = 0; j < COLS; j++){
							
								
								var sum = items_4_and_5_tickets.reduce(item4_reduce_create(rows[i],cols[j]),0);
								
								sheet.Cells(3+i,3+j).Value = sum;
							
							}
							
						}
						
						
						log("item5...")
						
						//item5
						var item5_rows = ["R 1.3", "R 1.3 Defects", "R 1.3 Minor Change", "Future releases", "not defined", "Minor changes", "HFs requested", "Any other"];
						var item5_cols = ["Work in Progress","Pending client action required","Under Customer Retest","Waiting for delivery in UTEST/PROD","Pending doc enhancement"];
						
						
						//map release number to UTSU format
						log("mapping release number to UTSU format")
						var mapping = load_mapping("utsu_rn_mapping.properties");
						
						items_4_and_5_tickets.forEach(function(elem){
							
							var rn 			= elem["Release number"].toUpperCase();
							var case_type 	= elem["Problem/Defect"].toUpperCase();
							var rn_ext = (rn+"|"+case_type);
							
							
							if(rn in mapping){
								
								elem["RN"]=mapping[rn];
								
								
							}else if( rn_ext in mapping){
								
								elem["RN"]=mapping[rn_ext];
								
								
							}else if(rn.trim() == "" ){
								
								elem["RN"] = "not defined";
								
								
							}else{
								
								log("***WARNING**** unrecognized RN: "+elem["Problem ID*+"]+" =>" +rn);
								
								
								elem["RN"] = "Any other";
								
							}
							
							
							
							
							if( elem["Status"] == "Pending Documentation Enhancement"){
								
								elem["STS"] = "Work in Progress";
								
							}else if(elem["Status"] == "Pending OMG Release decision"){
								
								elem["STS"] = "Work in Progress";
								
							}else if(elem["Status"] == "Pending Client Action Required"){
								
								elem["STS"] = "Work in Progress";
								
							}else{
								
								elem["STS"] = elem["Status"];
								
							}
							
							
						});
						
						
						ROWS = item5_rows.length, COLS= item5_cols.length;
						
						var item5_reduce_create =  function(row,col){
							
							return function(previousValue, currentValue, currentIndex, array) {
                
								if(currentValue["RN"] == row && currentValue["STS"] == col){
									 return previousValue+1;
								}else{
									 return previousValue;
								}
							
							};  
							
						};
						
						
						for(var i = 0; i < ROWS; i++){
							
							for(var j = 0; j < COLS; j++){
							
								
								var sum = items_4_and_5_tickets.reduce(item5_reduce_create(item5_rows[i],item5_cols[j]),0);
								
								sheet.Cells(18+i,3+j).Value = sum;
							
							}
							
						}

						
			
			
			
						var d = new Date();
        
						//CURRENT_FOLDER+"\\"+OUTPUT_FOLDER+"\\"+
						var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();
						new_list = CURRENT_FOLDER.replace(/\\/g, "/")+OUTPUT_FOLDER+"/"+UTSU_CL_TEMPLATE.replace("YYMMDD",current_date);
					
						log("saving UTSU consolidated list: " + new_list);
					
						book.SaveCopyAs(new_list);
		
						}catch(ex){
							
							 log(ex.message);
						
						}finally{
							
							try{
								book.Close();
								
							}catch(ex){
								
								log("***ERROR*** - error while trying to close the workbook");
								
							}
							
						}
						
						return new_list;
		
			
					log("process completed.")
					
				});
				
			})();
			break;
}



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

function load_working_directory(){
	
	var path = CURRENT_FOLDER+"/.workspace";
    var workspace = read_all_text_file(path);
	if(!workspace){
		
		return CURRENT_FOLDER+INPUT_FOLDER+"\\";
	
	}else{
		
		return workspace.trim();
	}
	
}

function save_working_directory(working_directory){
	
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

function load_properties(properties_file){
    
    
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


function load_mapping(mapping_file){
    
	var obj ={};
    
    var path = CURRENT_FOLDER+"/"+mapping_file;
    var lib = read_all_text_file(path);
    var rows = lib.split("\n");
    log("reading configuration...");
    for(var i= 0;i < rows.length; i++){
        var items = rows[i].trim().split("="); 
        var new_row = "";
        if(items.length==2){
			
			obj[items[0].toUpperCase()]=items[1]
        }
        
    }
    log("mapping loaded.");
	return obj
    
}

