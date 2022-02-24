load("core");
load("system");
load("helpers");

log("app folder\t"+APP_FOLDER);
log("current folder\t"+CURRENT_FOLDER)
log("current path\t"+CURRENT_PATH)
log("root folder\t"+ROOT_FOLDER)

function callback(resp){
	
	log(resp)
	
}
	
try{

	http_request('https://www.w3schools.com/','TEST',callback);

}catch(e){
	
	log('error:\t'+e);
	
}


http_request('https://www.w3schools.com/','GET',callback);
	
	