var debug = false;

/**
* global var section
*/
var menuLabel = "Lbc Alertes";
var menuMailSetupLabel = "Setup email";
var menuSearchOnlyLabel = "Rechercher maintenant";
var menuSearchLabel = "Rechercher & mailer";
var menuTestLabel = "Tester";
var menuLog = "Activer/Désactiver les logs";
var menuArchiveLog = "Archiver les logs";

function test() {

  return;
}

function get_month_id(month_str) {

  var months = ["janv","fév","mars","avril","mai","juin","juillet","août","sept","octobre","novembre","décembre"];
  if (months.indexOf(month_str) < 0)
    return -1;
  
  mt = months.indexOf(month_str); 
  return  mt;
}

function lbc(sendMail){
  
  // Arrange sendMail var for trigger call
  if(sendMail == null || sendMail != false){
    sendMail = true;
  }
  
  // Checks email settings
  var to = ScriptProperties.getProperty('email');
  if(to == "" || to == null ){
    Browser.msgBox("L'email du destinataire n'est pas définit. Allez dans le menu \"" + menuLabel + "\" puis \"" + menuMailSetupLabel + "\".");
    return -1;
  }// If recepient not OK
  

  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Données");
  var slog = ss.getSheetByName("Log");
  var i = 0; var nbSearchWithRes = 0; var nbResTot = 0;
  var body = ""; var corps = ""; var bodyHTML = ""; var corpsHTML = ""; var menu = ""; var searchURL = ""; var searchName = "";
  var stop = false; var nbPages_toCheck = 1; var page_checked = 0;
  var today = new Date(); var date_annonce = new Date();

  
  // Handles different search: reads sequentially column B data
  while((searchURL = sheet.getRange(2+i,2).getValue()) != ""){
    
    // Imports query (String)
    searchName = sheet.getRange(2+i,1).getValue();   
    // Import last checking date (Date)
    var date_last_check = new Date(sheet.getRange(2+i,3).getValue());
    var nbRes = 0;
    //debug_("date last : " + date_last_check);
    
    // Handles different pages
    while (page_checked < nbPages_toCheck) {
      
      // Add 'o=' and page index to the URL request
      var rep = UrlFetchApp.fetch(searchURL + '&o=' + page_checked).getContentText("iso-8859-15");
      
      // Are there any results available on LBC?
      // TODO: detect blanck page (without "Aucune annonce" string)
      if(rep.indexOf("Aucune annonce") > 0){
        stop = true;
      }
      
      var data = splitResult_(rep);
      data = data.substring(data.indexOf("<a"));
      var firsta = extractA_(data);
      var id = extractId_(firsta);

      // Iterates over html data over all <a> tags
      while(data.indexOf("<a") >= 0 && stop == false) {

        var a = extractA_(data);
        var ea = extractId_(a);
        var title = extractTitle_(data);
        var place = extractPlace_(data);
        var price = extractPrice_(data);
        var date = extractDate_(data);
        
        // Import date value from html data
        date_annonce = eval(uneval(today));
        if (date.indexOf("Aujourd'hui") >= 0) {
          //debug_("annonce de today");
        }
        else if (date.indexOf("Hier") >= 0) {
          date_annonce.setDate(today.getDate() - 1);
          //debug_("annonce de hier");
        }
        else {
          date_annonce.setFullYear = today.getFullYear;
          var day_str = date.substring( date.indexOf("<div>") + 5, date.indexOf(" ", date.indexOf("<div>")) );
          var month_str = date.substring( date.indexOf(" ", date.indexOf("<div>")) + 1, date.indexOf("</div>", date.indexOf("<div>")));
          var day_ann = parseInt( day_str );
          date_annonce.setMonth( get_month_id(month_str) );
          date_annonce.setDate( day_ann );
          
        }
        date_annonce.setHours(  date.substring( date.indexOf(":") - 2, date.indexOf(":") )  );
        date_annonce.setMinutes(  date.substring( date.indexOf(":") + 1, date.indexOf(":") + 3 ) );
        debug_("date ann : " + date_annonce );
        
        if (date_annonce > date_last_check) {
          // Nouvelle annonce!
          nbRes++;

          // Imports the image
          var image = extractImage_(data);
          // Build quick access HTML output
          body += "<li><a href=\"#" + ea + "\">" + title + "</a> (" + price + " euros - " + place + ")</li>";
          // Build main HTML output (Date, image, price, location)
          bodyHTML += "<li style=\"list-style:none;margin-bottom:20px; clear:both;background:#EAEBF0;border-top:1px solid #ccc;\"><div style=\"float:left;width:90px;padding: 20px 20px 0 0;text-align: right;\">"+ date +"<div style=\"float:left;width:200px;padding:20px 0;\"><a href=\"" + a + "\">"+ image +"</a> </div><div style=\"float:left;width:420px;padding:20px 0;\"><a name=\"" + ea + "\" href=\"" + a + "\" style=\"font-size: 14px;font-weight:bold;color:#369;text-decoration:none;\">" + title + "</a> <div>" + place + "</div> <div style=\"line-height:32px;font-size:14px;font-weight:bold;\">" + price + "</div></div></li>";

        }// date_annonce > date_last_check
        
        else {
          stop = true;
        }// NOT date_annonce > date_last_check
        
        //debug_("title : " + title + " \n Conditions boucle : balise <a> : " + data.indexOf("<a") + " stop : " + stop);        

        // Eliminates current <a> tag (treated)
        data = data.substring(data.indexOf("<a",10));
              
      }// HTML <a> tags loop       
      
      // If all <a> tags treated => next page
      page_checked++;        
    
    }// Pages loop   

    
    nbResTot += nbRes;
    if(nbRes > 1) {
      //plusieurs results, on créé une liste "accès rapide"
      menu += "<li><a href=\"#"+ searchName + "\">"+ searchName +" (" + nbRes + ")</a></li>";
      menu += "<ol type=\"1\">" + body + "</ol>";

    }// If multiple results
    
    corpsHTML += "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbRes + ")</a></p><ul>" + bodyHTML + "</ul>";
    
    // Add data to the log sheet
    if(ScriptProperties.getProperty('log') == "true" || ScriptProperties.getProperty('log') == null || ScriptProperties.getProperty('log') == ""){
      slog.insertRowBefore(2);
      slog.getRange("A2").setValue(searchName);
      slog.getRange("B2").setValue(nbRes);
      slog.getRange("C2").setValue(today.toString());
    }// Log enabled
    
    Logger.log("Recherche effectuée pour " + searchName);
    //debug_("Searched " + searchName);
    nbRes = 0;
    body = "";
    bodyHTML = "";
    stop = false;
    
    // Updates last check date on Main sheet (String)
    sheet.getRange(2+i,3).setValue(today.toString());
    // Request advance to next search & reset page index
    i++;
    page_checked = 0;
    
  }// Until sheet row contains URL (C column)

  menu = "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Accès rapide :</p><ul>" + menu + "</ul>";
  //debug_(menu);
  debug_("Nb de res tot: " + nbResTot);
  
  // If new results, create alert format & send mail
  if (nbResTot >= 1) {
    var title = "Alerte leboncoin.fr : " + nbResTot + " nouveau" + (nbResTot>1?"x":"") + " résultat" + (nbResTot>1?"s":"");
    corps = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
    corpsHTML = "<body>" + menu + corpsHTML + "</body>";
  }

  else {
    var title = "Alerte leboncoin.fr : Pas de nouveau résultat, dsl !";
    corps = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
    corpsHTML = "<body> See you next time ! </body>";
  }// if nbResTot >= 2  
  
  // Send mail if needed
  if(sendMail == true) {
    MailApp.sendEmail(to,title,corps,{ htmlBody: corpsHTML });
    debug_("Nb mail journailier restant : " + MailApp.getRemainingDailyQuota());
  }// if send needed
  

}// lbc function

function searchOnly(){
  lbc(false);
}

function setupMail(){
  if(ScriptProperties.getProperty('email') == "" || ScriptProperties.getProperty('email') == null ){
    var quest = Browser.inputBox("Entrez votre email, le programme ne vérifie pas le contenu de cette boite.", Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      Browser.msgBox("Ajout email annulé.");
      return false;
    }else{
      ScriptProperties.setProperty('email', quest);
      Browser.msgBox("Email " + ScriptProperties.getProperty('email') + " ajouté");
    }
  }else{
    var quest = Browser.inputBox("Entrez un email pour modifier l'email : " + ScriptProperties.getProperty('email') , Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      Browser.msgBox("Modification email annulé.");
      return false;
    }else{
      ScriptProperties.setProperty('email', quest);
      Browser.msgBox("Email " + ScriptProperties.getProperty('email') + " ajouté");
    }
  }
}

/**
* Extrait l'id de l'annonce LBC
*/
function extractId_(id){
  return id.substring(id.indexOf("/",25) + 1,id.indexOf(".htm"));
}

/**
* Extrait le lien de l'annonce
*/
function extractA_(data){
  return data.substring(data.indexOf("<a") + 9 , data.indexOf(".htm", data.indexOf("<a") + 9) + 4);
}

/**
* Extrait le titre de l'annonce
*/
function extractTitle_(data){
  return data.substring(data.indexOf("title=") + 7 , data.indexOf("\"", data.indexOf("title=") + 7) );
}

/**
* Extrait le lieu de l'annonce
*/
function extractPlace_(data){
return data.substring(data.indexOf("placement") + 11 , data.indexOf("</div>", data.indexOf("placement") + 11) );
}

/**
* Extrait le prix de l'annonce
*/
function extractPrice_(data){
// test à optimiser car c'est hyper bourrin [mlb]
data = data.substring(0,data.indexOf("clear",10)); //racourcissement de la longueur de data pour ne pas aller chercher le prix du proudit suivant
var isPrice = String(data.substring(data.indexOf("price"), data.indexOf("price")+250)).match(/price/gi);
if (isPrice) {
var price = data.substring(data.indexOf("price") + 7 , data.indexOf("</div>", data.indexOf("price") + 7) );
} else {
var price = "";
}
return price;
}

/**
* Extrait la date de l'annonce
*/
function extractDate_(data){
return data.substring(data.indexOf("date") + 6 , data.indexOf("class=\"image\"", data.indexOf("date") + 6) - 5);
}

/**
* Extrait l'image de l'annonce
*/
function extractImage_(data){
// test à optimiser car c'est hyper bourrin [mlb]
var isImage = String(data.substring(data.indexOf("image"), data.indexOf("image")+250)).match(/img/gi);
if (isImage) {
var image = data.substring(data.indexOf("class=\"image-and-nb\">") + 21, data.indexOf("class=\"nb\"", data.indexOf("class=\"image-and-nb\">") + 21) - 12);
} else {
var image = "";
}
return image;
}

/**
* Extrait la liste des annonces
*/
function splitResult_(text){
var debut = text.indexOf("<div class=\"list-lbc\">");
var fin = text.indexOf("<div class=\"list-gallery\">");
return text.substring(debut + "<div class=\"list-lbc\">".length,fin);
}

//Activer/Désactiver les logs
function dolog(){
  if(ScriptProperties.getProperty('log') == "true"){
    ScriptProperties.setProperty('log', false);
    Browser.msgBox("Les logs ont été désactivées.");
  }else if(ScriptProperties.getProperty('log') == "false"){
    ScriptProperties.setProperty('log', true);
    Browser.msgBox("Les logs ont été activées.");
  }else{
    ScriptProperties.setProperty('log', false);
    Browser.msgBox("Les logs ont été désactivées.");
  }
}

//Archiver les logs
function archivelog(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slog = ss.getSheetByName("Log");
  var today  = new Date();
  var newname = "LogArchive " + today.getFullYear()+(today.getMonth()+1)+today.getDate();
  slog.setName(newname);
  var newsheet = ss.insertSheet("Log",1);
  newsheet.getRange("A1").setValue("Recherche");
  newsheet.getRange("B1").setValue("Nb Résultats");
  newsheet.getRange("C1").setValue("Date");
  newsheet.getRange(1,1,2,3).setBorder(true,true,true,true,true,true);
}


function onOpen() {
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var entries = [{
name : menuMailSetupLabel,
functionName : "setupMail"
},
  null
,{
name : menuSearchLabel,
functionName : "lbc"
},{
name : menuSearchOnlyLabel,
functionName : "searchOnly"
},{
name : menuTestLabel,
functionName : "test"
},
null
,{
name : menuLog,
functionName : "dolog"
},{
name : menuArchiveLog,
functionName : "archivelog"
}];
sheet.addMenu(menuLabel, entries);
}

function onInstall()
{
onOpen();
}

/**
* Retourne la date
*/
function myDate_(){
var today = new Date();
debug_(today.getDate()+"/"+(today.getMonth()+1)+"/"+today.getFullYear());
return today.getDate()+"/"+(today.getMonth()+1)+"/"+today.getFullYear();
}

/**
* Retourne l'heure
*/
function myTime_(){
var temps = new Date();
var h = temps.getHours();
var m = temps.getMinutes();
if (h<"10"){h = "0" + h ;}
if (m<"10"){m = "0" + m ;}
debug_(h+":"+m);
return h+":"+m;
}

/**
* Debug
*/
function debug_(msg) {
if(debug != null && debug) {
Browser.msgBox(msg);
}
}




