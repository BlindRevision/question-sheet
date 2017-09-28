function getProblemList(){
  var problemList = (PropertiesService.getUserProperties().getProperty("CROWDCALL_PROBLEM_LIST") == null) ? [] : eval("(" + PropertiesService.getUserProperties().getProperty("CROWDCALL_PROBLEM_LIST") + ')');
  return problemList;  
}

function getProblemToCreateList(){
  var problemList = (PropertiesService.getUserProperties().getProperty("CROWDCALL_PROBLEM_TO_CREATE_LIST") == null) ? [] : eval("(" + PropertiesService.getUserProperties().getProperty("CROWDCALL_PROBLEM_TO_CREATE_LIST") + ')');
  return problemList;  
}

function authorizeServices(){
  var servicesToAuth = (PropertiesService.getUserProperties().getProperty("SERVICES_TO_AUTH") == null) ? [] : eval("(" + PropertiesService.getUserProperties().getProperty("SERVICES_TO_AUTH") + ')');
  if(servicesToAuth.length > 0){
    eval("showDialog"+servicesToAuth[0]+"()");
  }
  else{
    manageProblemsToCreate();
  }
}

function getService(name){
  for(var i=0;i<SERVICE_LIST.length;i++){
    if(SERVICE_LIST[i].serviceName == name) return SERVICE_LIST[i];
  }
  return null;
}

function stopSharing(mesData){
  var problemList = getProblemList();
  for(var i=0;i<problemList.length;i++){
    var p = problemList[i];
    if((p.row == mesData.row)&&(p.column == mesData.column)&&(p.sheetName == mesData.sheetName)&&(p.spreadsheetId == mesData.spreadsheetId)){
      var formulaLag = p.solution;
      var locale = SpreadsheetApp.getActive().getSpreadsheetLocale();
      if(locale == "es_ES"){
        formulaLag = formulaLag.replace(/\,/g,"%%COMMA%%");
        formulaLag = formulaLag.replace(/\;/g,",");
        formulaLag = formulaLag.replace(/\%\%COMMA\%\%/g,";");
      }
      if(p.solution != null) SpreadsheetApp.openById(p.spreadsheetId).getSheetByName(p.sheetName).getRange(p.row, p.column).setFormula(formulaLag);
      problemList.splice(i,1);
    }
  }
  PropertiesService.getUserProperties().setProperty("CROWDCALL_PROBLEM_LIST", JSON.stringify(problemList));
  return true;
}

function setSolution(row,col,sheetName,spreadsheetId,formula){
  var r = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName).getRange(row, col);
  var callbackURL = generateCallbackURL(row,col,sheetName,spreadsheetId,"openSidebar");
  
  // TODO: hacerlo bien
  var formulaLag = formula;
  var locale = SpreadsheetApp.getActive().getSpreadsheetLocale();
  if(locale == "es_ES"){
    formulaLag = formulaLag.replace(/\,/g,"%%COMMA%%");
    formulaLag = formulaLag.replace(/\;/g,",");
    formulaLag = formulaLag.replace(/\%\%COMMA\%\%/g,";");
  }
  
  r.setFormula('=HYPERLINK("'+callbackURL+'";'+formulaLag+')');
  var problemList = getProblemList();
  var problem = getProblemFromList(row,col,sheetName,spreadsheetId,problemList);
  problem["solution"] = formula;
  
  
  var responses = problem.replies;
  var message = "";
  for(var i=0;i<responses.length;i++){
    if((responses[i]["extractedData"]!=null)&&(responses[i]["extractedData"]==formula)){
      message = responses[i].reply;
      break;
    }
  }
  var noteText = "This is a Q&ASheet-fed cell.\r\nLink to thread: ";
  for(var i=0;i<problem.questionMessages.length;i++){
    noteText += "\r\n"+problem.questionMessages[0].link;
  }
  r.setNote(noteText);
  
  
  PropertiesService.getUserProperties().setProperty("CROWDCALL_PROBLEM_LIST", JSON.stringify(problemList));
  return true;
}

function getProblemFromList(row,column,sheetName,spreadsheetId,list){
  if(list == null) return false;
  for(var i=0;i<list.length;i++){
    var p = list[i];
    if((p.row == row)&&(p.column == column)&&(p.sheetName == sheetName)&&(p.spreadsheetId == spreadsheetId)) return p;
  }
  return null;
}

function addSubproblemToList(services,receivers,subproblemType,message,sheetName,spreadsheetId,expression,column,row,list,title){
  var subproblem = {};
  subproblem["title"] = title;
  subproblem["services"] = services;
  subproblem["subproblemType"] = subproblemType;
  subproblem["message"] = message;
  subproblem["sheetName"] = sheetName;
  subproblem["spreadsheetId"] = spreadsheetId;
  subproblem["expression"] = expression;
  subproblem["column"] = column;
  subproblem["row"] = row;
  subproblem["receivers"] = receivers;
  list.push(subproblem);
}

function showSidebar(mesData){
  var a = HtmlService.createTemplateFromFile('sidebar');
  a.column = (mesData.column == null) ? null : mesData.column;
  a.row = (mesData.row == null) ? null : mesData.row;
  a.spreadsheetId = (mesData.spreadsheetId == null) ? null : mesData.spreadsheetId;
  a.sheetName = (mesData.sheetName == null) ? null : mesData.sheetName;
  a.backupSheetName = (mesData.backupSheetName == null) ? null : mesData.backupSheetName;
  a.expression = (mesData.expression == null) ? null : mesData.expression;
  a.services = (mesData.services == null) ? null : mesData.services;
  a.message = (mesData.message == null) ? null : mesData.message;
  a.problemType = (mesData.problemType == null) ? null : mesData.problemType;
  a.title = (mesData.title == null) ? null : mesData.title;
  a.receivers = (mesData.receivers == null) ? null : mesData.receivers;
  a.replies = (mesData.replies == null) ? null : mesData.replies;
  a.posted = mesData.posted;
  
  var b = a.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("Q&A sheet").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(b);  
}

function reloadSidebar(row,column,sheetName,spreadsheetId){
  var problemList = getProblemList();
  var p = getProblemFromList(row,column,sheetName,spreadsheetId,problemList);
  showSidebar({
    column: column,
    row: row,
    sheetName: sheetName,
    spreadsheetId: spreadsheetId,
    backupSheetName: p.backupSheetName,
    expression: p.expression,
    services: p.services,
    message: p.message,
    problemType: p.problemType,
    title: p.title,
    receivers: p.receivers,
    replies: p.replies,
    posted: true
  });
}

function showShareDialog(mesData){  
  var a = HtmlService.createTemplateFromFile('sharingDialog');
  a.column = mesData.column;
  a.row = mesData.row;
  a.expression = mesData.expression;
  a.sheetName = mesData.sheetName;
  a.backupSheetName = mesData.backupSheetName;
  a.spreadsheetId = mesData.spreadsheetId;
  a.message = mesData.message;
  a.title = mesData.title;
  a.problemType = mesData.problemType;

  var html = a.evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(512)
  .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Sharing options");
}

function generateCallbackURL(row,column,sheetName,spreadsheetId,functionName){
  /*var url = PropertiesService.getUserProperties().getProperty("CALLBACK_URL");
  url += "?functionName="+functionName;
  url += "&sheetName="+sheetName;
  url += "&spreadsheetId="+spreadsheetId;
  url += "&row="+row;
  url += "&column="+column;
  return url;*/

  var url = "https://script.google.com/macros/d/" + PROJECT_KEY + "/usercallback?state=";
  var state = ScriptApp.newStateToken()
  .withArgument("sheetName", sheetName)
  .withArgument("spreadsheetId", spreadsheetId)
  .withArgument("row", row)
  .withArgument("column", column)
  .withMethod(functionName)
  .createToken();  
  return url+state;
}

function question() {
  var problemList = getProblemList();
  var sheetName = SpreadsheetApp.getActiveSheet().getName();
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var column = SpreadsheetApp.getActiveRange().getColumn();
  var row = SpreadsheetApp.getActiveRange().getRow();
  var sp = getProblemFromList(row,column,sheetName,spreadsheetId,problemList);
  if(sp == null) return "Question editing...";
  else return "Question posted";
}

function crowdTrigger(e){
  var sheetName = e.range.getSheet().getName();
  var spreadsheetId = e.source.getId();
  var column = e.range.getColumn();
  var row = e.range.getRow();
  var pL = getProblemList();
  
  var p = getProblemFromList(row,column,sheetName,spreadsheetId,pL);
  if(p != null){
    showSidebar({
      column: column,
      row: row,
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      expression: p.expression,
      services: p.services,
      message: p.message,
      problemType: p.problemType,
      title: p.title,
      receivers: p.receivers,
      replies: p.replies,
      posted: true
    });        
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var range = sh.getRange(row,column);
    var callbackURL = generateCallbackURL(row,column,sheetName,spreadsheetId,"openSidebar");
    range.setFormula('=HYPERLINK("'+callbackURL+'";"Question posted")');
    return true;
  }
  else{
    // TODO: crear copia oculta de la tabla
    PropertiesService.getUserProperties().setProperty("ASKING_QUESTION", 1);
    var questionCell = {row: row, column: column};
    PropertiesService.getUserProperties().setProperty("QUESTION_CELL", JSON.stringify(questionCell));
    
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var backupSheetName = spreadsheet.getSheetByName(sheetName).copyTo(spreadsheet).hideSheet().getName();
    //PropertiesService.getUserProperties().setProperty("BACKUP_SHEET_NAME", backupSheetName);
    showSidebar({
      column: column,
      row: row,
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      backupSheetName: backupSheetName,
      expression: ((e.value == null)||(e.value=="")) ? e.range.getFormula() : e.value,
      posted: false
    });        
  }
}

function getExpectedValuesInserted(){
  var a = PropertiesService.getUserProperties().getProperty("EXPECTED_VALUES_INSERTED");
  if(a != null) return true;
  else return false;
}

function getRandomDataInserted(){
  var a = PropertiesService.getUserProperties().getProperty("RANDOM_DATA_INSERTED");
  if(a != null) return true;
  else return false;  
}

function generateImages(rangeList,spreadsheet){
  //var rangeList = JSON.parse(ranges);
  var imageList = [];
  for(var i=0;i<rangeList.length;i++){
    //var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ss = SpreadsheetApp.openById(spreadsheet);
    var sheet = ss.getSheetByName(rangeList[i].sheet);
    var w = 0;
    var h = 0;
    var range = sheet.getRange(rangeList[i].range).getValues();
    var aux = rangeList[i].range.match(/[\w\$]+(?=:)/)[0];
    var a = sheet.getRange(aux);
    var firstC = a.getColumn()-1;
    for(var j=a.getRowIndex();j<a.getRowIndex()+range.length;j++){
      h += ss.getRowHeight(j);
    }
    for(var j=a.getColumnIndex();j<a.getColumnIndex()+range[0].length;j++){
      w += ss.getColumnWidth(j);
    }
    var dataRange = Charts.newDataTable();
    var width = range[0].length;
    var firstColumn = rangeList[i].range.match(/^\D+(?=\d)/)[0];
    var firstRow = rangeList[i].range.match(/\d+(?=:)/)[0];
    var columnIndexes = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var cc = columnIndexes.split("");
    for(var j=0;j<range[0].length;j++){
      if((firstC + j) > 25){
        var p1=cc[Math.floor((firstC+j)/26)-1];
        var p2=cc[(firstC+j)%26];
        dataRange = dataRange.addColumn(Charts.ColumnType.STRING, p1+p2);
      }
      else{
        dataRange = dataRange.addColumn(Charts.ColumnType.STRING, cc[(firstC+j)]);      
      }
    }
    for(var j=0;j<range.length;j++){
      var valueList = [];
      for (var k=0;k<range[j].length;k++){
        valueList[k]=range[j][k].toString();
      }
      dataRange.addRow(valueList);
    }
    
    var chart = Charts.newTableChart()
    .setDataTable(dataRange.build())
    .setOption('width', w+20)
    .setOption('height',h+50)
    .setFirstRowNumber(firstRow)
    .showRowNumberColumn(true);
    chart = chart.build();
    
    var picture = chart.getAs("image/png").setContentTypeFromExtension();
    imageList.push(picture);
  }
  return imageList;
}

function openSidebar(args){
  var row = args.parameters.row;
  var column = args.parameters.column;
  var sheetName = args.parameters.sheetName;
  var spreadsheetId = args.parameters.spreadsheetId;
  var problemList = getProblemList();
  
  var r = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName).getRange(row, column);
  
  loadMessages();
  
  var callbackURL = generateCallbackURL(row,column,sheetName,spreadsheetId,"openSidebar");
  //r.setFormula('=HYPERLINK("'+callbackURL+'";"Question posted")');  
  
  var p = getProblemFromList(row,column,sheetName,spreadsheetId,problemList);
  if(p != null){
     showSidebar({
      column: column,
      row: row,
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      expression: p.expression,
      services: p.services,
      message: p.message,
      problemType: p.problemType,
      title: p.title,
      receivers: p.receivers,
      replies: p.replies,
      posted: true
    });
  }   
  return HtmlService.createHtmlOutputFromFile("sidebar").setSandboxMode(HtmlService.SandboxMode.IFRAME); 
}

function manageProblemsToCreate(){
  var problemToCreateList = getProblemToCreateList();
  if(problemToCreateList == null) return true;
  var problemList = getProblemList();
  while(problemToCreateList.length != 0){
    var p = problemToCreateList[0];
    var problem = {};
    problem["sheetName"]=p.sheetName;
    problem["spreadsheetId"]=p.spreadsheetId;
    //problem["cellRanges"] = p.cellRanges;
    problem["services"]= p.services;
    problem["expression"]=p.expression;
    problem["row"] = p.row;
    problem["column"] = p.column;
    problem["problemType"] = p.problemType;
    problem["message"] = p.message;
    problem["title"] = p.title;
    problem["receivers"] = p.receivers;

    var imageList = [];
    if(p.contextImage != null){
      problem["contextImage"] = p.contextImage; 
      var request = UrlFetchApp.fetch(p.contextImage, {
        method: "GET",
      });
      imageList.push(request.getBlob());
    }
    
    var ret;      
    var messageList = [];
    
    for(var i=0;i<p.services.length;i++){
      var service = getService(problem.services[i]);
      
      var sharedTableUrl = null;
      var imageList, sharedTable;
      var regExp = /\=screenshot\(\w{1,2}\d{1,3}\:\w{1,2}\d{1,3}\)/i;
      var match = p.message.match(regExp);
      if((match != null)&&(match.length>0)){
        var sampleRangeString = String(match[0]).slice(12,-1);
        
        var ran = {
          sheet: p.sheetName,
          range: sampleRangeString
        };
        //imageList = generateImages([ran],p.spreadsheetId);

        imageList = [p.contextImage];
        var spreadsheet = SpreadsheetApp.openById(p.spreadsheetId);
        var sampleRange = spreadsheet.getSheetByName(p.sheetName).getRange(sampleRangeString);
        var lagSheet = spreadsheet.insertSheet();
        sampleRange.copyTo(lagSheet.getRange(sampleRangeString));
        sharedTable = SpreadsheetApp.create("Shared table");
        lagSheet.copyTo(sharedTable).setName(p.sheetName);
        sharedTable.deleteSheet(sharedTable.getSheets()[0]);
        spreadsheet.deleteSheet(lagSheet);
        var f = DriveApp.getFileById(sharedTable.getId());
        f.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
        sharedTableUrl = sharedTable.getUrl();
      }
      if(p.title != null) ret = service.publishMessage(p.message,p.receivers,imageList,p.title,sharedTableUrl);
      else ret = service.publishMessage(p.message,p.receivers,imageList,null,sharedTableUrl);
      
      if(ret.error != null){
        return ret;
      }
      
      var date = new Date();
      ret["creationDate"] = date;
      messageList.push(ret);
    }

    problem["questionMessages"] = messageList;

    problemToCreateList.splice(0, 1);
    
    var backupSheetName = p.backupSheetName;
    var spreadsheet = SpreadsheetApp.openById(problem.spreadsheetId);
    var sheet = spreadsheet.getSheetByName(problem.sheetName);
    var sheetIndex = sheet.getIndex();
    spreadsheet.deleteSheet(sheet);
    
    // Spreadsheet.getSheetByName doesn't work
    var sheetList = spreadsheet.getSheets();
    for(var i=0;i<sheetList.length;i++){
      if(sheetList[i].getName() == backupSheetName){
        var s = sheetList[i].setName(problem.sheetName).showSheet();
        spreadsheet.setActiveSheet(s);
        spreadsheet.moveActiveSheet(sheetIndex);
        spreadsheet.setActiveSheet(s);
        break;
      }
    }
        
    var r = SpreadsheetApp.openById(p.spreadsheetId).getSheetByName(p.sheetName).getRange(p.row, p.column);
    var callbackURL = generateCallbackURL(p.row,p.column,p.sheetName,p.spreadsheetId,"openSidebar");
    r.setFormula('=HYPERLINK("'+callbackURL+'";"Question posted")');  
    
    problemList.push(problem);
  }
  
  var userProperties = PropertiesService.getUserProperties();      
  userProperties.setProperty("CROWDCALL_PROBLEM_LIST", JSON.stringify(problemList));
  userProperties.setProperty("CROWDCALL_PROBLEM_TO_CREATE_LIST", JSON.stringify([]));
  return true;
}

function crowdFunc(mes){
  var userProperties = PropertiesService.getUserProperties();
  var problemToCreateList = getProblemToCreateList();
  var problemList = getProblemList();
  var range = SpreadsheetApp.openById(mes.spreadsheetId).getSheetByName(mes.sheetName).getRange(mes.row, mes.column);
  var p = getProblemFromList(mes.row,mes.column,mes.sheetName,mes.spreadsheetId,problemList);
  if(p != null){
    var callbackURL = generateCallbackURL(p.row,p.column,p.sheetName,p.spreadsheetId,"openSidebar");
    range.setFormula('=HYPERLINK("'+callbackURL+'";"Question posted")');
  }
  else{
    var problem = {
      services: mes.services,
      sheetName: mes.sheetName,
      backupSheetName: mes.backupSheetName,
      spreadsheetId: mes.spreadsheetId,
      expressiom: mes.expression,
      row: mes.row,
      column: mes.column,
      problemType: mes.problemType,
      message: mes.message,
      title: mes.title,
      receivers: mes.receivers
    };
    if(mes.contextImage != null){
      var blob = generateBlobFromImage(mes.contextImage);
      var imageURL = uploadImageToImgur(blob);
      problem["contextImage"] = imageURL;
    }
    problemToCreateList.push(problem);
    userProperties.setProperty("CROWDCALL_PROBLEM_TO_CREATE_LIST", JSON.stringify(problemToCreateList));
    userProperties.deleteProperty("ASKING_QUESTION");
    userProperties.deleteProperty("EXPECTED_VALUES_INSERTED");
    userProperties.deleteProperty("RANDOM_DATA_INSERTED");
    userProperties.deleteProperty("QUESTION_CELL");
    var servicesToAuthorize = [];
    for(var i=0;i<mes.services.length;i++){
      var service = getService(mes.services[i]);
      if(!service.isAuthorized()) servicesToAuthorize.push(mes.services[i]);
    }
    if(servicesToAuthorize.length > 0){
      userProperties.setProperty("SERVICES_TO_AUTH", JSON.stringify(servicesToAuthorize));
      authorizeServices();
      return false;
    }
    else{
      var b = manageProblemsToCreate();
      if(b.error != null){
        return "error";
      }
      else{
        reloadSidebar(mes.row,mes.column,mes.sheetName,mes.spreadsheetId)
        return "message posted";
      }
    }      
  }
}

function loadMessages(row,col,sheetName,spreadsheetId){
  var ret = [];
  var problemList =  getProblemList();
  
  var emptyMessages = true;
  var v;

  for(var q=0;q<problemList.length;q++){
    var problem = problemList[q];      
    var messages = [];
    for(var i=0;i<problem.questionMessages.length;i++){
      var service = getService(problem.questionMessages[i].service);
      var mes = service.getReplies(problem.questionMessages[i]);
      messages = messages.concat(mes);
    }
    
    if(messages.length > 0){
      var lag = {};
      
      if(problem.replies==null) problem["replies"] = [];
      var ind;
      for(var i=0;i<messages.length;i++){
        var found = false;
        for(var j=0;j<problem.replies.length;j++){
          if(messages[i].replyId == problem.replies[j].replyId){
            found = true;
            ind = j;
            break;
          }
        }
        var type = problem.problemType.trim();
        var dt = eval("new "+type+"Type()");
        var extracted = dt.extract(messages[i].reply,problem);
        if((extracted != null)&&(extracted != "")){
          messages[i]["extractedData"] = extracted.data;
        }
        if(!found){     
          problem.replies.push(messages[i]);
        }
        else{
          problem.replies[ind] = messages[i];
        }
      }
    }
    
    if((problem.replies != null)&&(problem.replies.length>0)){ 
      var cells = [];
      for(var b=0;b<problem.replies.length;b++){
        if(problem.replies[b].extractedData != null){
          cells.push(problem.replies[b]);
        }
      } 
      if(cells.length > 0){
        if(problem.solution == null){
          var range = SpreadsheetApp.openById(problem.spreadsheetId).getSheetByName(problem.sheetName).getRange(problem.row,problem.column);
          var callbackURL = generateCallbackURL(problem.row,problem.column,problem.sheetName,problem.spreadsheetId,"openSidebar");
          range.setFormula('=HYPERLINK("'+callbackURL+'";"Question answered ('+cells.length+')")');
        }
          
        if((problem.row == row)&&(problem.column == col)&&(problem.spreadsheetId == spreadsheetId)&&(problem.sheetName == sheetName)){
          ret = cells;
        }
      }
    }
  }
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("CROWDCALL_PROBLEM_LIST", JSON.stringify(problemList));
  return ret;
}

function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  var addonMenu = ui.createAddonMenu();  
  addonMenu.addItem("Refresh Answers","loadMessages").addToUi();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  if(triggers.length == 0){
     ScriptApp.newTrigger('tEdit')
     .forSpreadsheet(ss.getId())
     .onEdit()
     .create();
  }
}

function getSelectedRange(){
  var range = SpreadsheetApp.getActiveRange();
  if((range.getWidth()==1)&&(range.getHeight()==1)) return "empty";
  else return range.getA1Notation();
}

function tEdit(e){
  var asking = PropertiesService.getUserProperties().getProperty("ASKING_QUESTION");
 
  if(asking == 1){
    var formula = String(e.range.getFormula());
    var regExpCoverage = /coverage\([^\)]*\)/gi;
    if(!regExpCoverage.test(formula)){
      var questionCell = PropertiesService.getUserProperties().getProperty("QUESTION_CELL");
      if(questionCell != null){
        var l = JSON.parse(questionCell);
        if((e.range.getColumn() == l.column)||(e.range.getRow() == l.row)){
          e.range.setBackground("red");
          PropertiesService.getUserProperties().setProperty("EXPECTED_VALUES_INSERTED",1);
        }
      }
    }
  }
  
  var regExpCrowd = /question\([^\)]*\)/gi;
  var regExpCoverage = /coverage\([^\)]*\)/gi;


  if((e.range.getWidth() > 1)||(e.range.getHeight() > 1)){
    var f = e.range.getCell(1,1).getFormula();
    if(regExpCrowd.test(f)){
      var s = e.source;
      for(var i=1;i<=e.range.getHeight();i++){
        for(var j=1;j<=e.range.getWidth();j++){
          var r = e.range.getCell(i,j);
          crowdTrigger({"range":r,"source":s});
        }
      }
    }
  }
  else{
    var formula = String(e.range.getFormula());
    if(regExpCrowd.test(formula)) crowdTrigger(e);
    if(regExpCoverage.test(formula)) coverageTrigger(e);
    //if(formula == "=getMessages()") loadMessages(e.range.getColumn(),e.range.getRow(),e.source.getId(),e.range.getSheet().getName(),true,false);
  }
}

function coverageTrigger(e){
  var asking = PropertiesService.getUserProperties().getProperty("ASKING_QUESTION");
  if(asking == 1){
    PropertiesService.getUserProperties().setProperty("RANDOM_DATA_INSERTED",1);
  }
  var isNumber = true;
  var range = SpreadsheetApp.getActiveSheet().getRange(e.range.getFormula().slice(10,-1)).getValues();
  if (range == null) return true;
  for(var i=0;i<range.length;i++){
    if(isNumber){
      for(var j=0;j<range[i].length;j++){
        if(isNaN(parseInt(range[i][j]))){
          isNumber = false;
          break;
        }
      }
    }
    else break;
  }
    
  var posValues = [];
  if(isNumber){
    var max = range[0][0];
    var min = range[0][0];
    var isZero = false;
    for(var i=0;i<range.length;i++){
      for(var j=0;j<range[i].length;j++){
        if(range[i][j]>max) max = range[i][j];
        if(range[i][j]<min) min = range[i][j];
        if(range[i][j] == 0) isZero = true;
      }
    }

    if(!isZero) posValues.push(0);
    if(min>=0) posValues.push(-1);
    else posValues.push(min-Math.floor(Math.random()*100));      
    if(max<=0) posValues.push(1);
    else posValues.push(max+Math.floor(Math.random()*100));
    posValues.push(min+Math.floor(Math.random()*(max-min)));
    posValues.push("unknown");
    posValues.push("not aplicable");
    //return posValues[Math.floor(Math.random() * posValues.length)];
  }
  else{
    var maxLength = range[0][0].length;
    var minLength = range[0][0].length;
    var isEmpty = false;
    for(var i=0;i<range.length;i++){
      for(var j=0;j<range[i].length;j++){
        if(String(range[i][j]).length>maxLength) maxLength = String(range[i][j]).length;
        if(String(range[i][j]).length<minLength) minLength = String(range[i][j]).length;
        if(range[i][j] == "") isEmpty = true;
      }
    }
    if(!isEmpty) posValues.push("");
    
    var randomString = "";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    for( var i=0; i <= maxLength; i++ )
        randomString += possible.charAt(Math.floor(Math.random() * possible.length));
    posValues.push(randomString);

    if(minLength > 1){
      randomString = "";
      for( var i=0; i < (minLength-1); i++ )
        randomString += possible.charAt(Math.floor(Math.random() * possible.length));
      posValues.push(randomString);
    }
    
    posValues.push("unknown");
    posValues.push("not aplicable");

    //return posValues;
    //return posValues[Math.floor(Math.random() * posValues.length)]; 
  }
  var initialRow = e.range.getRow();
  var initialCol = e.range.getColumn();
  //e.range.setValue("initialCol: "+initialCol+" initialRow: "+initialRow);
  for(var i=0;i<posValues.length;i++){
    //SpreadsheetApp.getActiveSheet().getRange(e.range.getRow(), e.range.getColumn()).setValue("hola");
    if(i>0){
      var r = SpreadsheetApp.getActiveSheet().getRange(initialRow+i, initialCol);
      if(r.getValue()==""){
        r.setValue(posValues[i]);
      }
      else{
        break;
      }
    }
    else{
      e.range.setValue(posValues[i]);
    }
  }
}

function coverage(range){ 
  return "Generating random data...";  
}

function getRangeHtml(range){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = ss.getActiveSheet().getRange(range);
  // Optional init, to ensure the spreadsheet config overrides the script's
  var conv = SheetConverter.init(ss.getSpreadsheetTimeZone(),
                                 ss.getSpreadsheetLocale());
  // Grab an array for formatted content
  var array = conv.convertRange(r);
  // Get a html table version, with all formatting
  var html = conv.convertRange2html(r);
  return html;
}

function generateBlobFromImage(img){  
  var contentType = "image/png";
  var sliceSize = 512;
  var byteCharacters = Utilities.base64Decode(img.slice(22));
  var blob = Utilities.newBlob(byteCharacters, "image/png", "imageProba");
  return blob;
}