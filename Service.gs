function Service(){
  this.getAuthorizationProtocol = function(){
    return this.authorizationProtocol;
  }
  this.getFriendlistsSupport = function(){
    return this.friendlistsSupport;
  }
  this.isReceiverMandatory = function(){
    return this.receiverMandatory;
  }
  this.getName = function(){
    return this.serviceName;
  }
}

function showDialog(serviceName){ 
  var service = getService(serviceName);
  if(service!=null){
    if(service.getAuthorizationProtocol() == "OAuth2"){
      var ser = service.getService();
      if(!ser.hasAccess()){
        var authorizationUrl = ser.getAuthorizationUrl();
        var template = HtmlService.createTemplate(
          '<a onclick="setTimeout(function(){google.script.host.close()},1000);" href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
          'Refresh the page when authorization complete.');
        template.authorizationUrl = authorizationUrl;
        var page = template.evaluate();
        SpreadsheetApp.getActive().show(page);
      }
    }
    else if(service.getAuthorizationProtocol() == "OAuth1"){
      if(!service.isAuthorized()){
        var m = service.getRequestToken();
        var authorizationUrl = service.getAuthorizationURL();
        var template = HtmlService.createTemplate(
          '<a onclick="setTimeout(function(){google.script.host.close()},1000);" href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
          'Refresh the page when authorization complete.');
        template.authorizationUrl = authorizationUrl;
        var page = template.evaluate();
        SpreadsheetApp.getActive().show(page);
      }
    }
  }
  else{
    var template = HtmlService.createTemplate(
      'Error. Service not found. Try refreshing the page.');
    var page = template.evaluate();
    SpreadsheetApp.getActive().show(page);
  }
}

function authCallback(request,serviceName){
  var s = getService(serviceName);
  if(s!=null){
    var service = eval("new "+serviceName+"Service().getService()");
    
    var isAuthorized;
    var code = request.parameter.code;
    var error = request.parameter.error;
    if (error) {
      if (error == 'access_denied') {
        return false;
      } else {
        throw 'Error authorizing token: ' + error;
      }
    }
        
    if(serviceName == "StackOverflow"){

      var so = new StackOverflowService();
      var redirectUri = "https://script.google.com/macros/d/"+PROJECT_KEY+"/usercallback";
      
      var response = UrlFetchApp.fetch("https://stackexchange.com/oauth/access_token", {
        method: 'POST',
        payload: {
          code: code,
          client_id: new String(so.clientID),
          client_secret: so.clientSecret,
          redirect_uri: redirectUri
        },
        contentType: 'application/x-www-form-urlencoded',
        muteHttpExceptions: true
      });   
      
      var token = service.parseToken_(response.getContentText());
      if (response.getResponseCode() != 200) {
        isAuthorized = false;
        var reason = token.error ? token.error : response.getResponseCode();
      }
      service.saveToken_(token);
      isAuthorized = true;
    }
    else{
      isAuthorized = service.handleCallback(request);
    }
    if (isAuthorized) {
      var servicesToAuth = (PropertiesService.getUserProperties().getProperty("SERVICES_TO_AUTH") == null) ? [] : eval("(" + PropertiesService.getUserProperties().getProperty("SERVICES_TO_AUTH") + ')');
      servicesToAuth.splice(servicesToAuth.indexOf(serviceName),1);
      PropertiesService.getUserProperties().setProperty("SERVICES_TO_AUTH", JSON.stringify(servicesToAuth));
      authorizeServices();
      return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
      var problemList = (userProperties.getProperty("CROWDCALL_PROBLEM_LIST") == null) ? [] : eval("(" + userProperties.getProperty("CROWDCALL_PROBLEM_LIST") + ')');
      if(problemList.length>0){
        var problem = problemList[0];
        var a = HtmlService.createTemplateFromFile('sidebar');
        a.col = problem.subproblems[0].column;
        a.row = problem.subproblems[0].row;
        a.spreadsheetId = problem.spreadsheetId;
        a.sheetName = problem.sheetName;
        a.expression = problem.subproblems[0].expression;
        a.service = problem.service;
        a.message = problem.subproblems[0].message;
        a.type = problem.subproblems[0].subproblemType;
        a.title = problem.subproblems[0].title;
        a.receivers = null;
        var b = a.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(300);
        SpreadsheetApp.getUi().showSidebar(b);
      }
      PropertiesService.getUserProperties().deleteProperty("CROWDCALL_PROBLEM_TO_CREATE_LIST");
      return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
  }
  else{
    return HtmlService.createHtmlOutput('Error. Service not found');
  }
}

function generateTitle(message){
  var title = "Google Spreadsheet formula ";
  var regExp = /what|which|how|where/i;
  var whPos = message.search(regExp);
  if(whPos != -1){
    var lag = message.slice(whPos).search(" ");
    if(lag == -1) return title;
    whPos += lag+1;
    var pointPos = message.slice(whPos).search(/\.|\?/i);
    if((pointPos != -1)&&(pointPos < 20)) title += message.slice(whPos,whPos+pointPos-1);
    else title += message.slice(whPos,whPos+30);
  }
  return title;
}
