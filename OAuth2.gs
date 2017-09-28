function OAuth2Service(){
  this.authorizationProtocol = "OAuth2";
  this.getAuthorizationProtocol = function(){
    return this.authorizationProtocol;
  }
  this.getService = function(){
     //var l = this.scopes.join(" ");
     return OAuth2.createService(this.serviceName)
     .setAuthorizationBaseUrl(this.authURL)
     .setTokenUrl(this.tokenURL)
     .setClientId(this.clientID)
     .setClientSecret(this.clientSecret)
     .setProjectKey(this.projectKey)
     .setCallbackFunction(this.callbackFunction)
     .setPropertyStore(PropertiesService.getUserProperties())
     .setScope(this.scopes)
     .setParam('access_type', 'offline')
     .setParam('approval_prompt', 'force')
     .setTokenFormat(this.tokenFormat);
  }
  this.isAuthorized = function(){
    var ser = this.getService();
    return ser.hasAccess();
  }
}
OAuth2Service.inheritsFrom(Service);

function installTriggers(){
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('tEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  ScriptApp.newTrigger('tChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  
  ScriptApp.newTrigger('triTime')
    .timeBased()
    .everyHours(1)
    .create();
}