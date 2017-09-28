function OAuth1Service(){
  this.authorizationProtocol = "OAuth1";
  this.getAuthorizationProtocol = function(){
    return this.authorizationProtocol;
  }
  this.getOAuthToken = function(){
    var a = PropertiesService.getUserProperties().getProperty("OAuth1."+this.serviceName);
    if(a!=null){
      var obj = JSON.parse(a);
      return obj.oauthToken;
    }
    return null;
  }
  this.getOAuthSecret = function(){
    var a = PropertiesService.getUserProperties().getProperty("OAuth1."+this.serviceName);
    if(a!=null){
      var obj = JSON.parse(a);
      return obj.oauthSecret;
    }
    return null;
  }
  this.getAuthorization = function(method,baseURL,params,headers){
    var nonce = "";
    var mask = '';
    mask += 'abcdefghijklmnopqrstuvwxyz';
    mask += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    mask += '0123456789';
    for (var i = 32; i > 0; --i) nonce += mask[Math.round(Math.random() * (mask.length - 1))];
    var a = {
      "oauth_consumer_key":this.consumerKey,
      "oauth_signature_method":this.signatureMethod,
      "oauth_timestamp": Math.round(new Date().getTime() / 1000),
      "oauth_nonce": nonce,
      "oauth_version": "1.0"
    };
    if(this.getOAuthToken()!=null) a["oauth_token"] = this.getOAuthToken();
    var b = {};
    if(headers != null){
      for(var i in headers) a[i] = headers[i];
    }
    for(var i in a) b[i] = a[i];
    for(var i in params) b[i] = params[i];
    var keysOrdered = Object.keys(b).sort();
    var c = "";
    for(var i=0;i<keysOrdered.length;i++){
      c += encodeURIComponent(keysOrdered[i]).replace(/!/g, '%21');
      c += "=";
      c += encodeURIComponent(b[keysOrdered[i]]).replace(/!/g, '%21');
      if(i!=(keysOrdered.length-1)) c += "&";
    }
    var d = "";
    d += method.toUpperCase() + "&";
    d += encodeURIComponent(baseURL).replace(/!/g, '%21') + "&";
    d += encodeURIComponent(c).replace(/!/g, '%21');
    
    var signingKey = "";
    signingKey += encodeURIComponent(this.consumerSecret).replace(/!/g, '%21') + "&";
    if(this.getOAuthSecret()!=null) signingKey += encodeURIComponent(this.getOAuthSecret()).replace(/!/g, '%21');
    var e = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, d, signingKey);
    a["oauth_signature"] = Utilities.base64Encode(e);
    var dst = "OAuth "
    for(var i in a){
      if(dst!="OAuth ") dst += ", ";
      dst += encodeURIComponent(i).replace(/!/g, '%21') + '="';
      dst += encodeURIComponent(a[i]).replace(/!/g, '%21') + '"';
    }
    return dst;
  }
  this.getCallbackURL = function(){
    var url = "https://script.google.com/macros/d/" + this.projectKey + "/usercallback?state=";
    var state = ScriptApp.newStateToken()
    .withMethod(this.callbackFunction)
    .withTimeout(3600)
    .createToken();
    return url + state;
  }
  this.getAuthorizationURL = function(){
    var url = this.authorizationURL + "?oauth_token=";
    var oauthToken = this.getOAuthToken();
    return url+oauthToken
  }
  this.isAuthorized = function(){
    var a = PropertiesService.getUserProperties().getProperty("OAuth1."+this.serviceName);
    if(a!=null){
      var obj = JSON.parse(a);
      return obj.authorized;
    }
    return false;
  }
  this.getRequestToken = function(){
    var method = "POST";
    var headers = {"oauth_callback":this.getCallbackURL()};
    var params = {};
    var endpoint = this.requestURL;
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      headers: {
        Authorization: this.getAuthorization(method,endpoint,params,headers)
      }
    });
    var args = response.getContentText().split("&");
    var oauthToken = String(args[0]).split("=")[1];
    var oauthSecret = String(args[1]).split("=")[1];
    var a = {"oauthToken":oauthToken,"oauthSecret":oauthSecret,"authorized":false};
    PropertiesService.getUserProperties().setProperty("OAuth1."+this.serviceName, JSON.stringify(a));
    return true;
  }
  this.getAccessToken = function(verifier){
    var method = "POST";
    var headers = {"oauth_verifier":verifier};
    var params = {};
    var endpoint = this.accessURL;
    var auth = this.getAuthorization(method,endpoint,params,headers);
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      headers: {
        Authorization: auth
      }
    });
    var a = String(response);
    var b = a.split("&");
    var oauthToken = b[0].split("=")[1];
    var oauthSecret = b[1].split("=")[1];
    var c = {"oauthToken":oauthToken,"oauthSecret":oauthSecret,"authorized":true};
    PropertiesService.getUserProperties().setProperty("OAuth1."+this.serviceName, JSON.stringify(c));
    return true;  
  }
}
OAuth1Service.inheritsFrom(Service);

function oauth1Callback(request,serviceName){
  var service = eval("new "+serviceName+"Service()");
  var verifier = request.parameters.oauth_verifier[0];
  var b = service.getAccessToken(verifier);
  var servicesToAuth = (PropertiesService.getUserProperties().getProperty("SERVICES_TO_AUTH") == null) ? [] : eval("(" + PropertiesService.getUserProperties().getProperty("SERVICES_TO_AUTH") + ')');
  servicesToAuth.splice(servicesToAuth.indexOf(serviceName),1);
  PropertiesService.getUserProperties().setProperty("SERVICES_TO_AUTH", JSON.stringify(servicesToAuth));
  authorizeServices();
  return HtmlService.createHtmlOutput('Success! You can close this tab.'); 
}