function TwitterService(){
  this.serviceName = "Twitter";
  this.consumerKey = /*CONSUMER_KEY*/;
  this.consumerSecret = /*CONSUMER_SECRET*/;
  this.accessURL = "https://api.twitter.com/oauth/access_token";
  this.requestURL = "https://api.twitter.com/oauth/request_token";
  this.authorizationURL = "https://api.twitter.com/oauth/authorize"; 
  this.projectKey = PROJECT_KEY;
  this.callbackFunction = "twitterCallback";
  this.signatureMethod = "HMAC-SHA1";
  this.friendlistsSupport = false;
  this.receiverRequired = false;
  this.imageSupport = true;
  this.titleRequired = false;
  this.iconURL = "https://icons.iconarchive.com/icons/danleech/simple/256/twitter-icon.png";
  this.showInShareBar = true;
  this.getName = function(){
    return this.serviceName;
  }
  this.publishImage = function(image){ 
    var boundary = "cuthere";
    
    var requestBody = Utilities.newBlob(
      "--"+boundary+"\r\n"
      + "Content-Disposition: form-data; name=\"media\"; filename=\""+image.getName()+"\"\r\n"
    + "Content-Type: " + image.getContentType()+"\r\n\r\n").getBytes();
    
    requestBody = requestBody.concat(image.getBytes());
    requestBody = requestBody.concat(Utilities.newBlob("\r\n--"+boundary+"--\r\n").getBytes());
    
    try{
      var params = requestBody
      var endpoint = "https://upload.twitter.com/1.1/media/upload.json";
      var method = "POST";
      var auth = this.getAuthorization(method,endpoint,null,null);
      var request = UrlFetchApp.fetch(endpoint, {
        method: method,
        headers: {
          Authorization: auth
        },
        contentType: "multipart/form-data; boundary="+boundary,
        payload: requestBody
      });
      var obj = JSON.parse(request);
      return obj.media_id_string;
    } catch (e) {
      Logger.log(e);
    }
    return "imposible publicar";
  };
  this.publishMessage = function(message,friendlists,images,title,sharedTable){    
    
    var img = [];
 
    /*var regExp = /\=example\(\w{1,2}\d{1,3}\:\w{1,2}\d{1,3}\)/i;
    var matches = message.match(regExp);
    if((matches!=null)&&(matches.length>0)){
      
    }*/
    
    /*
    var regExp = /\[IMAGE\][^\[]+\[\/IMAGE\]/gi;
    var imgs = message.match(regExp);
    
    if(imgs != null){
      var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
      for(var i=0;i<imgs.length;i++){
        var lag = [JSON.parse(imgs[i].slice(7,-8))];
        var imageList = generateImages(lag,spreadsheetId);
        img.push(this.publishImage(imageList[0]));
      }
    }*/
    
    /*var regExp = /\[SHARED_TABLE\][^\[]+\[\/SHARED_TABLE\]/gi;
    var sharedTables = mes.match(regExp);
    if(sharedTables != null){
      for(var i=0;i<sharedTables.length;i++){
        var lag = sharedTables[i].slice(14,-15);
        mes = mes.replace(sharedTables[i],"[Shared table]("+lag+")");
      }
    }*/

    var endpoint = "https://api.twitter.com/1.1/statuses/update.json";
    var method = "POST";
    var params = {"status":message};
    if(img.length>0) params["media_ids"]=img.toString();
    var auth = this.getAuthorization(method,endpoint,params,null);
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      headers: {
        Authorization: auth
      },
      payload: params
    });
    var obj = JSON.parse(response);
    return {"service":"Twitter","messageId":obj.id_str,"authorName":obj.user.name,"authorUsername":"@"+obj.user.screen_name,"userPhoto":obj.user.profile_image_url};
  }
  this.getReplies = function(message){
    var method = "GET";
    var endpoint = "https://api.twitter.com/1.1/statuses/home_timeline.json";
    var auth = this.getAuthorization(method,endpoint,null,null);
    try {
      var result = UrlFetchApp.fetch(endpoint, {
        method: method,
        headers: {
          Authorization: auth
        }
      });
      var obj = JSON.parse(result);
      var responses = [];
      var i;
      for(i=0;i<obj.length;i++){
        if(obj[i].in_reply_to_status_id_str == message.messageId){
          var a = obj[i].text.replace(message.authorUsername,"");
          responses.push({"service":"Twitter","replyId":String(obj[i].id_str),"reply":a,"authorName":obj[i].user.name,"authorUsername":"@"+obj[i].user.screen_name,"authorPhoto":obj[i].user.profile_image_url});
        }
      }
      return responses;
    } catch (e) {
      return [];
    }
  }
  this.replyMessage = function(message,replyText){
    var method = "POST";
    var endpoint = "https://api.twitter.com/1.1/statuses/update.json";
    var mes = "@"+message.authorUsername+" "+replyText;
    var params = {"status":mes,"in_reply_to_status_id":String(message.replyId)};
    var params2 = {"status":encodeURI(mes),"in_reply_to_status_id":String(message.replyId)};
    var auth = this.getAuthorization(method,endpoint,params,null);
    var result = UrlFetchApp.fetch(endpoint, {
      method: method,
      headers: {
        Authorization: auth
      },
      payload: params2
    });
    return true;
    //var obj = JSON.parse(result);
  }
  this.replyMessage2 = function(message,replyText){
    var method = "POST";
    var endpoint = "https://api.twitter.com/1.1/statuses/update.json";
    var mes = "@"+message.authorUsername+" "+replyText;
    var params = {"status":mes,"in_reply_to_status_id":String(message.messageId)};
    var params2 = {"status":encodeURI(mes),"in_reply_to_status_id":String(message.messageId)};
    var auth = this.getAuthorization(method,endpoint,params,null);
    var result = UrlFetchApp.fetch(endpoint, {
      method: method,
      headers: {
        Authorization: auth
      },
      payload: params2
    });
    return true;
    //var obj = JSON.parse(result);
  };
}
TwitterService.inheritsFrom(OAuth1Service);
SERVICE_LIST.push(new TwitterService());

function twitterCallback(request){
  return oauth1Callback(request,"Twitter");
}

function showDialogTwitter(){
  showDialog("Twitter");
}

function selectTwitter(){
  var up = PropertiesService.getUserProperties();
  var service = up.setProperty("SERVICE_TO_USE","Twitter");
  reloadMenu();
}