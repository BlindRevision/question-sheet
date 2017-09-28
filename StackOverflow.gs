function StackOverflowService(){
  this.serviceName = "StackOverflow";
  this.authURL = "https://stackexchange.com/oauth";
  this.tokenURL = "https://stackexchange.com/oauth/access_token";
  this.clientID = /*CLIENT_ID*/;
  this.clientSecret = /*CLIENT_SECRET*/;
  this.projectKey = PROJECT_KEY;
  this.callbackFunction = "stackOverflowCallback";
  this.key = /*KEY*/;
  this.scopes = "write_access read_inbox no_expiry private_info";
  this.titleRequired = true;
  this.iconURL = "http://blog.grio.com/wp-content/uploads/2012/09/stackoverflow.png";
  this.showInShareBar = true;
  this.imageSupport = true;
  this.friendlistsSupport = false;
  this.receiverRequired = false;
  this.getReputation = function (){
    var service = this.getService();
    var accessToken = service.getAccessToken();
    var response = UrlFetchApp.fetch('https://api.stackexchange.com/2.2/me?site=stackoverflow&access_token='+accessToken+'&key='+this.key);
    var obj = JSON.parse(response);
    return obj.items[0].reputation;
  };
  this.editQuestion = function(questionId,newTitle,newMessage){
    var service = this.getService();
    var endpoint = "https://api.stackexchange.com/2.2/questions/"+questionId+"/edit";
    var method = "POST";
    var params = {
      site:"stackoverflow",
      title:title,
      body: mes,
      tags: "google-spreadsheet;formula;spreadsheet",
      key: this.key,
      access_token: service.getAccessToken()
    };
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      payload: params,
      muteHttpExceptions: true
    });
    var obj = JSON.parse(response);
    if(obj.error_message != null){
      return {"error":obj.error_message};
    }
    else{
      var a = obj.items[0];
      return {"service":"StackOverflow","link":a.link,"messageId":a.question_id,"authorName":a.owner.display_name,"userPhoto":a.owner.profile_image};
    }

  }
  this.publishMessage = function(message,friendlists,images,title,sharedTable){
    Logger.log("publish message");
    var service = this.getService();
    var accessToken = service.getAccessToken();
    var endpoint = "https://api.stackexchange.com/2.2/questions/add";
    var method = "POST";
    
    var mes = message;
    
    var reputation = this.getReputation();

    var regExp = /\=screenshot\(\w{1,2}\d{1,3}\:\w{1,2}\d{1,3}\)/i;
    var sampleRanges = message.match(regExp);
    if((sampleRanges != null)&&(sampleRanges.length>0)){
      var sampleRangeString = String(sampleRanges[0]).slice(12,-1);
      var sampleRange = SpreadsheetApp.getActiveSheet().getRange(sampleRangeString);
      if((sampleRange.getHeight()>3)||(sampleRange.getWidth()>3)){
        if(reputation > 10){
          var imageURL = images[0];//uploadImageToImgur(images[0]);
          var lagMes = "![Table]("+imageURL+")";
          lagMes += "\r\n\r\n[Shared table]("+sharedTable+")";
          mes = mes.replace(sampleRanges[0],lagMes);
        }
        else{
          var imageURL = images[0];//uploadImageToImgur(images[0]);
          var lagMes = "[Table]("+imageURL+")";
          lagMes += "\r\n\r\n[Shared table]("+sharedTable+")";
          mes = mes.replace(sampleRanges[0],lagMes);     
        }
      }
      else{
        var sheetName = SpreadsheetApp.getActiveSheet().getName();
        var rangeList = [{range: sampleRangeString, sheet: sheetName}];
        var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
        var lagMes = generateRangeMarkdown(rangeList,spreadsheetId);
        lagMes += "\r\n\r\n[Shared table]("+sharedTable+")";
        mes = mes.replace(sampleRanges[0],lagMes);     
      }
    }
    var params = {
      site:"stackoverflow",
      title:title,
      body: mes,
      tags: "google-spreadsheet;formula;spreadsheet",
      key: this.key,
      access_token: accessToken
    };
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      payload: params,
      muteHttpExceptions: true
    });
    var obj = JSON.parse(response);
    Logger.log(obj);
    if(obj.error_message != null){
      return {"error":obj.error_message};
    }
    else{
      var a = obj.items[0];
      return {"service":"StackOverflow","link":a.link,"messageId":a.question_id,"authorName":a.owner.display_name,"userPhoto":a.owner.profile_image};
    }
  };
  this.getReplies = function(message){
    var endpoint = "https://api.stackexchange.com/2.2/questions/"+message.messageId+"/answers?site=stackoverflow&filter=!3yXvhCa-PLfHWn._i";
    var method = "GET";
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
    });
    var obj = JSON.parse(response);
    var responses = [];
    var i;
    for(i=0;i<obj.items.length;i++){
      var a = obj.items[i];
      responses.push({"service":"StackOverflow","replyId":a.answer_id,"votes":a.score,"authorReputation":a.owner.reputation,"reply":a.body,"authorName":a.owner.user_id,"authorUsername":a.owner.display_name,"authorPhoto":a.owner.profile_image});
    }
    return responses;
  };
  this.replyMessage = function(message,replyText){
    var service = this.getService();
    var endpoint = "https://api.stackexchange.com/2.2/posts/"+message.messageId+"/comments/add";
    var method = "POST";
    var params = {
      site:"stackoverflow",
      body: replyText,
      key: this.key,
      access_token: service.getAccessToken()
    };
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      payload: params
    });
    var obj = JSON.parse(response);
    Logger.log(obj);
    return true;
  }
  this.replyMessage2 = function(message,replyText){
    var service = this.getService();
    var endpoint = "https://api.stackexchange.com/2.2/posts/"+message.messageId+"/comments/add";
    var method = "POST";
    var params = {
      site:"stackoverflow",
      body: replyText,
      key: this.key,
      access_token: service.getAccessToken()
    };
    var response = UrlFetchApp.fetch(endpoint, {
      method: method,
      payload: params
    });
    var obj = JSON.parse(response);
    Logger.log(obj);
    return true;
  }
}
StackOverflowService.inheritsFrom(OAuth2Service);
SERVICE_LIST.push(new StackOverflowService());

function searchRelatedQuestions(title){
  var endpoint = "https://api.stackexchange.com/2.2/similar/?site=stackoverflow&sort=relevance&title="+encodeURIComponent(title);
  var method = "GET";
  var response = UrlFetchApp.fetch(endpoint, {
    method: method
  });
  var obj = JSON.parse(response);
  var questionList = [];
  for(var i=0;i<obj.items.length;i++){
    questionList.push({"title":obj.items[i].title,"link":obj.items[i].link});
  }
  /*var a = HtmlService.createTemplateFromFile('questionList');
  a.questionList = questionList;
  a.form = form;
  var b = a.evaluate().setTitle("CrowdCall").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(b); 
  return true;*/
  return questionList;
}

function stackOverflowCallback(request) {
  return authCallback(request,"StackOverflow");
}

function showDialogStackOverflow(){
  showDialog("StackOverflow");
}

function selectStackOverflow(){
  var up = PropertiesService.getUserProperties();
  var service = up.setProperty("SERVICE_TO_USE","StackOverflow");
  reloadMenu();
}

function uploadImageToImgur(image){
  //var imageURL = PropertiesService.getUserProperties().getProperty("IMAGE_URL");
  //if(imageURL != null) return imageURL;
  var boundary = "cuthere";
  
  var requestBody = Utilities.newBlob(
    "--"+boundary+"\r\n"
    + "Content-Disposition: form-data; name=\"fkey\"\r\n\r\n"
    + "ed3fe1ec9f11892a684041b2e64fe826"+"\r\n"
    +"--"+boundary+"\r\n"
    + "Content-Disposition: form-data; name=\"source\"\r\n\r\n"
    + "computer"+"\r\n"
    +"--"+boundary+"\r\n"
    + "Content-Disposition: form-data; name=\"filename\"; filename=\""+image.getName()+"\"\r\n"
  + "Content-Type: " + image.getContentType()+"\r\n\r\n").getBytes();
  
  requestBody = requestBody.concat(image.getBytes());
  requestBody = requestBody.concat(Utilities.newBlob("\r\n--"+boundary+"--\r\n").getBytes());
  
  var params = requestBody
  var endpoint = "https://stackoverflow.com/upload/image?https=true";
  var method = "POST";
  
  var request = UrlFetchApp.fetch(endpoint, {
    method: method,
    contentType: "multipart/form-data; boundary="+boundary,
    payload: requestBody
  });
  
  var regExp = /http[^"]+"/i;
  Logger.log(request.getContentText());
  var lag = request.getContentText().match(regExp);
  if(lag == null) return {"error":"no image found"};
  var url = String(lag[0].slice(0,-1));
  //Logger.log("URL - "+url);
  //PropertiesService.getUserProperties().setProperty("IMAGE_URL", url);
  return url;
}

function generateRangeMarkdown(rangeList,spreadsheet){
  var ret = "\r\n";
  for(var i=0;i<rangeList.length;i++){
    var ss = SpreadsheetApp.openById(spreadsheet);
    var sheet = ss.getSheetByName(rangeList[i].sheet);
    var range = sheet.getRange(rangeList[i].range).getValues();
    var aux = rangeList[i].range.match(/[\w\$]+(?=:)/)[0];
    var a = sheet.getRange(aux);
    var firstC = a.getColumn()-1;
    var dataRange = Charts.newDataTable();
    var width = range[0].length;
    var firstColumn = rangeList[i].range.match(/^\D+(?=\d)/)[0];
    var firstRow = rangeList[i].range.match(/\d+(?=:)/)[0];
    var columnIndexes = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var cc = columnIndexes.split("");
    ret += "    |    ";
    var maxColLength = [];
    for(var j=0;j<range.length;j++){
      for(var k=0;k<range[j].length;k++){
        var le = range[j][k].toString().length;
        if(maxColLength.length <= k){
          if(le>4) maxColLength[k] = le;
          else maxColLength[k] = 4;
        }
        else if(maxColLength[k] < le){
          maxColLength[k] = le;
        }
      }
    }
    
    for(var j=0;j<range[0].length;j++){
      var numWhites;
      if((firstC + j) > 25){
        var p1=cc[Math.floor((firstC+j)/26)-1];
        var p2=cc[(firstC+j)%26];
        numWhites = maxColLength[j] - 3;
        ret += "| "+p1+p2;
      }
      else{    
        ret += "| "+cc[(firstC+j)];
        numWhites = maxColLength[j] - 2;
      }
      for(var n=0;n<numWhites;n++) ret += " ";
    }
    ret += "|  \r\n";
    
    //ret += "|:--:";
    ret += "    -----";
    for(var j=0;j<maxColLength.length;j++){
      //ret += "|:";
      ret += "--";
      for(var n=1;n<(maxColLength[j]-1);n++) ret += "-";
      //ret += ":";
      ret += "-";
    }
    //ret += "|";
    ret += "-";
    
    for(var j=0;j<range.length;j++){
      ret += "  \r\n";
      var v = parseInt(firstRow)+j;
      ret += "    |"+v;
      for(var n=String(v).length;n<4;n++) ret += " ";
      for (var k=0;k<range[j].length;k++){
        var lag = range[j][k].toString();
        ret += "|"+lag;
        for(var n=lag.length;n<maxColLength[k];n++) ret += " ";
      }
      ret += "|";
    }
  }
  ret += "  \r\n";
  return ret;
}