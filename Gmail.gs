function GmailService(){
  this.serviceName = "Gmail";
  this.imageSupport = true;
  this.friendlistsSupport = true;
  this.titleRequired = true;
  this.receiverRequired = true;
  this.showInShareBar = false;
  this.isAuthorized = function(){
    return true;
  }
  this.getContacts = function(){
    //var groups = GroupsApp.getGroups();
    var list = [];
    /*for(var i=0;i<groups.length;i++){
      var g = groups[i];
      list.push({"name":g.getEmail(),"nameToShow":"Grupo - "+g.getEmail(),"id":g.getEmail()});
    }*/
    var contacts = ContactsApp.getContacts();
    for(var i=0;i<contacts.length;i++){
      var g = contacts[i];
      if(g.getEmails().length>0){
        for(var j=0;j<g.getEmails().length;j++){
          if((g.getFullName()!=null)&&(g.getFullName()!="")) list.push({"name":g.getEmails()[j].getAddress(),"fullName":g.getFullName(),"id":g.getEmails()[j].getAddress()});
          else list.push({"name":g.getEmails()[j].getAddress(),"fullName":null,"id":g.getEmails()[j].getAddress()});
        }
      }
    }
    return list;
  };
  this.publishMessage = function(message,friendlists,images,subject,sharedTable){
    var regExp = /\=screenshot\(\w{1,2}\d{1,3}\:\w{1,2}\d{1,3}\)/i;
    var sampleRanges = message.match(regExp);
    var attachmentList = {};
    var mes = message;
    var mesHTML = message;
    
    if((sampleRanges != null)&&(sampleRanges.length>0)){
      var img = UrlFetchApp.fetch(images[0]).getBlob();
      //attachmentList["context"] = images[0];
      attachmentList["context"] = img;
      mesHTML = mesHTML.replace(sampleRanges[0],'<br/><img src="cid:context" /><br/><a href="'+sharedTable+'">Shared table</a><br/>');
    }
        
    mesHTML = mesHTML.replace(/\n/g,"<br/>");
    if(sharedTable == null){
      GmailApp.sendEmail(friendlists.toString(), subject, mes)
    }
    else{
      GmailApp.sendEmail(friendlists.toString(), subject, mes, {
        inlineImages: attachmentList,
        htmlBody: mesHTML
      })
    }
    Utilities.sleep(2000);
    var thread = GmailApp.search("in:sent",0,1)[0];
    var mes = thread.getMessages()[0];
    var userName = mes.getFrom();
    return {"service":"Gmail","threadId":thread.getId(),"messageId":mes.getId(),"authorUsername":mes.getFrom(),"link":thread.getPermalink()};
  };
  this.getReplies = function(message){
    var responseList = [];
    var mes = GmailApp.getThreadById(message.threadId).getMessages();
    for(var i=0;i<mes.length;i++){
      if(mes[i].getId()!=message.messageId){
        var regExp = /\<.+\@.+\>/;
        var address = mes[i].getFrom().match(regExp);
        var a = {"service":"Gmail","originalThreadId":message.threadId,"subject":mes[i].getSubject(),"service":"Gmail","replyId":mes[i].getId(),"reply":mes[i].getPlainBody()};
        if(address != null){
          a["authorUsername"]=address[0].slice(1,-1);
          a["authorName"]=mes[i].getFrom().slice(0,mes[i].getFrom().indexOf(address[0]));
        }
        else a["authorUsername"] = mes[i].getFrom();
        responseList.push(a);
      }
    }
    return responseList;
  }
  this.replyMessage = function(message,replyText){
    var a = GmailApp.getMessageById(message.replyId);
    var mes = a.reply(replyText);
  }
  this.replyMessage2 = function(message,replyText){
    var a = GmailApp.getMessageById(message.messageId);
    var mes = a.replyAll(replyText);   
  }
}
SERVICE_LIST.push(new GmailService());


function selectGmail(){
  var up = PropertiesService.getUserProperties();
  var service = up.setProperty("SERVICE_TO_USE","Gmail");
  reloadMenu();
}