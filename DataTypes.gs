function DataType(){
  this.getName = function(){
    return this.name;
  }
  this.getSubtypes = function(){
    var list = [];
    for(var i=0;i<TYPE_LIST.length;i++){
      if((TYPE_LIST[i] != this.name)&&(eval("new "+TYPE_LIST[i]+"Type()") instanceof eval(this.name+"Type"))) list.push(TYPE_LIST[i]); 
    }
    return list;
  }
}

function FunctionType(){
  this.name = "Function";
  this.matchCell = function(value){
    var regExp = /[^\s\(]+\(.*\)/i;
    return regExp.test(value);
  }
  this.extractRangesFromCrowd = function(expression,sheet){
    var regExp = /question\([^\)]*\)/gi;
    var rangesRegExp = /((('[\w\s]+'!)|(\w+!))?(([A-z]{1,2}\d+(:[A-z]{1,2}\d+)?)|([A-z]{1,2}:[A-z]{1,2})|(\d+:\d+))(;(('[\w\s]+'!)|(\w+!))?(([A-z]{1,2}\d+(:[A-z]{1,2}\d+)?)|([A-z]{1,2}:[A-z]{1,2})|(\d+:\d+)))*)(?=\))/g;
    var rangeList = [];
    var crowdF = String(expression.match(regExp)[0]);
    var lag = crowdF.match(rangesRegExp);
    if(lag!=null){
      var a = lag[0].split(';');
      for(var i=0;i<a.length;i++){
        var h = a[i].split("!");
        if(h.length==2){
          if(h[0].charAt(0) == "'"){
            rangeList.push({"sheet":h[0].slice(1,-1),"range":h[1]});
          }
          else{
            rangeList.push({"sheet":h[0],"range":h[1]});
          }
        }
        else{
          rangeList.push({"sheet":sheet,"range":a[i]});
        }
      }
    }
    return rangeList;
  }
  this.extractRanges = function(expression,sheet){
    var a = /(('[\w\s]+'!)|(\w+!))?((\$?[A-z]{1,2}\$?\d+(:\$?[A-z]{1,2}\$?\d+)?)|(\$?[A-z]{1,2}:\$?[A-z]{1,2})|(\$?\d+:\$?\d+))/g;
    var b = expression.match(a);
    var rangeList = [];
    for(var i=0;i<b.length;i++){
      var r = SpreadsheetApp.getActiveSpreadsheet().getRange(b[0]);
      var c = b[i];
      var d = c.split("!");
      var e = {};
      var f;
      if(d.length > 1){
        if(d[0].charAt(0) == "'") e["sheet"] = d[0].slice(1,-1);
        else e["sheet"] = d[0];
        e["range"] = d[1];
      }
      else{
        e["sheet"] = sheet;
        e["range"] = d[0];
      }
      rangeList.push(e);
    }
    return rangeList;
  }
  this.matchRanges = function(crowdRanges,replyRanges){
    for(var i=0;i<replyRanges.length;i++){
      var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(replyRanges[i].sheet);
      var r = s.getRange(replyRanges[i].range);
      var z = false;
      for(var j=0;j<crowdRanges.length;j++){
        var a = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(crowdRanges[j].sheet);
        var b = a.getRange(crowdRanges[j].range);
        if((crowdRanges[j].sheet==replyRanges[i].sheet)&&(r.getRow()>=b.getRow())&&(r.getLastRow()<=b.getLastRow())&&(r.getColumn()>=b.getColumn())&&(r.getLastColumn()<=b.getLastColumn())){
          z = true;
          break;
        }
        if((crowdRanges[j].sheet==replyRanges[i].sheet)&&(r.getRow()>=b.getRow())&&(r.getLastRow()<=b.getLastRow())&&(r.getColumn()==1)){ // rangos de fila (por mejorar)
          var c = replyRanges[i].range.split("!");
          if(c[c.length-1].search(/[A-z]/g)==-1) z = true;
          break;
        }
        if((crowdRanges[j].sheet==replyRanges[i].sheet)&&(r.getColumn()>=b.getColumn())&&(r.getLastColumn()<=b.getLastColumn())&&(r.getRow()==1)){ // rangos de columna (por mejorar)
          //if(String(r.getRow()).
          var c = replyRanges[i].range.split("!");
          if(c[c.length-1].search(/\d/g)==-1) z = true;
          break;
        }
      }
      if(!z) return false;
    }
    return true;
  }
  this.extract = function(message,problem){
    //var rangeListCrowd = this.extractRangesFromCrowd(problem.expression,problem.sheetName);
    
    var regExp
    
    
    var regExp = /[^\=\s\(]+\(.*\)/i;    
    var res = regExp.exec(message);
    if(res!=null){
      /*if(rangeListCrowd.length != 0){
        var a = this.extractRanges(res[0],problem.sheetName);
        if(!this.matchRanges(rangeListCrowd,a)) return "";
      } */ 
      return {"data":res[0],"rest":message.replace(res[0],"")};
    }
    return "";
  }
  this.handleSolution = function(expression,solution,range,problem){
    var regExp = /question\([^\)]*\)/gi;
    var la = expression.replace(regExp,solution);
    var col = range.getColumn();
    var row = range.getRow();
    if(la.charAt(0) == "=") la = la.slice(1);
    range.setFormula('=HYPERLINK("'+problem.messages[0].link+'";'+la+')');
    var responses = problem.responses;
    var message = "";
    for(var i=0;i<responses.length;i++){
      if((responses[i]["extractedData"]["complete"])&&(responses[i]["extractedData"][row+"-"+col]==solution)){
        message = responses[i].reply;
        break;
      }
    }
    range.setNote("This is a CrowdCall-fed cell. Click to go to the Forum discussion. Do not forget to thank the contributors.\r\n\r\nExtracted from this message:\r\n\r\n"+message);
  }
}
FunctionType.inheritsFrom(DataType);
TYPE_LIST.push("Function");

function StringType(){
  this.name = "String";
  this.matchFormItem = function(value){
    var regExp = /^.+$/;
    var a = regExp.exec(value);
    if(a == null) return null;
    return a[0];
  }
  this.matchCell = function(value){
    var regExp = /.*/;
    return regExp.test(value);
  }
  this.extract = function(message,subproblem,problem){
    var regExp = /\$[^$]+\$/i;
    var res = regExp.exec(message);
    if(res!=null) return {"data":res[0].substring(1,res[0].length-1),"rest":message.replace(res[0],"")};
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;
    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
  this.addItemToForm = function(form,label){
    var item = form.addTextItem();
    item.setTitle(label);
    return item;
  }
}
StringType.inheritsFrom(DataType);
TYPE_LIST.push("String");

function NumberType(){
  this.name = "Number";
  this.matchFormItem = function(value){
    var regExp = /^\d*(\.|\,)?\d+$/;
    var a = regExp.exec(value);
    if(a == null) return null;
    return a[0];
  }
  this.matchCell = function(value){
    var regExp = /^\d*(\.|\,)?\d+$/;
    return regExp.test(value);
  }
  this.extract = function(message,subproblem,problem){
    var regExp = /\$\d*(\.|\,)?\d+\$/i;
    var res = regExp.exec(message);
    if(res!=null) return {"data":res[0].substring(1,res[0].length-1),"rest":message.replace(res[0],"")};
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;
    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
}
NumberType.inheritsFrom(StringType);
TYPE_LIST.push("Number");

function OptionType(options){
  this.options = options;
  this.name = "Option";
  this.matchFormItem = function(value){
    if(this.options.indexOf(value) != -1) return value;
    return null;
  }
  this.matchCell = function(value){
    return false;
  }
  this.extract = function(message,subproblem,problem){
    var foundMatches = [];
    for(var i=0;i<this.options.length;i++){
      var optRegExp = new RegExp(this.options[i],"i");
      if(message.match(optRegExp) != null){
        foundMatches.push(this.options[i]);
      }
    }
    if(foundMatches.length == 1){
      return {"data":foundMatches[0],"rest":message.replace(new RegExp(foundMatches[0],"i"),"")};
    }
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;
    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
  this.addItemToForm = function(form,label){
    var item = form.addListItem();
    item.setTitle(label);
    var choices = [];
    for(var i=0;i<this.options.length;i++){
      choices.push(item.createChoice(this.options[i]));
    }
    item.setChoices(choices);
    return item;
  }
}
OptionType.inheritsFrom(DataType);
TYPE_LIST.push("Option");


function EmailType(){
  this.name = "Email";
  this.matchFormItem = function(value){
    var regExp = /^[A-z0-9\.\_\%\+\-]+\@[A-z0-9\.\-]+\.[A-z]{2,4}$/;
    var a = regExp.exec(value);
    if(a == null) return null;
    return a[0];
  }
  this.matchCell = function(value){
    var regExp = /^[A-z0-9\.\_\%\+\-]+\@[A-z0-9\.\-]+\.[A-z]{2,4}$/;
    return regExp.test(value);
  }
  this.extract = function(message,subproblem,problem){
    var regExp = /[A-z0-9\.\_\%\+\-]+\@[A-z0-9\.\-]+\.[A-z]{2,4}/i;
    var res = regExp.exec(message);
    if(res!=null) return {"data":res[0].substring(1,res[0].length-1),"rest":message.replace(res[0],"")};
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;
    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
}
EmailType.inheritsFrom(StringType);
TYPE_LIST.push("Email");


function URLType(){
  this.name = "URL";
  this.matchFormItem = function(value){
    var regExp = /^(http(s)?:\/\/.)?(www\.)?[-A-z0-9\@\:\%\.\_\+\~\#\=]{2,256}\.[A-z]{2,6}\b([-A-z0-9\@\:\%\_\+\.\~\#\?\&\/\/\=]*)$/gi;

    var a = regExp.exec(value);
    if(a == null) return null;
    return a[0];
  }
  this.matchCell = function(value){
    var regExp = /^(http(s)?:\/\/.)?(www\.)?[-A-z0-9\@\:\%\.\_\+\~\#\=]{2,256}\.[A-z]{2,6}\b([-A-z0-9\@\:\%\_\+\.\~\#\?\&\/\/\=]*)$/gi;
    return regExp.test(value);
  }
  this.extract = function(message,subproblem,problem){
    var regExp = /(http(s)?:\/\/.)?(www\.)?[-A-z0-9\@\:\%\.\_\+\~\#\=]{2,256}\.[A-z]{2,6}\b([-A-z0-9\@\:\%\_\+\.\~\#\?\&\/\/\=]*)/gi;
    var res = regExp.exec(message);
    if(res!=null) return {"data":res[0].substring(1,res[0].length-1),"rest":message.replace(res[0],"")};
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;
    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
}
URLType.inheritsFrom(StringType);
TYPE_LIST.push("URL");


function YYYYMMDDType(){
  this.name = "YYYYMMDD";
  this.matchFormItem = function(value){
    var regExp = /^\d{4}-\d{1,2}-\d{1,2}$/i;
    var a = regExp.exec(value);
    if(a == null) return null;
    var aux = a[0].split("-");
    return aux[0]+"/"+aux[1]+"/"+aux[2];
  }
  this.matchCell = function(value){
    var regExp = /^\d{4}\/\d{1,2}\/\d{1,2}$/i;
    return regExp.test(value);
  }
  this.extract = function(message,subproblem,problem){
    var regExp = /\d{4}\/\d{1,2}\/\d{1,2}/i;
    var res = regExp.exec(message);
    if(res!=null){
      return {"data":res[0],"rest":message.replace(res[0],"")};
    }
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;

    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
  this.addItemToForm = function(form,label){
    var item = form.addDateItem();
    item.setTitle(label);
    return item;
  }
}
YYYYMMDDType.inheritsFrom(StringType);
TYPE_LIST.push("YYYYMMDD");

function DDMMYYYYType(){
  this.name = "DDMMYYYY";
  this.matchFormItem = function(value){
    var regExp = /^\d{4}-\d{1,2}-\d{1,2}$/i;
    var a = regExp.exec(value);
    if(a == null) return null;
    var aux = a[0].split("-");
    return aux[2]+"/"+aux[1]+"/"+aux[0];
  }
  this.matchCell = function(value){
    var regExp = /^\d{1,2}\/\d{1,2}\/\d{4}$/i;
    return regExp.test(value);
  }
  this.extract = function(message,subproblem,problem){
    var regExp = /\d{1,2}\/\d{1,2}\/\d{4}/i;
    var res = message.match(regExp);
    if(res!=null){
      return {"data":res[0],"rest":message.replace(res[0],"")};
    }
    return "";
  }
  this.handleSolution = function(expression,solution,range){
    var regExp;
    if(expression.indexOf("=ask(")==0) regExp = /\=ask\([^\)]*\)/gi;
    else regExp = /ask\([^\)]*\)/gi;

    var la = expression.replace(regExp, solution);
    range.setValue(la);
  }
  this.addItemToForm = function(form,label){
    var item = form.addDateItem();
    item.setTitle(label);
    return item;
  }
}

DDMMYYYYType.inheritsFrom(StringType);
TYPE_LIST.push("DDMMYYYY");