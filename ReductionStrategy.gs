function ReductionStrategy(){
  this.getName = function(){
    return this.name;
  }
  this.supportsService = function(serviceName){
    if(this.services==null) return true;
    for(var i=0;i<this.services.length;i++){
      if(this.services[i]==serviceName)return true;
    }
    return false;
  }
  this.supportsDataType = function(dataType){
    if(this.dataTypes==null) return true;
    for(var i=0;i<this.dataTypes.length;i++){
      if(this.dataTypes[i]==dataType)return true;
    }
    return false;
  }  
}

function ChooseStrategy(){
  this.name = "Choose";
  this.getBestOption = function(options){
    for(var i=0;i<options.length;i++){
      if(options[i][0]!="") return options[i][0];
    }
    return "";
  }
}
ChooseStrategy.inheritsFrom(ReductionStrategy);
STRATEGY_LIST.push("Choose");

function RepeatStrategy(){
  this.name = "Repeat";
  this.getBestOption = function(options){
    var i;
    var count = {};
    for(i=0;i<options.length;i++){
      if(options[i]!=""){
        if(options[i] in count) count[options[i]] ++;
        else{
          count[options[i]] = 1;
        }
      }
    }
    var key;
    var maxCount = 0;
    var maxKey;
    for(key in count){
      if(count[key]>maxCount){
        maxKey = key;
        maxCount = count[key];
      }
    }
    return maxKey;    
  }
}
RepeatStrategy.inheritsFrom(ReductionStrategy);
STRATEGY_LIST.push("Repeat");