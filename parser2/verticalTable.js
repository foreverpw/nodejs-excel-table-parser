var Table = require('./table');
var util = require('util');

function VerticalTable(startXIndex,startYIndex){
  Table.apply(this,arguments);

  var lineSize = 0;

  this.parseExcelLine = function(data){
    if(this.end){
      return;
    }
    if(!this.headers||this.headers.length==0){
      for(var i = this.startXIndex; i < data.length; i++){
        if(!data[i]||data[i]==''){
          break;
        }
      }
      lineSize = i - this.startXIndex;
      for(var j =0;j<lineSize-1; j++){
        var line = [];
        this.tableData.push(line);
      }
    }
    var newData = data.slice(this.startXIndex,this.startXIndex+lineSize);
    if(!newData[0]||newData[0]==''){
      this.end = true;
      return;
    }
    for(var i in newData){
      if(i==0){
        this.headers.push(newData[0]);
      }else{
        this.tableData[i-1].push(newData[i]);
      }
    }
    
  }
  
  this.getJson = function(){
    
  }    
}

util.inherits(VerticalTable, Table);

module.exports = VerticalTable;