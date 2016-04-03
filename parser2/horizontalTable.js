var Table = require('./table');
var util = require('util');

function HorizontalTable(startXIndex,startYIndex){
  Table.apply(this,arguments);
  
  var columnSize = 0;
  
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
      columnSize = i - this.startXIndex;
      this.headers = data.slice(this.startXIndex,this.startXIndex+columnSize);
      return;
    }
    var newData = data.slice(this.startXIndex,this.startXIndex+columnSize);
    var lastLine = true;
    for(var i in newData){
      if(newData[i]&&newData[i]!=''){
        lastLine = false;
      }
    }
    if(lastLine){
      this.end = true;
      return;
    }
    this.tableData.push(newData);
  };
}

util.inherits(HorizontalTable, Table);

module.exports = HorizontalTable;