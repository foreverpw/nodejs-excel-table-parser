var Table = require('./table');
var util = require('util');

function HorizontalTable(startHorizontalIndex){
  Table.call(this)
  
  var startHorizontalIndex = startHorizontalIndex;
  var columnSize = 0;
  
  this.parseExcelLine = function(data){
    if(this.end){
      return;
    }
    if(!this.headers||this.headers.length==0){
      for(var i = startHorizontalIndex; i < data.length; i++){
        if(!data[i]||data[i]==''){
          break;
        }
      }
      columnSize = i - startHorizontalIndex;
      this.headers = data.slice(startHorizontalIndex,startHorizontalIndex+columnSize);
      return;
    }
    var newData = data.slice(startHorizontalIndex,startHorizontalIndex+columnSize);
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
  }
}

util.inherits(HorizontalTable, Table);

module.exports = HorizontalTable;