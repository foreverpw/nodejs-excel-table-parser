var Table = require('./table');
var util = require('util');

function VerticalTable(startHorizontalIndex){
  Table.call(this)
  
  var startHorizontalIndex = startHorizontalIndex;
  var lineSize = 0;

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
      lineSize = i - startHorizontalIndex;
      for(var j =0;j<lineSize-1; j++){
        var line = [];
        this.tableData.push(line);
      }
    }
    var newData = data.slice(startHorizontalIndex,startHorizontalIndex+lineSize);
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
}

util.inherits(VerticalTable, Table);

module.exports = VerticalTable;