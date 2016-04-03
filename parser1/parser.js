var xlsx = require('node-xlsx');
var fs = require('fs');
var HorizontalTable = require('./horizontalTable');
var VerticalTable = require('./verticalTable');

/*var obj = xlsx.parse(fs.readFileSync('test.xlsx')); 
console.log(obj);*/




module.exports.parse = function(path,callback){
  fs.readFile(path, function (err,data) {
    if (err) {    
      return console.log(err);
    }
    var sheets = xlsx.parse(data);
    var resultTables = [];
    for(var i in sheets){
      var tables = [];
      var sheet = sheets[i];
      for(var j in sheet.data){
        var lineData = sheet.data[j];
        if(j==0&&lineData[0]&&lineData[0]!=''){
          var table = new HorizontalTable(0);
          tables.push(table);
        }
        for(var m in tables){
          tables[m].parseExcelLine(lineData);
        }
        for(var k in lineData){
          if(lineData[k]&&lineData[k]=='$table_vertical'){
            tables.push(new VerticalTable(k));
          }else if(lineData[k]&&lineData[k]=='$table'){
            tables.push(new HorizontalTable(k));
          }
        }
      }
      resultTables.push(tables);
      callback(resultTables);
    }
  });     
};