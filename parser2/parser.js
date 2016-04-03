var XLSX = require('xlsx');
var fs = require('fs');
var HorizontalTable = require('./horizontalTable');
var VerticalTable = require('./verticalTable');

var horizontalHeaderColor = 'FFC000';
var verticalHeaderColor = '92D050';

module.exports.parse = function(path,callback){
  fs.readFile(path, function (err,data) {
    if (err) {    
      return console.log(err);
    }
    var resultTables = [];
    var workbook = XLSX.read(data, {type: 'buffer',cellStyles:true});
    var sheet_name_list = workbook.SheetNames;
    sheet_name_list.forEach(function(sheetName) { /* iterate through sheets */
      var worksheet = workbook.Sheets[sheetName];
      var sheetData = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: true});
      var tables = [];
      for (var z in worksheet) {//detect tables
        if(z[0] === '!') continue;
        if(worksheet[z].s&&worksheet[z].s.fgColor.rgb == horizontalHeaderColor){
          var range = XLSX.utils.decode_range(z);
          var previousRange = JSON.parse(JSON.stringify(range));
          previousRange.s.c--;
          var previousCell = worksheet[XLSX.utils.encode_range(previousRange).split(':')[0]];
          if(!previousCell||!previousCell.s||previousCell.s.fgColor.rgb != horizontalHeaderColor){
            tables.push(new HorizontalTable(range.s.c,range.s.r));              
          }
        }else if(worksheet[z].s&&worksheet[z].s.fgColor.rgb == verticalHeaderColor){
          var range = XLSX.utils.decode_range(z);
          var previousRange = JSON.parse(JSON.stringify(range));
          previousRange.s.r--;
          var previousCell = worksheet[XLSX.utils.encode_range(previousRange).split(':')[0]];
          if(!previousCell||!previousCell.s||previousCell.s.fgColor.rgb != verticalHeaderColor){
            tables.push(new VerticalTable(range.s.c,range.s.r));              
          }
        }
      }
      for(var i in sheetData){
        var lineData = sheetData[i];
        for(var j in tables){
          if(tables[j].startYIndex <= i&&!tables[j].end){
            tables[j].parseExcelLine(lineData);
          }
        }
      }
      resultTables.push(tables);
    });
    callback(resultTables);
  });
};

