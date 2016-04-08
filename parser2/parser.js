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
        if(worksheet[z].s&&worksheet[z].s.fgColor.rgb){
          var fgColor = worksheet[z].s.fgColor.rgb;
          var decodedAddress = XLSX.utils.decode_cell(z);
          var topAddress = XLSX.utils.encode_cell({c:decodedAddress.c,r:decodedAddress.r-1});
          var topCell = worksheet[topAddress];
          var leftAddress = XLSX.utils.encode_cell({c:decodedAddress.c-1,r:decodedAddress.r});
          var leftCell = worksheet[leftAddress];
          if((topCell&&topCell.s&&topCell.s.fgColor.rgb==fgColor)||
              (leftCell&&leftCell.s&&leftCell.s.fgColor.rgb==fgColor)){
                
          }else{
            var bottomAddress = XLSX.utils.encode_cell({c:decodedAddress.c,r:decodedAddress.r+1});
            var bottomCell = worksheet[bottomAddress];
            var rightAddress = XLSX.utils.encode_cell({c:decodedAddress.c+1,r:decodedAddress.r});
            var rightCell = worksheet[rightAddress];
            if(rightCell&&rightCell.s&&rightCell.s.fgColor.rgb==fgColor){
              tables.push(new HorizontalTable(decodedAddress.c,decodedAddress.r));  
            }else if(bottomCell&&bottomCell.s&&bottomCell.s.fgColor.rgb==fgColor){
              tables.push(new VerticalTable(decodedAddress.c,decodedAddress.r));              
            }  
          }
        }
        /*if(worksheet[z].s&&worksheet[z].s.fgColor.rgb == horizontalHeaderColor){
          var address = XLSX.utils.decode_cell(z);
          var previousCell = worksheet[XLSX.utils.encode_cell({c:address.c-1,r:address.r})];
          if(!previousCell||!previousCell.s||previousCell.s.fgColor.rgb != horizontalHeaderColor){
            tables.push(new HorizontalTable(address.c,address.r));              
          }
        }else if(worksheet[z].s&&worksheet[z].s.fgColor.rgb == verticalHeaderColor){
          var address = XLSX.utils.decode_cell(z);
          var previousCell = worksheet[XLSX.utils.encode_cell({c:address.c,r:address.r-1})];
          if(!previousCell||!previousCell.s||previousCell.s.fgColor.rgb != verticalHeaderColor){
            tables.push(new VerticalTable(address.c,address.r));              
          }
        }*/
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

