var Table = require('./table');
var util = require('util');
var XLSX = require('xlsx');

function VerticalTable(){
  Table.apply(this,arguments);

  this.parse = function(worksheet,range){
    for(var R = range.s.r;R<=range.e.r;R++){
		var address = XLSX.utils.encode_cell({c:range.s.c, r:R});
		var cell = worksheet[address];
		this.headers.push(cell.w.trim());
	}
	for(var C = range.s.c+1; C <= range.e.c; C++){
		var record = {};
		for(var R = range.s.r;R<=range.e.r;R++){
			var address = XLSX.utils.encode_cell({c:C, r:R});
			var cell = worksheet[address];
			record[this.headers[R-range.s.r]] = cell.w.trim();
		}
		this.records.push(record);
	}
  }
}

util.inherits(VerticalTable, Table);

module.exports = VerticalTable;