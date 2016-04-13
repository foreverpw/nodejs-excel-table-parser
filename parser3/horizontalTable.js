var Table = require('./table');
var util = require('util');
var XLSX = require('xlsx');

function HorizontalTable() {
	Table.apply(this, arguments);

	this.parse = function(worksheet, range) {
		for (var C = range.s.c; C <= range.e.c; C++) {
			var address = XLSX.utils.encode_cell({
				c : C,
				r : range.s.r
			});
			var cell = worksheet[address];
			this.headers.push(cell.w.trim());
		}
		for (var R = range.s.r + 1; R <= range.e.r; R++) {
			var record = {};
			for (var C = range.s.c; C <= range.e.c; C++) {
				var address = XLSX.utils.encode_cell({
					c : C,
					r : R
				});
				var cell = worksheet[address];
				if (cell != null) {
					record[this.headers[C - range.s.c]] = cell.w.trim();
				}
			}
			this.records.push(record);
		}
	}
}

util.inherits(HorizontalTable, Table);

module.exports = HorizontalTable;