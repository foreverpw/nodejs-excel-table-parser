var XLSX = require('xlsx');
var fs = require('fs');
var HorizontalTable = require('./horizontalTable');
var VerticalTable = require('./verticalTable');
var path = require('path');
var async = require('async');

module.exports.parse = function(paths, callback) {
	var result = [];
	async.eachSeries(paths, function(filePath, filecb) {
		fs.readFile(filePath, function(err, content) {
			if (!err) {
				var workbook = XLSX.read(content, {
					type : 'buffer',
					cellStyles : true
				});
				var sheet_name_list = workbook.SheetNames;
				console.log(sheet_name_list);
				var resultTables = [];
				async.eachSeries(sheet_name_list, function(sheetName, cb) {
					var tables = parseSheet(workbook.Sheets[sheetName]);
					resultTables.push({
						sheetName : sheetName,
						tables : tables
					});
					cb(null);
				},
				// Final callback after each item has been iterated over.
				function(err) {
					var fileName = path.basename(filePath);
					result.push({
						fileName : fileName,
						sheets : resultTables
					});
					filecb(err);
				});
			}
		});
	},
	// Final callback after each item has been iterated over.
	function(err) {
		callback(result);
	});
};

function parseSheet(worksheet) {
	if (!worksheet['!ref']) {
		return;
	}
	var range = XLSX.utils.decode_range(worksheet['!ref']);

	var tableRanges = [];
	var tables = [];
	console.log('rage');
	console.log(range);
	// horizontal
	for (var R = range.s.r; R <= range.e.r; ++R) {
		// if this line is header;
		for (var C = range.s.c; C <= range.e.c; ++C) {
			var cell_address = XLSX.utils.encode_cell({
				c : C,
				r : R
			});
			var cell = worksheet[cell_address];
			var top_address = XLSX.utils.encode_cell({
				c : C,
				r : R - 1
			});
			var top_cell = worksheet[top_address];
			var left_address = XLSX.utils.encode_cell({
				c : C - 1,
				r : R
			});
			var left_cell = worksheet[left_address];
			if ((!top_cell || top_cell.w.trim() == '')
					&& (!left_cell || left_cell.w.trim() == '') && cell
					&& cell.w.trim() != '') {
				var notInTables = true;
				for ( var i in tableRanges) {
					if (inRange(tableRanges[i], {
						c : C,
						r : R
					})) {
						notInTables = false;
					}
				}
				if (notInTables) {
					console.log('notInTables------');
					var hTableRange = isHorizontalTable(worksheet, {
						c : C,
						r : R
					});
					console.log(hTableRange);
					if (hTableRange) {
						console.log('#find horizontal table on R:' + R + ',C:'
								+ C);
						tableRanges.push(hTableRange);
						var ht = new HorizontalTable();
						ht.parse(worksheet, hTableRange);
						tables.push(ht);
					} else {
						var vTableRange = isVerticalTable(worksheet, {
							c : C,
							r : R
						});
						if (vTableRange) {
							console.log(sheetName);
							console.log('#find vertical table on R:' + R
									+ ',C:' + C);
							tableRanges.push(vTableRange);
							var vt = new VerticalTable();
							vt.parse(worksheet, vTableRange);
							tables.push(vt);
						}
					}
				}
			}
		}

	}
	console.log('notInTables:' + notInTables);
	return tables;
}

function inRange(range, address) {
	if (range.s.c <= address.c && address.c <= range.e.c
			&& range.s.r <= address.r && address.r <= range.e.r) {
		return true;
	} else {
		return false;
	}
}

function isHorizontalTable(worksheet, position) {
	console.log(position);
	var range = XLSX.utils.decode_range(worksheet['!ref']);
	var end_C = position.c;
	for (var C = position.c + 1; C <= range.e.c; ++C) {
		var address = XLSX.utils.encode_cell({
			c : C,
			r : position.r
		});
		var cell = worksheet[address];
		if (!cell || cell.w.trim() == '') {
			break;
		}
		end_C = C;
	}
	var end_R = position.r;
	for (var R = position.r + 1; R <= range.e.r; R++) {
		var tableEnd = true;
		for (var C = position.c; C <= end_C; ++C) {
			var address = XLSX.utils.encode_cell({
				c : C,
				r : R
			});
			var cell = worksheet[address];
			if (cell && cell.w.trim() != '') {
				tableEnd = false;
			}
		}
		if (tableEnd) {
			break;
		}
		end_R = R;
	}
	// special 1
	var width = end_C - position.c + 1;
	var height = end_R - position.r + 1;
	if (height < 3 && width < 2 || height == 1) {
		console.log('return false1');
		return false;
	}
	// special 2
	if (height < 3 && width > 2) {
		// headers are different
		var headerList = [];
		for (var C = position.c; C <= end_C; ++C) {
			var address = XLSX.utils.encode_cell({
				c : C,
				r : position.r
			});
			var cell = worksheet[address];
			for ( var i in headerList) {
				if (cell.w.trim() == headerList[i].trim()) {
					console.log('return false2');
					return false;
				}
			}
			headerList.push(cell.w);
		}
		var tableRange = {
			s : {
				c : position.c,
				r : position.r
			},
			e : {
				c : end_C,
				r : end_R
			}
		}
		return tableRange;
	}
	// iterate next 20/totalRows row between position.c to end_C
	var columns_with_data = [];
	var iterateSize = (end_R - position.r) > 20 ? 20 : end_R - position.r
	for (var R = position.r + 1; R <= position.r + iterateSize; R++) {
		var cols = [];
		for (var C = position.c; C <= end_C; ++C) {
			var address = XLSX.utils.encode_cell({
				c : C,
				r : R
			});
			var cell = worksheet[address];
			if (cell && cell.w.trim() != '') {
				cols.push(C);
			}
		}
		columns_with_data.push(cols);
	}
	var sumOfDistance = 0;
	for (var i = 0; i < columns_with_data.length - 1; i++) {
		var row1 = columns_with_data[i];
		var row2 = columns_with_data[i + 1];
		var distArray = levenshteinenator(row1.join(''), row2.join(''));
		var dist = (distArray[distArray.length - 1][distArray[distArray.length - 1].length - 1])
				/ ((row1.length + row2.length) / 2);
		sumOfDistance += dist;
	}
	var averageDistance = sumOfDistance / (iterateSize - 1);
	// console.log('sumOfDistance:'+sumOfDistance);
	// console.log('averageDistance:'+averageDistance);
	if (averageDistance < 0.3) {
		var sumTotalDiff = 0;
		var ignoreColumns = 0;
		for (var C = position.c; C <= end_C; ++C) {
			var sumColDiff = 0;
			for (var R = position.r + 1; R < position.r + iterateSize; R++) {
				var address1 = XLSX.utils.encode_cell({
					c : C,
					r : R
				});
				var cell1 = worksheet[address1];
				var value1 = cell1 ? cell1.w.trim() : '';
				var address2 = XLSX.utils.encode_cell({
					c : C,
					r : R + 1
				});
				var cell2 = worksheet[address2];
				var value2 = cell2 ? cell2.w.trim() : '';
				var distArray = levenshteinenator(value1, value2);
				if (value1.length > 20) {
					sumColDiff = 200;
					break;
				}
				var dist = (distArray[distArray.length - 1][distArray[distArray.length - 1].length - 1]);
				sumColDiff += dist;
			}
			if (sumColDiff > 100) {
				ignoreColumns++;
			} else {
				sumTotalDiff += sumColDiff / (iterateSize - 1);
			}
		}
		var allCellAverageDiff = sumTotalDiff / (width - ignoreColumns);
		console.log('allCellAverageDiff:'+allCellAverageDiff);
		if (allCellAverageDiff < 5) {
			var tableRange = {
				s : {
					c : position.c,
					r : position.r
				},
				e : {
					c : end_C,
					r : end_R
				}
			}
			return tableRange;
		} else {
			return false;
		}
	}
}

function isVerticalTable(worksheet, position) {
	// console.log(position);
	var range = XLSX.utils.decode_range(worksheet['!ref']);
	var end_R = position.r;
	for (var R = position.r + 1; R <= range.e.r; ++R) {
		var address = XLSX.utils.encode_cell({
			c : position.c,
			r : R
		});
		var cell = worksheet[address];
		if (!cell || cell.w.trim() == '') {
			break;
		}
		end_R = R;
	}
	var end_C = position.c;
	for (var C = position.c + 1; C <= range.e.c; C++) {
		var tableEnd = true;
		for (var R = position.r; R <= end_R; ++R) {
			var address = XLSX.utils.encode_cell({
				c : C,
				r : R
			});
			var cell = worksheet[address];
			if (cell && cell.w.trim() != '') {
				tableEnd = false;
			}
		}
		if (tableEnd) {
			break;
		}
		end_C = C;
	}
	// special 1
	var width = end_R - position.r + 1;
	var height = end_C - position.c + 1;
	if (height < 3 && width < 2 || height == 1) {
		return false;
	}
	// special 2
	if (height < 3 && width > 2) {
		// headers are different
		var headerList = [];
		for (var R = position.r; R <= end_R; ++R) {
			var address = XLSX.utils.encode_cell({
				c : position.c,
				r : R
			});
			var cell = worksheet[address];
			for ( var i in headerList) {
				if (cell.w.trim() == headerList[i].trim()) {
					return false;
				}
			}
			headerList.push(cell.w);
		}
		var tableRange = {
			s : {
				c : position.c,
				r : position.r
			},
			e : {
				c : end_C,
				r : end_R
			}
		}
		return tableRange;
	}
	// iterate next 20/totalRows row between position.c to end_C
	var columns_with_data = [];
	var iterateSize = (end_C - position.c) > 20 ? 20 : end_C - position.c
	for (var C = position.c + 1; C <= position.c + iterateSize; C++) {
		var cols = [];
		for (var R = position.r; R <= end_R; ++R) {
			var address = XLSX.utils.encode_cell({
				c : C,
				r : R
			});
			var cell = worksheet[address];
			if (cell && cell.w.trim() != '') {
				cols.push(R);
			}
		}
		columns_with_data.push(cols);
	}
	var sumOfDistance = 0;
	for (var i = 0; i < columns_with_data.length - 1; i++) {
		var row1 = columns_with_data[i];
		var row2 = columns_with_data[i + 1];
		var distArray = levenshteinenator(row1.join(''), row2.join(''));
		var dist = (distArray[distArray.length - 1][distArray[distArray.length - 1].length - 1])
				/ ((row1.length + row2.length) / 2);
		sumOfDistance += dist;
	}
	var averageDistance = sumOfDistance / (iterateSize - 1);
	// console.log('sumOfDistance:'+sumOfDistance);
	// console.log('averageDistance:'+averageDistance);
	if (averageDistance < 0.3) {
		var sumTotalDiff = 0;
		var ignoreColumns = 0;
		for (var R = position.r; R <= end_R; ++R) {
			var sumColDiff = 0;
			for (var C = position.c + 1; C < position.c + iterateSize; C++) {
				var address1 = XLSX.utils.encode_cell({
					c : C,
					r : R
				});
				var cell1 = worksheet[address1];
				var value1 = cell1 ? cell1.w.trim() : '';
				var address2 = XLSX.utils.encode_cell({
					c : C + 1,
					r : R
				});
				var cell2 = worksheet[address2];
				var value2 = cell2 ? cell2.w.trim() : '';
				var distArray = levenshteinenator(value1, value2);
				if (value1.length > 20) {
					sumColDiff = 200;
					break;
				}
				var dist = (distArray[distArray.length - 1][distArray[distArray.length - 1].length - 1]);
				sumColDiff += dist;
			}
			if (sumColDiff > 100) {
				ignoreColumns++;
			} else {
				sumTotalDiff += sumColDiff / (iterateSize - 1);
			}
		}
		var allCellAverageDiff = sumTotalDiff / (width - ignoreColumns);
		// console.log('allCellAverageDiff:'+allCellAverageDiff);
		if (allCellAverageDiff < 3) {
			var tableRange = {
				s : {
					c : position.c,
					r : position.r
				},
				e : {
					c : end_C,
					r : end_R
				}
			}
			return tableRange;
		} else {
			return false;
		}
	}
}

function levenshteinenator(a, b) {
	var cost;
	var m = a.length;
	var n = b.length;

	// make sure a.length >= b.length to use O(min(n,m)) space, whatever that is
	if (m < n) {
		var c = a;
		a = b;
		b = c;
		var o = m;
		m = n;
		n = o;
	}

	var r = [];
	r[0] = [];
	for (var c = 0; c < n + 1; ++c) {
		r[0][c] = c;
	}

	for (var i = 1; i < m + 1; ++i) {
		r[i] = [];
		r[i][0] = i;
		for (var j = 1; j < n + 1; ++j) {
			cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
			r[i][j] = minimator(r[i - 1][j] + 1, r[i][j - 1] + 1,
					r[i - 1][j - 1] + cost);
		}
	}

	return r;
}

/**
 * Return the smallest of the three numbers passed in
 * 
 * @param Number
 *            x
 * @param Number
 *            y
 * @param Number
 *            z
 * @return Number
 */
function minimator(x, y, z) {
	if (x <= y && x <= z)
		return x;
	if (y <= x && y <= z)
		return y;
	return z;
}
