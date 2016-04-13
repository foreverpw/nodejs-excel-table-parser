function Table(){
  this.headers = [];
  // this.tableData = [];
  this.records = [];
  this.end = false;

  this.getDescription = function(){
	return {totalColumns:this.headers.length,totalItems:this.records.length};
  }  
}

module.exports = Table;