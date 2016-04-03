// var parser = require('./parser1/parser');
var parser = require('./parser2/parser');

parser.parse(__dirname + '/test2.xlsx',function(tables){
  console.log(tables);
});