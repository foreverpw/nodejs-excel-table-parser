var leven = require('levenshtein');
var jaro = require('');

var str1 = 'sdafheiohrfkjdnckxm,hfewfnwhuia4y7t8rughifow7sdfjaksdhfudfjgraeihrhurhjrjhehgurehgghregeruguerguerhngiuerguerhguerguierhnguerhguerhng8484yr8hhfdjs';
var str2 = '234rt5mcbvbewiusdjhfdddsnfjkewhriutegfjkdbzxkesdfjnncxjvkbknzmxcbvncbvmnxbvmnxbvmnxcbvmnxcbvuewbiuwhfrio3289436yt78eiufufhsdiufheui2i1479y8treuh';

// var str1 = 'Allen wang';
// var str2 = 'johnny pool';

console.time('leven');
for(var i = 0; i < 1000; i++){
  var dis = new leven(str1,str2);
}
console.timeEnd('leven');

console.time('jaro');
for(var i = 0; i < 1000; i++){
  var dis = jaro(str1,str2);
}
console.timeEnd('jaro');