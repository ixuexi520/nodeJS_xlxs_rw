var request = require("request");
// write data to excel 
var fs = require('fs');
var xlsx = require('node-xlsx');

var date = new Date();
var y = date.getFullYear();
var m = date.getMonth() + 1;
var d = date.getDate();

if (m < 10) {
  m = '0' + m;
}
if (d < 10) {
  d = '0' + d;
}

var today = y + '' + m + '' + d;

var stock = [2454, 2317,  2330, 1476, 2002];
var stockIndex = 0;
var num = '',
    name = '',
    price = '';
    titleName = new Array();

var run = function() {
  var stockNo = stock[stockIndex];
  var jsonUrl = "http://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=" + today + "&stockNo=" + stockNo;
  if (stockNo) {
    console.log(stockNo);
    request({
      url: jsonUrl,
      method: "GET"
    }, function(error, response, body) {
      if (error || !body) {
        return;
      } else {
          // 如果沒有資料，會出現 404 的 html 網頁，此時就重新抓取
        if (body.indexOf('html') != -1) {
          console.log('Error !! reload');
          // run2();
        } else {
          b = JSON.parse(body);
          var json = b.data;
          var title = b.title.split(' ');
          //titleName.push('TTT');
//          console.log('title : '+ json);
          titleName.push(title[2]);
          var data = json[json.length - 1];
          num = num + title[1] + ',';
          name = name + title[2] + ',';
          price = price + data[data.length - 3] + ',';
          stockIndex = stockIndex + 1;
          run();
          console.log('one time: '+ titleName);
        }
      }
    });
  } else {
    console.log('T'+num);
    console.log(name);
    console.log('E'+price);
var buffer = xlsx.build([
    {
        name:'sheet1',
        data:stock
    }        
]);
 
//将文件内容插入新的文件中
    //writeXls();
    fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});   //生成excel
    // sheet();
  }
};

var sheet = function() {
  var parameter = {
    sheetUrl: '試算表網址',
    sheetName: '工作表名稱',
    num: num,
    name: name,
    price: price
  }
  request({
    url: 'Google App Script 網址',
    method: "GET",
    qs: parameter
  }, function(error, response, body) {
    console.log(body);
  });
}

function writeXls() {
    var buffer = xlsx.build([
        {
            name:'sheet1',
            data:titleName   
        }
    ]);
    fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});   //生成excel
}

run();


