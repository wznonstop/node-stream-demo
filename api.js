/**
 * api.js
 *@author wuzhe
 *create 2017-12-26 18:47
 */
var express = require('express');
var app = express();
var Excel = require('exceljs');


app.get('/excel', function (req, res, next) {
  res.set('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

  res.setHeader("Content-Disposition","attachment;filename=2017.12.26.xlsx")
  res.set('Set-Cookie', 'fileDownload=true; path=/')
  var data = [
    {
      "province": "山西",
      "city": "大同，太原"
    },
    {
      "province": "黑龙江",
      "city": "佳木斯，哈尔滨"
    },
    {
      "province": "安徽",
      "city": "合肥，安庆"
    }
  ];
  var options = {
    stream: res,
    useStyles: true,
    useSharedStrings: true
  }

  var start_time = Date.now();
  var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
  var worksheet = workbook.addWorksheet('Sheet');
  worksheet.columns = [
    { header: '去过的省', key: 'province' },
    { header: '具体的市', key: 'city' }
  ];

  for (var i = 0; i < data.length; i ++) {
    worksheet.addRow(data[i]).commit();
  }
  // worksheet.commit()
  workbook.commit()
    .then(function() {
      // the stream has been written
      var end_time = new Date();
      var duration = end_time - start_time;

      console.log("创建excel程序执行完毕，用时：", duration);
    });
})

app.listen(3000, function () {
  console.log('listening port 3000~~~~')
})