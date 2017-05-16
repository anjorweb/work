
环境：Nodejs

工具：exceljs

功能：读取一个json文件，分析数据结构，生成excel并写入数据


```javascript
var Excel = require('exceljs');
var fs = require("fs");

var start_time = new Date();
var fileNameDate = start_time.getFullYear()+"-"+(1+start_time.getMonth())+'-'+start_time.getDate();
var workbook = new Excel.stream.xlsx.WorkbookWriter({
    filename: 'export/'+ fileNameDate +'.xlsx'
});
var worksheet = workbook.addWorksheet('Sheet');

worksheet.columns = [
    { header: '姓名', key: 'name' },
    { header: '公司', key: 'company' },
    { header: '电话', key: 'tel' },
    { header: '邮箱', key: 'email' },
    { header: '创建时间', key: 'createAt' },
    { header: '来源', key: 'media' }

];

fs.readFile("jsonfile/userInfo.json", "utf-8", function(error, config) {
    if (error) {
        console.log(error);
        console.log("config文件读入出错");
    }
    var jsonData = JSON.parse(config.toString());
    forJSON(jsonData);
});
var userData = [];
function forJSON(data){
    var jsonData = data;
    if (jsonData.results.length > 0){
        jsonData.results.forEach(function (res) {
            userData.push({
                name:res.name,
                company:res.company,
                tel:res.tel,
                email:res.email,
                createdAt:res.createdAt,
                media:res.media?res.media:"其他"
            });
        });

        var length = userData.length;
// 当前进度
        var current_num = 0;
        var time_monit = 400;
        var temp_time = Date.now();

        console.log('开始添加数据');
// 开始添加数据
        for(let i in userData) {
            worksheet.addRow(userData[i]).commit();
            current_num = i;
            if(Date.now() - temp_time > time_monit) {
                temp_time = Date.now();
                console.log((current_num / length * 100).toFixed(2) + '%');
            }
        }
        console.log('添加数据完毕：', (Date.now() - start_time));
        workbook.commit();

        var end_time = new Date();
        var duration = end_time - start_time;

        console.log('用时：' + duration);
        console.log("程序执行完毕");
    }
}


```
