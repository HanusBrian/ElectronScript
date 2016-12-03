var app = require('app');
var BrowerWindow = require('browser-window');
var fs = require('fs');
var path = require('path');
var electron = require('electron');
var ipcMain = electron.ipcMain;
var XLS = require('xlsjs');

app.on('ready', () => {
    var win = new BrowerWindow({
        width: 900,
        height: 700,
    });

    win.loadURL('file://' + __dirname + '/index.html');

});

ipcMain.on('file-drop', (event, path) => {
    if (path) {
        console.log("File info: " + path);
        parseFile(path, event);
    } else {
        console.log("no files found");
    }
});

function parseFile(file, event) {
    var workbook = XLS.readFile(file, { type: "binary" });

    var dataArr = XLS.utils.sheet_to_row_object_array(workbook.Sheets['Sheet1']);

    for (var i = 0; i < dataArr.length; i++) {
        console.log(dataArr[i]);
    }
    fs.writeFile('helptextData.json',
        JSON.stringify(dataArr, null, 4),
        'utf-8',
        (err) => {
            if (err) console.log(err.message);
            else {
                console.log('File successfully written');
                event.sender.send('file-done');
            }
        }
    );
}
ipcMain.on('open-helptext-window', () =>{
    var helptextWindow = new BrowerWindow({
        width: 900,
        height: 700,
    });

    helptextWindow.loadURL('file://' + __dirname + '/output/displayHelptext.html');
});