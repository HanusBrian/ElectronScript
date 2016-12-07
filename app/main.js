var app = require('app');
var BrowerWindow = require('browser-window');
var fs = require('fs');
var path = require('path');
var electron = require('electron');
var ipcMain = electron.ipcMain;

let win

function createWindow() {
    win = new BrowerWindow({
        width: 900,
        height: 700,
    });

    win.loadURL('file://' + __dirname + '/index.html');
    win.webContents.openDevTools();

    // Emitted when the window is closed.
    win.on('closed', () => {
        // Dereference the window object, usually you would store windows
        // in an array if your app supports multi windows, this is the time
        // when you should delete the corresponding element.
        win = null;
    });
}

app.on('ready', createWindow);

// Quit when all windows are closed.
app.on('window-all-closed', () => {
    // On macOS it is common for applications and their menu bar
    // to stay active until the user quits explicitly with Cmd + Q
    if (process.platform !== 'darwin') {
        app.quit()
    }
})

app.on('activate', () => {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (win === null) {
        createWindow()
    }
});

ipcMain.on('open-helptext-window', () => {
    var helptextWindow = new BrowerWindow({
        width: 900,
        height: 700,
    });

    helptextWindow.loadURL('file://' + __dirname + '/../temp/displayHelptext.html');

    helptextWindow.on('closed', () => {
        var dirPath = __dirname + '/../temp/splitFiles';
        try { var files = fs.readdirSync(dirPath); }
        catch (e) { return; }
        if (files.length > 0)
            for (var i = 0; i < files.length; i++) {
                var filePath = dirPath + '/' + files[i];
                if (fs.statSync(filePath).isFile())
                    fs.unlinkSync(filePath);
                else
                    rmDir(filePath);
            }
        fs.rmdirSync(dirPath);
        // Dereference the window object, usually you would store windows
        // in an array if your app supports multi windows, this is the time
        // when you should delete the corresponding element.
        helptextWindow = null;
    });
});