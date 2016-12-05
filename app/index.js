var electron = require("electron");
var ipcRenderer = electron.ipcRenderer;
var shell = electron.shell;
var remote = electron.remote;
var ipcMain = remote.require('./main');
var fs = require('fs');
var path = require('path');
var zip = require('zip-folder');
var XLS = require('xlsjs');

function handleDrop(e) {
  e.preventDefault();
  var files = e.target.files || (e.dataTransfer && e.dataTransfer.files);
  if (files) {
    console.log(files[0].path);
    parseFile(files[0].path);
  }
  return false;
}

function parseFile(file) {
  var workbook = XLS.readFile(file, { type: "binary" });

  var dataArr = XLS.utils.sheet_to_row_object_array(workbook.Sheets['Sheet1']);

  fs.writeFile('./helptextData.json',
    JSON.stringify(dataArr, null, 4),
    'utf-8',
    (err) => {
      if (err) console.log(err.message);
      else {
        console.log('File successfully written');
        displayHelptext();
      }
    }
  );
}

function displayHelptext() {
  var data = require('./helptextData.json');
  writeHelptext(data, openHelptextWindow);
}

function openHelptextWindow() {
  ipcRenderer.send('open-helptext-window');
}

function writeHelptext(data) {
  if (fs.existsSync(__dirname + '/app/output/'))
    this.rmDir(__dirname + '/app/output/');

  fs.writeFile('./app/output/masterHelptext.htm', '', ()=>{});

  fs.mkdir('./app/output', () => {
    fs.mkdir('/app/output/splitFiles', () => {
      var buttons = `
      <html><body>
        <header>
        <a href="./masterHelptext.htm" download>
          <div>Download Master File</div>
        </a>
        <a href="./splitFiles.zip" download>
          <div>Download Split Text</div>
        </a>
        </header>
        `;
      fs.writeFile('./app/output/displayHelptext.html', buttons, () => { });
      var tag = "";
      var color = "";
      for (var i = 0; i < data.length; i++) {
        switch (data[i]["Criticality"]) {
          case 'Critical': color = "red"; break;
          case 'Major': color = "orange"; break;
          case 'NonCritical': color = "orange"; break;
          case 'Minor': color = "blue"; break;
          case 'Variable': color = "purple"; break;
        }
        tag = `
        <head>
          <meta http-equiv=Content-Type content='text/html; charset=windows-1252'>
          <meta name=Generator content='Microsoft Word 14 (filtered)'>
        </head>
        <body lang=EN-US>
          <div class=WordSection1>
            <table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0 align=left style='border-collapse:collapse;border:none;margin-left:6.75pt;margin-right:6.75pt; margin-bottom:10px;'>
              <tr style='background-color:`
          + color +
          `;'>
                <td width=638 valign=top style='width:6.65in;border:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;'>
                  <p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt; margin-left:0in;text-align:justify'><b><span style='font-family:'Arial','sans-serif''>`
          + data[i]["Question Number"] +
          `<span style='color:green'>&nbsp; </span><span style='color:white'>`
          + data[i]["Question Text"] +
          `</span></span></b></p>
                </td>
              </tr>
              <tr>
                <td width=638 valign=top style='width:6.65in;border:solid windowtext 1.0pt; border-top:none;padding:0in 5.4pt 0in 5.4pt'>
                  <p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt;margin-left:0in;line-height:normal'><b><span style='font-family:'Arial','sans-serif''><br>  Threshold: </span></b><span style='font-family:'Arial','sans-serif';font-weight:normal' [innerHtml]='data.threshold'>`
          + data[i]["Threshold"] +
          `<br><br></span></p>
                </td>
              </tr>
              <tr>
                <td width=638 valign=top style='width:6.65in;border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt'>
                  <p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt;margin-left:0in;line-height:normal'><b><span style='font-family:'Arial','sans-serif''><br>  Assessment:</span></b></p>
                  <p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt;line-height:normal'><span style='font-family:'Arial','sans-serif';color:black;font-weight:normal' [innerHtml]='data.assessment'>`
          + data[i]["Recommended Assessment Criteria"] +
          `</span></p>
                </td>
              </tr>
              <tr style='height:110.15pt'>
                <td width=638 valign=top style='width:6.65in;border:solid windowtext 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:110.15pt'>
                  <p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt;margin-left:0in;line-height:normal'><b><span style='font-family:'Arial','sans-serif''><br> Picklist:</span></b></p>
                  <ul>
                    <li style='list-style-type: disc'>
                      <p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt;line-height:normal'>
                        <span style='font-family:'Arial','sans-serif';color:black;font-weight:normal'>`
          + data[i]["Picklist"] +
          `</span></p><p class=MsoNormal style='margin-top:4.0pt;margin-right:0in;margin-bottom:4.0pt;margin-left:57.2pt;line-height:normal'>
                    </li>
                  </ul>
            </table>
          </div>
        </body>
      `
        fs.writeFile('./app/output/splitFiles/' + data[i]["Question Number"] + '.htm', tag, () => { });
        fs.appendFile('./app/output/masterHelptext.htm', tag, () => { });
        fs.appendFile('./app/output/displayHelptext.html', tag, () => { });
      }
      fs.appendFile('./app/output/splitFiles', '</body></html>', () => { });

      zip('./app/output/splitFiles/', './app/output/splitFiles.zip', function (err) {
        if (err) console.log(err);
      });
    })
  });
  openHelptextWindow();
}

function handleDragover(e) {
  e.preventDefault();
  e.stopPropagation();
  return false;
}

function handleDragover(e) {
  e.preventDefault();
  return false;
}

var rmDir = function (dirPath) {
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
};