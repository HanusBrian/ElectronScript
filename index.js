var electron = require("electron");
var ipcRenderer = electron.ipcRenderer;
var shell = electron.shell;
var remote = electron.remote;
var ipcMain = remote.require('./main');
var fs = require('fs');
var path = require('path');

function handleDrop(e) {
  e.preventDefault();
  var files = e.target.files || (e.dataTransfer && e.dataTransfer.files);
  if (files) {
    console.log(files[0].path);
    ipcRenderer.send('file-drop', files[0].path);
  }
  return false;
}

ipcRenderer.on('file-done', (event) => {
  var data = require('./helptextData.json');
  writeHelptext(data);
  ipcRenderer.send('open-helptext-window');
})

function writeHelptext(data) {
  this.rmDir(__dirname + '/output/');
  fs.writeFile('./output/masterHelptext.htm','', ()=>{});
  var buttons = `
  <body>
  <input id="download-button" type="button" value="Download Master HelpText" onClick="window.location.href='masterTemplate.htm'">
  <input id="download-button" type="button" value="Download Split Question Folder" onClick="window.location.href='splitFiles'">
  </body>
  `;
  fs.writeFile('./output/displayHelptext.htm', buttons, () => {

  fs.mkdir('output', () => {
    fs.mkdir('output/splitFiles', () => {
      for (var i = 0; i < data.length; i++) {
        var tag = "";
        var color = "";
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
        fs.writeFileSync('./output/splitFiles/' + data[i]["Question Number"] + '.htm', tag);
        fs.appendFileSync('./output/masterHelptext.htm', tag);
        fs.appendFileSync('./output/displayHelptext.htm', tag);
      }
      console.log(tag);
    })
  });
  });
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

rmDir = function (dirPath) {
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