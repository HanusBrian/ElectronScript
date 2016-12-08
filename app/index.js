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
  console.log("drop felt");

  var files = e.target.files || (e.dataTransfer && e.dataTransfer.files);
  if (files) {
    console.log(files[0].path);
    var workbook = XLS.readFile(files[0].path, { type: "binary" });
    var dataArr = XLS.utils.sheet_to_row_object_array(workbook.Sheets['Sheet1']);
    console.log(dataArr);

    fs.mkdir(__dirname + '/../temp/', () => {
      writeHelptext(dataArr, openHelptextWindow);
    });
  }
  return false;
}

function openHelptextWindow() {
  ipcRenderer.send('open-helptext-window');
}

function writeHelptext(data, callback) {

  fs.writeFile(__dirname + '/../temp/masterHelptext.htm', '', () => {
    fs.mkdir(__dirname + '/../temp/splitFiles', () => {
      var buttons =
        `
          <html>
          <head>
          <link rel="stylesheet" href="
          ` + __dirname +
        `/../app.asar/index.css"></link>
          </head>
          <body>
          <header class="header">
            <a href="./masterHelptext.htm" download>
              <div class="htButton">Download Master File</div>
            </a>
            <a href="./splitFiles.zip" download>
              <div class="htButton">Download Split Text</div>
            </a>
            <div class="htButton" onClick="handleClick(event)">Back</div>
          </header>
        `
        ;
      fs.writeFile(__dirname + '/../temp/displayHelptext.html', buttons, () => {

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
            <style>
              .table {
                border: 1px solid black;
                margin: 0px;
                padding: 0px;
                align:left; 
                border-collapse:collapse;
                border:none;
                margin-left:6.75pt;
                margin-right:6.75pt; 
                margin-bottom:10px;'
              }
              .para {
                margin-top:4.0pt;
                margin-right:0in;
                margin-bottom:4.0pt; 
                margin-left:0in;
                text-align:justify;
                line-height:normal;
              }
              .tabdat {
                vertical-align: top;
                width: 6.65in;
                border:solid windowtext 1.0pt;
                border-top:none;
                padding: 0in 5.4pt 0in 5.4pt;
                height: 110.15pt'
              }
              .span1 {
                font-family:'Arial','sans-serif';
              }
              .span2 {
                font-family:'Arial','sans-serif';
                color:black;
                font-weight:normal;
              }
              .parend {
                margin-top:4.0pt;
                margin-right:0in;
                margin-bottom:4.0pt;
                margin-left:57.2pt;
                line-height:normal;
              }
            </style>
            </head>
            <body lang=EN-US>
            <div class=WordSection1>
              <table class="table">
                <tr style='background-color:`
            + color +
            `;border-top: 1px solid black'>
                <td class="tabdat">
                  <p class="para">
                  <b><span class="span1">`
            + data[i]["Question Number"] +
            `<span style='color:green'>&nbsp; </span><span style='color:white'>`
            + data[i]["Question Text"] +
            `</span></span></b></p>
                  </td>
                </tr>
                <tr>
                  <td class="tabdat">
                    <p class="para"><b><span class="span1"><br>  Threshold: </span></b><span class="span2">`
            + data[i]["Threshold"] +
            `<br><br></span></p>
                  </td>
                </tr>
                <tr>
                  <td class="tabdat">
                    <p class="para"><b><span class="span1"><br>  Assessment:</span></b></p>
                    <p class="para"><span class="span2">`
            + data[i]["Recommended Assessment Criteria"] +
            `</span></p>
                  </td>
                </tr>
                <tr style='height:110.15pt'>
                  <td class="tabdat">
                    <p class="para"><b><span class="span1"><br> Picklist:</span></b></p>
                    <ul>`;
          if (data[i]["Picklist"]) {
            while (data[i]["Picklist"].indexOf("\n") > -1) {
              tag += `<li style='list-style-type: disc'>
                        <p class="para">
                          <span class="span2">`;
              tag += data[i]["Picklist"].substring(0, data[i]["Picklist"].indexOf("\n"));
              tag += "</span></p></li>";
              data[i]["Picklist"] = data[i]["Picklist"].substring(data[i]["Picklist"].indexOf("\n") + 1);
            }
            tag += `<li style='list-style-type: disc'>
                        <p class="para">
                          <span class="span2">`;
            tag += data[i]["Picklist"];
            tag += "</span></p></li>";
          }
          tag += `</span></p><p class="parend"></li>
            </span></p><p class="parend">
                      </li>
                    </ul>
              </table>
            </div>
          </body>
          `;
          fs.writeFileSync(__dirname + '/../temp/splitFiles/' + data[i]["Question Number"] + '.htm', tag);
          fs.appendFileSync(__dirname + '/../temp/masterHelptext.htm', tag);
          fs.appendFileSync(__dirname + '/../temp/displayHelptext.html', tag);
        }
        fs.appendFileSync(__dirname + '/../temp/displayHelptext.html',
          `
        </body>
        <script>
          var electron = require("electron");
          var ipcRenderer = electron.ipcRenderer;
          function handleClick(e) {  
            ipcRenderer.send("home-page");
          }
        </script>
        </html>
        
        `);

        zipFiles();
      });
    });
  });
  callback();
}

function zipFiles() {
  zip(__dirname + '/../temp/splitFiles',
    __dirname + '/../temp/splitFiles.zip', function (err) {
      if (err) console.log(err);
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

function handleDragend(e) {
  e.preventDefault();
  return false;
}