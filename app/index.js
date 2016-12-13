var electron = require("electron");
var ipcRenderer = electron.ipcRenderer;
var shell = electron.shell;
var remote = electron.remote;
var ipcMain = remote.require('./main');
var fs = require('fs');
var path = require('path');
var zip = require('zip-folder');
var XLS = require('xlsjs');
var dev = true;

function handleDrop(e) {
  e.preventDefault();
  console.log("drop felt");

  var files = e.target.files || (e.dataTransfer && e.dataTransfer.files);
  if (files) {
    console.log(files[0].path);
    var workbook = XLS.readFile(files[0].path, { type: "binary" });
    var dataArr = XLS.utils.sheet_to_row_object_array(workbook.Sheets['Sheet1']);
    console.log(dataArr);
    formatDataArr(dataArr);

    fs.mkdir(__dirname + '/../temp/', () => {
      writeHelptext(dataArr, openHelptextWindow);
    });
  }
  return false;
}

function formatDataArr(dataArr) {
  for (var i = 0; i < dataArr.length; i++) {
    var rac = dataArr[i]["Recommended Assessment Criteria"];
    if (rac) {
      if (rac.indexOf("\u2022") > -1) {
        rac = parseOl(rac);
      }
      if (rac.indexOf('\n')) {
        rac = parseNl(rac);
      }
    }
    if (dataArr[i]["Picklist"]) {
      parsePicklist(dataArr[i]["Picklist"]);
    }
  }
}

function parseOl(data) {
  var temp = "";
  var isFirstLi = true;
  while (data.indexOf("\u2022") > -1) {
    temp += data.substring(0, data.indexOf("\u2022"));

    if (isFirstLi) {
      temp += "<ul><li>";
      isFirstLi = false;
    }
    else
      temp += "<li>";

    if (data.indexOf("\n") > -1) {
      temp += data.substring(data.indexOf("\u2022") + 1, data.indexOf('/n') + 1);
      temp += "</li>";
      data = data.indexOf('/n'+ 1, data.indexOf("\u2022") + 1);
    }
    else {
      temp += data.substring(data.indexOf("\u2022") + 1);
      temp += "</li>";
    }
  }
  return temp;
}

function parseNl(data) {
  var temp = "";

  while (data.indexOf("\n") > -1) {
    temp += data.substring(0, data.indexOf("\n"));
    temp += "<br><br>";
    data = data.substring(data.indexOf("\n") + 1);
  }
  temp += data;
  return temp;
}

function parsePicklist(data) {
  var temp = "";
  while (data.indexOf("\n") > -1) {
    temp += `<li style='list-style-type: disc'>
                  <p class="para">
                    <span class="span2">`;
    temp += data.substring(0, data.indexOf("\n"));
    temp += "</span></p></li>";
    data = data.substring(data.indexOf("\n") + 1);
  }
  temp += `<li style='list-style-type: disc'>
                        <p class="para">
                          <span class="span2">`;
  temp += data;
  temp += "</span></p></li>";
  return temp;
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
          ` + __dirname
      if (dev)
        buttons += './index.css">';
      else
        buttons += '/../app.asar/index.css">';
      buttons += `
        </link>
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
            case 'Informational': color = '#3D8F22'; break;
            case 'Critical': color = "#E00034"; break;
            case 'Major': color = "#EEAF00"; break;
            case 'NonCritical': color = "#EEAF00"; break;
            case 'Minor': color = "#007AC9"; break;
            case 'Variable': color = "purple"; break;
          }
          tag = `
            <head>
            <meta http-equiv=Content-Type content='text/html; charset="utf-8"'>
            <meta name=Generator content='utf-8'>
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
          tag += data[i]["Picklist"];
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