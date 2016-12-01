  var electron = require("electron");
  var ipcRenderer = electron.ipcRenderer;
  var shell = electron.shell;
  var remote = electron.remote;
  var ipcMain = remote.require('./main');

  window.addEventListener("dragover",function(e){
  	e = e || event;
  	e.preventDefault();
	},false);
  window.addEventListener("drop",function(e){
    e = e || event;
    e.preventDefault();
  },false);

  function handleDrop(e) {
	  var files = e.target.files || (e.dataTransfer && e.dataTransfer.files);
	  if(files) {
		  console.log(files[0].path);
		ipcRenderer.send('file-drop', files[0].path);
	  }
    	ipcRenderer.on('file-done', (event) => {
		console.log("I see that file is done!!!");
	})
	  return false;
  }
  // function handleDragover(e){
  //   e.preventDefault();
	//   return false;
  // }
  // function handleDragend(e){ 
  //   e.preventDefault();
	//   return false;
  // }

