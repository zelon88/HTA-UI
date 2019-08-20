// / HTA-UI Desktop Application Template
// / https://github.com/zelon88/HTA-UI
// / https://github.com/zelon88

// / Author: Justin Grimes
// / Date: 8/18/2019
// / <3 Open-Source

// / Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
// / https://www.gnu.org/licenses/gpl-3.0.html

// / This HTA application template started out on the Microsoft TechNet website and has served me well.
// / I hope that someone out there can make as much use out of it as I was able to. 

// / --------------------
// / A function to format a prettified date for the UI.
function humanDate() {
  var dateVar = new Date(); 
  return dateVar.toLocaleString(); }
// / --------------------

// / --------------------
// / A function to call the VB sub saveSettings which just displays the "Save Complete" message.
function callVBSave() {
  saveSettings(); }
// / --------------------

// / --------------------
// / A function to read the contents of a text file and replace the <body> of a page with it's contents.
function readFile(path) { 
  var fso = new ActiveXObject('Scripting.FileSystemObject'),
    iStream=fso.OpenTextFile(path, 1, false);
  while(!iStream.AtEndOfStream) { 
    document.body.innerHTML += iStream.ReadLine() + '<br/>'; }        
  iStream.Close(); }
// / --------------------

// / --------------------
// / A function to read the contents of a text file and return the results.
function readFile2(path) { 
  var fso = new ActiveXObject('Scripting.FileSystemObject'),
    iStream=fso.OpenTextFile(path, 1, false);
  var data = "";
  while(!iStream.AtEndOfStream) { 
    data += iStream.ReadLine() + '<br/>'; }        
  iStream.Close(); 
  return data; }
// / --------------------

// / --------------------
// / A function to toggle the visibility of the selected element between "block" and "none".
function toggleVisibility(id) {
  var e = document.getElementById(id);
  if(e.style.display == 'block')
     e.style.display = 'none';
  else
     e.style.display = 'block'; }
// / --------------------
