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

// / --------------------
// / A function to save the current settings to the settings cache.
// / Change these to actual settings, with the name of the file set to the name of the setting.
// / Example:  LogDir setting should be stored in LogDir.dat and the setting name is LogDir.
function updateSetting(setting) {
  var input = '';
  if (setting == 'setting1') { 
  	var file = 'Cache/setting1.dat';
  	var input = document.getElementById('setting1').value; }
  if (setting == 'setting2') { 
  	var file = 'Cache/setting2.dat';
    var input = document.getElementById('setting2').value; }
  if (setting == 'setting3') { 
    var file = 'Cache/setting3.dat';
    var input = document.getElementById('setting3').value; }
  if (setting == 'setting4') { 
    var file = 'Cache/setting4.dat';
    var input = document.getElementById('setting4').value; }
  if (setting == 'setting5') { 
    var file = 'Cache/setting5.dat';
    var input = document.getElementById('setting5').value; }
  var data = input;
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var s = fso.OpenTextFile(file, 2, true);
  s.WriteLine(data);
  s.Close(); }
// / --------------------

// / --------------------
// / A function to load the log location from the settings cache.
// / Load the setting by name coinciding with setting file.
// / Example:  LogDir setting can be retrieved from LogDir.dat using name LogDir.
function getSetting(setting) {
  var fso = new ActiveXObject('Scripting.FileSystemObject'),
  iStream = fso.OpenTextFile('Cache/' + setting + '.dat', 1, false);
  while(!iStream.AtEndOfStream) { 
    data = iStream.ReadLine(); }        
  iStream.Close(); 
  return data; }
// / --------------------
