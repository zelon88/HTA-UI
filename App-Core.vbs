'HTA-UI Desktop Application Template
'https://github.com/zelon88/HTA-UI
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/14/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'Portions of the UI-Core.vbs file are licensed under the Microsoft Limited Public License.
'Copies of all applicable software licenses can be found in the "Documentation" directory.

'This HTA application template started out on the Microsoft TechNet website and has served me well.
'I hope that someone out there can make as much use out of it as I was able to. 

Option Explicit

'--------------------------------------------------
'Define global variables for the session.
Dim objShell, BinaryToRun, Command, run

Set objShell = CreateObject("WScript.Shell")
'--------------------------------------------------

'--------------------------------------------------
'Bootstrap some other program or code in the Binaries folder.
'Example for bootstrapping a PHP script.
'  Bootstrap(binariesDirectory & "PHP\php.exe", scriptsDirectory & "PHP\test.php")
'The above function call uses the Bootstrap() function to call 
'Binaries\PHP\php.exe with an argument that evaluates to Scripts\PHP\test.php.
'The result will be that the PHP binary is used to execute a PHP script.
Function Bootstrap(BinaryToRun, Command)
  run = binariesDirectory & BinaryToRun & Command 
  objShell.exec(run)
  Bootstrap = objShellExec.StdOut.ReadAll
End Function