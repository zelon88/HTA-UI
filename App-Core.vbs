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
'  Bootstrap("PHP\php.exe", scriptsDirectory & "PHP\test.php")
'The above function call uses the Bootstrap() function to call 
'Binaries\PHP\php.exe with an argument that evaluates to Scripts\PHP\test.php.
'The result will be that the PHP binary is used to execute a PHP script.
'If Async is set to TRUE, HTA-UI will wait for the command to finish before continuing.
Function Bootstrap(BinaryToRun, Command, Async)
  Dim objShell, objShellExec, run, tempFile, tempData
  tempFile = tempDirectory & "temp.txt"
  If Async = TRUE Then 
    async = TRUE
  Else 
    async = ""
  End If
  Set objShell = CreateObject("WScript.Shell")
  run = "C:\Windows\System32\cmd.exe /c " & binariesDirectory & BinaryToRun & " " & Command & " > " & tempFile
  objShell.Run run, 0, async
  Set tempData = objFSO.OpenTextFile(tempFile, 1)
  Bootstrap = tempData.ReadAll()
  tempData.Close
  objFSO.DeleteFile(tempFile)
End Function
