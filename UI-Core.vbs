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
'Large portions of code in this file were borrowed from the Microsoft TechNet website on 8/14/2019 
'in accordance with the Microsoft Limited Public License...
'https://gallery.technet.microsoft.com/scriptcenter/796bd584-0fdb-43bc-a5d2-aa5fc99a5e5d

'--------------------------------------------------
'Define global variables for the session.
Dim objFSO, strComputer, objWMIService, scriptsDirectory, binariesDirectory, _
 colItems, objItem, intHorizontal, intVertical, nLeft, nTop, sItem, helpLocSetting, _
 version, currentDirectory, appName, developerName, developerURL, windowHeight, windowWidth

version = "v1.1"

helpLocSetting = "https://github.com/zelon88"
appName = "HTA-UI"
developerName = "Justin Grimes"
developerURL = "https://github.com/zelon88"

windowHeight = 660
windowWidth = 600

Const sMenuItems = "File,Settings,Help" 
Const sFile = "Exit" 
Const sSettings = "View Settings"
Const sHelp = "Help, About" 
Const sHTML = "&nbsp;&nbsp;&nbsp;#sItem#&nbsp;&nbsp;&nbsp;" 
Dim dMenus, sMenuOpen 
Set objFSO = CreateObject("Scripting.FileSystemObject")
currentDirectory = objFSO.GetAbsolutePathName(".")
scriptsDirectory = currentDirectory & "\Scripts\"
binariesDirectory = currentDirectory & "\Binaries\"
strComputer = "."
'--------------------------------------------------

'--------------------------------------------------
'Load the main application window.
Sub Window_OnLoad 
  Dim entry 
  Set dMenus = createObject("Scripting.Dictionary") 
  For Each entry In Split(sMenuItems, ",") 
    menu.innerHTML = menu.innerHTML & "&nbsp;<span id=" & entry _ 
      & " style='padding-bottom:2px' onselectstart=cancelEvent>&nbsp;" _ 
      & entry & "&nbsp;</span>&nbsp;&nbsp;" 
    dMenus.Add entry, Split(eval("s" & entry), ",") 
  Next 
  sMenuOpen = "" 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Resize the application window.
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
For Each objItem in colItems
  'Max screen width in pixels.
  intHorizontal = objItem.ScreenWidth
  'Max screen height in pixels.
  intVertical = objItem.ScreenHeight
Next
  window.resizeTo windowHeight,windowWidth
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes on mouse hover.
Sub menu_onmouseover 
  clearmenu 
  with window.event.srcElement 
    if .parentElement.ID = "menu" then 
      .style.border = "thin outset" 
      .style.cursor = "arrow" 
    end if 
  end with 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when mouse leaves hover.
Sub menu_onmouseout 
  with window.event.srcElement 
    .style.border = "none" 
    .style.cursor = "default" 
  end with 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when mouse hovers over a dropdown menu item.
Sub dropmenu_onmouseover 
  with window.event 
    .srcElement.style.cursor = "arrow" 
    .cancelbubble = true 
    .returnvalue = false 
  end with 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when a user hovers over a dropdown menu selection.
sub SubMenuOver 
  with window.event.srcElement 
    if .ID = "dropmenu" then exit sub 
    .style.backgroundcolor = "darkblue" 
    .style.color = "white" 
    .style.cursor = "arrow" 
  end with 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when mouse leaves hover over a dropdown menu selection.
sub SubMenuOut 
  with window.event.srcElement 
    .style.backgroundcolor = "lightgrey" 
    .style.color = "black" 
    .style.cursor = "default" 
  end with 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when a user clicks on a menu item.
Sub menu_onclick 
  Dim oEL, oItem 
  if sMenuOpen <> "" then exit sub 
  with window.event.srcElement 
    if .ID <> "menu" then 
      .style.border = "thin inset" 
      nLeft = .offsetLeft 
      ntop  = .offsetTop + replace(menu.style.Height, "px", "") - 5 
      sMenuOpen = trim(.innertext) 
      with dropmenu 
        with .style 
          .border = "thin outset" 
          .backgroundcolor = "lightgrey" 
          .position = "absolute" 
          .left = nLeft 
          .top = nTop 
          .width = "100px" 
          .zIndex = "101"
        end with 
        for each sItem in dMenus.Item(sMenuOpen) 
          set oEL = document.createElement("SPAN") 
          .appendChild(oEL) 
          with oEl 
            .ID = sItem 
            .style.height = "20px" 
            .style.width = dropmenu.style.width 
            .style.zIndex = "102"
            .innerHTML = Replace(sHTML, "#sItem#", trim(sItem)) 
            set .onmouseover = getRef("SubMenuOver") 
            set .onmouseout = getRef("SubMenuOut") 
            set .onclick = getRef("SubMenuClick") 
            set .onselectstart = getRef("cancelEvent") 
          end with
          set oEL = document.createElement("BR") 
          .appendChild(oEL) 
        next 
      end with
    end if 
  end with
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle when an event is cancelled.
sub cancelEvent 
  window.event.returnValue = false 
end sub
'--------------------------------------------------

'--------------------------------------------------
'Handle when a user deselects a menu.
sub clearmenu 
  dropmenu.innerHTML = "" 
  dropmenu.style.border = "none" 
  dropmenu.style.backgroundcolor = "transparent" 
  if sMenuOpen <> "" then 
    document.getElementByID(sMenuOpen).style.border = "none" 
    sMenuOpen = "" 
  end if 
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle when a user clicks on a submenu.
Sub SubMenuClick 
  sItem = trim(window.event.srcElement.innerText) 
  clearmenu 
  Select Case lcase(sItem) 
    case "exit" 
      window.close  
    case "view settings"
      document.location = "Settings.hta"
    case "about" 
      msgbox version & ". " & vbCRLF & vbCRLF & "Developed by " & developerName & "."_ 
        & vbCRLF & vbCRLF & developerURL, _ 
        vbOKOnly + vbInformation, "About "& appName 
    case else 
      msgbox "You can get support for '" & appName & "' by visiting: " _ 
      & vbCRLF & vbCRLF & helpLocSetting, vbOKOnly + vbInformation, appName & " Help"
  end Select 
end sub 
'--------------------------------------------------
