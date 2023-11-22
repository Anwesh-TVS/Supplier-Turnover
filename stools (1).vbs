'=======================================================================
'SAP Logon screen
'=======================================================================
'The below section will create an SAP session.

set WshShell = CreateObject("WScript.Shell")

Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")

WScript.Sleep 3000

   Do While proc.Status = 0

      WScript.Sleep 100

      if proc.Status = 0 then
        Exit Do
      end if

   Loop
'==========================================================
Set Arg = WScript.Arguments

Set SapGui = GetObject("SAPGUI")

Set Appl = SapGui.GetScriptingEngine

'Set Connection = Appl.Openconnection("PRD", True)
Set Connection = Appl.Openconnection(Cstr(Arg(3)), True)

on error resume next

    Set session = Connection.Children(0)

    If Err.Number <> 0 Then
        'msgbox "SAP Scripting was disabled"
        Err.Clear
        Wscript.Quit
    end if

session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Arg(0)
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Arg(1)

'session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "IT-RPA"
'session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "It-RP@123#"

session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "E"

session.findById("wnd[0]").sendVKey 0

'file_path = Cstr(Arg(2))

variant_name = ucase(Cstr(Arg(4)))
tcode = ucase(Cstr(Arg(5)))
'variant_name = "RPA-STO"
'tcode = "ZMMR030"

set Arg = Nothing
'==================================================================================================

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

'===========================================================================
' Authon       :  K.Prabakaran
' Created on   :  02-05-2022
' Purpose      :  CP - Supplier Tools, service,supplier turnover report
' Modified on  :  -
' M.Purpose    :  -
' Version      :  1.5.4.0
'===========================================================================
TerminateExcel()
Set fos = CreateObject("Scripting.FileSystemObject")
GetCurrentFolder= fos.GetAbsolutePathName(".")
folder_paths = cstr(GetCurrentFolder)

' folder_paths = cstr(GetCurrentFolder) & "\CP"

' if Not filesys.FolderExists(folder_paths) Then
'    Set newfolder = filesys.CreateFolder(folder_paths)
' end if

'variant_name = "RPA-TOOL" 'RPA-SERVICE

' if variant_name = "RPA-TOOL" then
'    file_name = "TOOL.XLSX"
' elseif variant_name = "RPA-SERVICE" then
'    file_name = "SERVICE.XLSX"
' elseif variant_name = "RPA-STO" then
'    file_name = "STO.XLSX"
' end if

file_name = variant_name & ".xlsx"
file_name_output = variant_name & "_OUTPUT.xlsx"

if fos.FileExists(folder_paths & "\" & file_name) then
   fos.DeleteFile(folder_paths & "\" & file_name)
end if

if fos.FileExists(folder_paths & "\" & file_name_output) then
   fos.DeleteFile(folder_paths & "\" & file_name_output)
end if
'===============================================================================
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = tcode
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = variant_name
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV"
' session.findById("wnd[0]/usr/ctxtLISTU").setFocus
' session.findById("wnd[0]/usr/ctxtLISTU").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[8]").press

'=============================================================================================
' Hide column in the report
if tcode = "ME2L" then
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"ICON_PO_HIST"
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ICON_PO_HIST"
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&COL_INV"
   session.findById("wnd[0]/tbar[1]/btn[43]").press
end if

if variant_name = "RPA-STO" and tcode = "ZMMR030" then
   'session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select

   ' FROM HERE code changes has been performed######################################################


   'If Not IsObject(application) Then
   'Set SapGuiAuto  = GetObject("SAPGUI")
   'Set application = SapGuiAuto.GetScriptingEngine
   'End If
   'If Not IsObject(connection) Then
      'Set connection = application.Children(0)
   'End If
   'If Not IsObject(session) Then
      'Set session    = connection.Children(0)
   'End If
   'If IsObject(WScript) Then
      'WScript.ConnectObject session,     "on"
      'WScript.ConnectObject application, "on"
   'End If
   'session.findById("wnd[0]").maximize
   'session.findById("wnd[0]/tbar[0]/okcd").text = "ZMMR030"
   'session.findById("wnd[0]").sendVKey 0
   'session.findById("wnd[0]/tbar[1]/btn[17]").press
   'session.findById("wnd[1]/usr/txtV-LOW").text = "RPA-STO"
   'session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
   'session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 7
   'session.findById("wnd[1]/tbar[0]/btn[8]").press
   ''''''''''''''''''''''''''''''''''''''
   ' Get the current month
   currentMonth = Month(Now)

   ' Calculate the next month
   nextMonth = currentMonth + 1

   ' If the current month is between 01 and 03 (January to March), use the current year; otherwise, use the next year
   If currentMonth >= 1 And currentMonth <= 3 Then
      currentYear = Year(Now)

   ' Calculate the next year
      nextYear = currentYear
      currentYear = currentYear - 1
      ' Format the dates using the current year
      formattedDateLow = "04." & Right("0000" & currentYear, 4)
      formattedDateHigh = "03." & Right("0000" & nextYear, 4)
   Else
      ' Format the dates using the next year
      currentYear = Year(Now)

   ' Calculate the next year
      nextYear = currentYear +1

      formattedDateLow = "04." & Right("0000" & currentYear, 4)
      formattedDateHigh = "03." & Right("0000" & nextYear , 4)
   End If

   ' Set the text fields with the formatted dates

   session.findById("wnd[0]/usr/txtSO_YRMN-LOW").text = formattedDateLow
   session.findById("wnd[0]/usr/txtSO_YRMN-HIGH").text = formattedDateHigh


   session.findById("wnd[0]/usr/txtSO_YRMN-HIGH").setFocus
   session.findById("wnd[0]/usr/txtSO_YRMN-HIGH").caretPosition = 7
   session.findById("wnd[0]/tbar[1]/btn[8]").press
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 3,"EBELP"
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "3"
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
   session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
   ''''''''''''''''''''''
   'session.findById("wnd[1]/tbar[0]/btn[0]").press
   'session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_paths
   'session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
   'session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 36
   'session.findById("wnd[1]/tbar[0]/btn[0]").press
   'session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SRPA-STO.xlsx"
   'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8s
   'session.findById("wnd[1]/tbar[0]/btn[0]").press
end if
'=============================================================================================
'session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\OneDrive - TVS Motor Company Ltd\wwwroot\RPA\CP"

session.findById("wnd[1]/tbar[0]/btn[0]").press 

session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_paths
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
'MsgBox folder_paths
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
session.findById("wnd[1]/tbar[0]/btn[0]").press
'=============================================================================================
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

TerminateExcel()

'============================================================================================================
' Log-off SAP - START
'============================================================================================================
session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
session.findById("wnd[0]").sendVKey 0
'============================================================================================================
' Log-off SAP - END
'============================================================================================================

Sub TerminateExcel

    'msgbox "Excel kill method called"

    Dim Process 

    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'EXCEL.EXE'")

    Process.Terminate

    Next

End Sub
