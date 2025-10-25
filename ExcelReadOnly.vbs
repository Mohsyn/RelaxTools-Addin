-------------------------------------------------------------------------------
Script to enable "Open as read-only" on right-clicking Excel files

ExcelReadOnly.vbs

Copyright (c) 2015 Y. Watanabe

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php
-------------------------------------------------------------------------------
Tested on: Windows 7 + Excel 2010 / Windows 8 + Excel 2013
-------------------------------------------------------------------------------
Reference sites below

Untitled - Display "Open as read-only" in the right-click menu (Excel & Word)
https://sites.google.com/site/universeof/tips/openasreadonly
-------------------------------------------------------------------------------

Option Explicit

On Error Resume Next

If WScript.Arguments.Count = 0 Then

    'Run yourself as administrator
    With CreateObject("Shell.Application")
        .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ dummy", "", "runas"
    End With
    
    WScript.Quit
    
End If


With WScript.CreateObject("WScript.Shell")

    'Removed 'Extended' key so menu appears without pressing shift
    .RegDelete "HKCR\Excel.Sheet.8\shell\OpenAsReadOnly\Extended"
    .RegDelete "HKCR\Excel.Sheet.12\shell\OpenAsReadOnly\Extended"
    .RegDelete "HKCR\Excel.SheetMacroEnabled.12\shell\OpenAsReadOnly\Extended"

    Err.Clear

   'Enable read-only
    .RegWrite "HKCR\Excel.Sheet.8\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.Sheet.12\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.SheetMacroEnabled.12\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"

End With

If Err.Number = 0 Then
    MsgBox "You have successfully changed the association.", vbInformation + vbOkOnly, "read-only enable"
Else
    MsgBox "An error has occurred.", vbCritical + vbOkOnly, "Read-only enable"
End IF


 
