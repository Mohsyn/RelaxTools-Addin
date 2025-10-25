' -------------------------------------------------------------------------------
' RelaxTools-Addin installation script Ver.1.0.6
' -------------------------------------------------------------------------------
'Reference site
'A certain SE's tweet
' How to automatically install/uninstall add-ins in Excel using VBScript
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
' Modification
' 1.0.6 Modified installation path to use Application.UserLibraryPath.
' 1.0.5 Fixed to install VBS that opens the book of the same name for reference.
' 1.0.4 Delete VBS for multi-process as it is no longer needed.
' 1.0.3 Fixed to copy VBS for multi-process.
' 1.0.3 Fixed to copy images folder.
' 1.0.2 Corrected the case where the add-in file obtained from the Internet with Windows Update is not loaded in Excel.
' A warning and properties window will be displayed to ask you to "unblock".
' -------------------------------------------------------------------------------
Option Explicit
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin
Dim imageFolder
Dim appFile
Dim objWshShell
Dim objFileSys
Dim strPath
Dim objFolder
Dim objFile

'アドイン情報を設定 
addInName = "RelaxTools Addin" 
addInFileName = "Relaxtools.xlam"
appFile = "rlxAliasOpen.vbs"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

IF Not objFileSys.FileExists(addInFileName) THEN
    MsgBox "Extract the zip file and run it.", vbExclamation, addInName 
    WScript.Quit 
END IF

IF MsgBox(addInName & "Would you like to install it?" & vbCrLf &  "Version 4.0.0 Please note that the settings will not be inherited after or before that.", vbYesNo + vbQuestion, addInName) = vbNo Then 
    WScript.Quit 
End IF

'Excel instantiation 
With CreateObject("Excel.Application") 

    'Creating the installation path'Excel instantiation 
    strPath = .UserLibraryPath
    imageFolder = objWshShell.SpecialFolders("Appdata") & "\RelaxTools-Addin\"

    'Create the installation folder if it does not exist
    IF Not objFileSys.FolderExists(strPath) THEN
        objFileSys.CreateFolder(strPath)
    END IF

    installPath = strPath & addInFileName

    'File copy (overwrite) 
    objFileSys.CopyFile  addInFileName ,installPath , True

    'Create an image folder if it doesn't exist
    IF Not objFileSys.FolderExists(imageFolder) THEN
        objFileSys.CreateFolder(imageFolder)
    END IF

    'Copy (overwrite) image folder 
    objFileSys.CopyFolder  "Source\customUI\images" ,imageFolder , True

    'Copy (overwrite) file 
    objFileSys.CopyFile  appFile, imageFolder & appFile, True

    'Add-in registration 
    .Workbooks.Add
    Set objAddin = .AddIns.Add(installPath, True) 
    objAddin.Installed = True

    'Excel end 
    .Quit

End WIth

IF Err.Number = 0 THEN 
    MsgBox "Add-in installation has completed.", vbInformation, addInName 

    'プロパティファイル表示
    CreateObject("Shell.Application").NameSpace(strPath).ParseName(addInFileName).InvokeVerb("properties")
    MsgBox "Files obtained from the Internet may be blocked by Excel." & vbCrlf & "Open the properties window and click Unblock." & vbCrLf & vbCrLf & "If Unblock is not displayed in the properties, no action is required.", vbExclamation, addInName 

ELSE 
    MsgBox "An error has occurred." & vbCrLF & "If Excel is running, please close it.", vbExclamation, addInName 
    WScript.Quit 
End IF

If MsgBox("Do you want to enable Explorer right-click (open book with the same name for reference)?" & vbCrLf & "Administrator privileges are required to run.", vbYesNo + vbQuestion, addInName) <> vbNo Then 
    objWshShell.Run "rlxAliasOpen.vbs /install", 1, true
End IF

If MsgBox("Enable Explorer right-click (read-only in Excel)?" & vbCrLf & "Administrator privileges are required to run it.", vbYesNo + vbQuestion, addInName) = vbNo Then 
    WScript.Quit 
End IF

objWshShell.Run "ExcelReadOnly.vbs", 1, true

Set objFileSys = Nothing
Set objWshShell = Nothing 

