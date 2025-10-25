'------------------------------------------------------------------------------
' Set the Excel file cursor to the home position
'
' ExcelSetHomePosition.vbs
' Version 1.0.0
'
' Copyright (c) 2015 Y.Watanabe
'
' This software is released under the MIT License.
' http://opensource.org/licenses/mit-license.php
'------------------------------------------------------------------------------
' Operation confirmation: Windows 7 + Excel 2010 / Windows 8 + Excel 2013
'------------------------------------------------------------------------------
' for Used
' (1) Place this script in the folder of the Excel file to set the home position.
' (2) Rewrite the script's "extension" and "read password" as necessary.
' (3) Execute the script.
' (4) Display the results in a text file.
'
'------------------------------------------------------------------------------
    Option Explicit

    Dim objFs, strMsg, SH
    Dim objDic, XL, WB, FL, LogName
    dim varPatterns, strKey, varPass, p
    Dim IE
    Dim strTitle
    
    strTitle = "Home position setting"
    
    If MsgBox("Set the home position for the Excel files under the same folder." & vbCrLf & "Are you sure?？" & VbCrLf & VbCrLf & "☆Promise☆" & vbCrLf & "Please back up your Excel file in advance.", vbYesNo + vbQuestion, strTitle) = vbNo Then 
        WScript.Quit 
    End IF

    Set IE = WScript.CreateObject("InternetExplorer.Application")
 
    IE.Navigate "about:blank"
    Do While IE.busy
        WScript.Sleep(100)
    Loop
    Do While IE.Document.readyState <> "complete"
        WScript.Sleep(100)
    Loop
    IE.Document.body.innerHTML = "<b id=""msg"">Setting home position<br>Please wait...</b>"
    IE.AddressBar = False
    IE.ToolBar = False
    IE.StatusBar = False
    IE.Height = 120
    IE.Width = 300
    IE.Left = 0
    IE.Top = 0
    IE.Document.Title = strTitle
    IE.Visible = True
    
    On Error Resume Next

    Set objFs =  WScript.CreateObject("Scripting.FileSystemObject")
    Set objDic = WScript.CreateObject("Scripting.Dictionary")
    
    '--------------------------------------------------------------
    ' Describe the extension to be processed using regular expressions
    '--------------------------------------------------------------
    varPatterns = Array("\.xls$", "\.xlsx$", "\.xlsm$")
    
    '--------------------------------------------------------------
    ' If you have a read password, enter it here (multiple passwords can be specified)
    '--------------------------------------------------------------
    varPass = Array("", "", "")
    
    FileSearch objFs, objFs.GetParentFolderName(WScript.ScriptFullName), varPatterns, objDic

    LogName = objFs.GetBaseName(WScript.ScriptFullName) & ".txt"
    Set FL = objFs.CreateTextFile(LogName)

    FL.WriteLine "☆=Start setting home position (" & Now() & ")☆="
    FL.WriteLine "Number of files processed:" & objDic.Count

    If objDic.Count > 0 Then
        
        Set XL = WScript.CreateObject("Excel.Application")

        For Each strKey In objDic.Keys
        
           'If password is specified
            For Each p In varPass
                Err.Clear
                Set WB = XL.WorkBooks.Open(objDic(strKey),,False,,p,"",True,,,False)
                If Err.Number = 0 Then
                    Exit For
                End If
            Next
            
            Select Case True
                Case Err.Number <> 0
                    FL.WriteLine "error => " & objDic(strKey)
                    FL.WriteLine "          " & Err.Description
                    
                Case WB.ReadOnly 
  	                FL.WriteLine "error => " & objDic(strKey)
                    FL.WriteLine "          Book is read-only"
                    
                Case Else
                    setAllA1 WB

                    XL.DisplayAlerts = False
                    WB.Save
                
                    If Err.Number <> 0 Or WB.Saved = False Then
                        FL.WriteLine "error => " & objDic(strKey)
                        FL.WriteLine "          " & Err.Description
                    Else
                        FL.WriteLine "Processed => " & objDic(strKey)
                    End If
                
                    XL.DisplayAlerts = True
            End Select
            
            'If there is an instance Close
            If Not IsNothing(WB) Then
                WB.Close
                Set WB = Nothing
            End If
        Next

        XL.Quit

        Set XL = Nothing

    End If

    FL.WriteLine "☆= Home position setting finished (" & Now() & ")☆="
    FL.Close
    Set FL = Nothing

    Set objDic = Nothing
    Set objFs =  Nothing

    With CreateObject("Shell.Application")
        .ShellExecute(LogName)
    End With

    IE.Quit
    'MsgBox "Processing completed.", vbInformation + VbOkOnly, strTitle

'--------------------------------------------------------------
'Set the selection position of all sheets to A1
'--------------------------------------------------------------
Sub setAllA1(WB)

    Dim WS
    Dim WD

    For Each WS In WB.Worksheets
        If WS.visible Then
            WS.Activate
'Problem where multiple cells are not deselected when setting A1 #65 #66
            WS.Range("A1").Select
            WB.Windows(1).ScrollRow = 1
            WB.Windows(1).ScrollColumn = 1
            WB.Windows(1).Zoom = 100
        End If
    Next

    'Make it the first image displayed.
    For Each WS In WB.Worksheets
        If WS.visible  Then
            WS.Select
            Exit For
        End If
    Next

End Sub

'--------------------------------------------------------------
'　Subfolder search
'--------------------------------------------------------------
Private Sub FileSearch(objFs, strPath, varPatterns, objDic)

    Dim objfld
    Dim objfl
    Dim objSub
    Dim f, objRegx
    
    Set objfld = objFs.GetFolder(strPath)

    'Get file name
    For Each objfl In objfld.files
    
        Dim blnFind
        blnFind = False

	    Set objRegx = CreateObject("VBScript.RegExp")
        For Each f In varPatterns
            objRegx.Pattern = f
            If objRegx.Test(objfl.name) Then
                blnFind = True
                Exit For
            End If
        Next
	    Set objRegx = Nothing
        
        If blnFind Then
            objDic.Add objFs.BuildPath(objfl.ParentFolder.Path, objfl.name), objFs.BuildPath(objfl.ParentFolder.Path, objfl.name)
        End If
    Next
    
    'With subfolder search
    For Each objSub In objfld.SubFolders
        FileSearch objFs, objSub.Path, varPatterns, objDic
    Next

End Sub

