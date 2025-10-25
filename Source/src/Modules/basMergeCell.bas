Attribute VB_Name = "basMergeCell"
Option Explicit
'Sub key()
'    Application.OnKey "^%{RIGHT}", "SizeToWidest"
'    Application.OnKey "^%{LEFT}", "SizeToNarrowest"
'    Application.OnKey "^%{UP}", "SizeToShortest"
'    Application.OnKey "^%{DOWN}", "SizeToTallest"
'
'End Sub

'SizeToWidest
Sub SizeToWidest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    Dim blnValue As Boolean
    
    Dim strSel As String
    
    On Error GoTo e
    
    Dim blnOnly1 As Boolean
    
    Application.CutCopyMode = False
    
    If Selection(1).MergeArea.Columns.Count = 1 Then
        blnOnly1 = True
    End If
    
    On Error GoTo e
    
    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.Count).Row
    lngRight = Selection(Selection.Count).Column + 1

    For j = lngRight To Cells.Columns.Count
        blnMerge = False
        blnValue = False
        For i = lngTop To lngBottom

            With Cells(i, j)
            
                If .MergeCells Then
                    blnMerge = True
                End If
            
                If .Value <> "" Or .HasFormula Then
                    blnValue = True
                End If
                
            End With

        Next
        
        If blnMerge = False Then

            If blnValue = True Then
                If MsgBox("拡張先に文字または式がありますが続行しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
                    Exit For
                End If
            End If

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(lngBottom, j - 1))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Columns(r.Columns(2).Column).Insert Shift:=xlToRight
                
                If blnOnly1 Then
                    For i = lngTop To lngBottom
            
                        If .Cells(i, lngLeft).Address = .Cells(i, lngLeft).MergeArea(1).Address Then
                            .Cells(i, lngLeft).MergeArea.Resize(, 2).Merge
                        End If
            
                    Next
                End If

                Set s = .Range(r.Address).Resize(, r.Columns.Count + 1)

                s.Cut Destination:=Range(s.Address)
                
                Set s = Range(strSel)
                s.Resize(, s.Columns.Count + 1).Select

            End With

            Application.ScreenUpdating = True

            Exit For

        End If
    Next
    
    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE

End Sub
'SizeToNarrowest
Sub SizeToNarrowest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    
    Dim strSel As String
    
    On Error GoTo e
    
    If Selection(1).MergeArea.Columns.Count <= 1 Then
        Exit Sub
    End If
    
    Application.CutCopyMode = False

    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.Count).Row
    lngRight = Selection(Selection.Count).Column + 1

    For j = lngRight To Cells.Columns.Count
        blnMerge = False
        For i = lngTop To lngBottom

            If Cells(i, j).MergeCells Then
                blnMerge = True
                Exit For
            End If

        Next
        If blnMerge = False Then

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(lngBottom, j - 1))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Columns(r.Columns(2).Column).Delete Shift:=xlToLeft

                Set s = .Range(r.Address).Resize(, r.Columns.Count - 1)

                s.Cut Destination:=Range(s.Address)
                
            
            End With

            Set s = Range(strSel)
            s.Resize(, s.Columns.Count - 1).Select
            
            Application.ScreenUpdating = True

            Exit For

        End If
    Next

    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE
End Sub
'SizeToTallest
Sub SizeToTallest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    Dim blnValue As Boolean

    Dim strSel As String
    
    On Error GoTo e
    
    Dim blnOnly1 As Boolean
    
    Application.CutCopyMode = False
    
    If Selection(1).MergeArea.Rows.Count = 1 Then
        blnOnly1 = True
    End If
    
    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.Count).Row + 1
    lngRight = Selection(Selection.Count).Column

    For i = lngBottom To Cells.Rows.Count
        blnMerge = False
        blnValue = False
        For j = lngLeft To lngRight

            
            With Cells(i, j)
            
                If .MergeCells Then
                    blnMerge = True
                End If
            
                If .Value <> "" Or .HasFormula Then
                    blnValue = True
                End If
                
            End With

        Next
        
        If blnMerge = False Then

            If blnValue = True Then
                If MsgBox("拡張先に文字または式がありますが続行しますか？", vbOKCancel + vbQuestion, C_TITLE) <> vbOK Then
                    Exit For
                End If
            End If
            
            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(i - 1, lngRight))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Rows(r.Rows(2).Row).Insert Shift:=xlDown
                
                If blnOnly1 Then
                    For j = lngLeft To lngRight
            
                        If .Cells(lngTop, j).Address = .Cells(lngTop, j).MergeArea(1).Address Then
                            .Cells(lngTop, j).MergeArea.Resize(2).Merge
                        End If
            
                    Next
                End If

                Set s = .Range(r.Address).Resize(r.Rows.Count + 1)

                s.Cut Destination:=Range(s.Address)

            End With
            
            Set s = Range(strSel)
            s.Resize(s.Rows.Count + 1).Select

            Application.ScreenUpdating = True

            Exit For

        End If
    Next
    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE
End Sub
'SizeToShortest
Sub SizeToShortest()

    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    Dim i As Long
    Dim j As Long

    Dim blnMerge As Boolean
    
    Dim strSel As String
    
    On Error GoTo e
    
    If Selection(1).MergeArea.Rows.Count <= 1 Then
        Exit Sub
    End If
    
    Application.CutCopyMode = False
    
    strSel = Selection.Address

    lngLeft = Selection(1).Column
    lngTop = Selection(1).Row
    lngBottom = Selection(Selection.Count).Row + 1
    lngRight = Selection(Selection.Count).Column

    For i = lngBottom To Cells.Rows.Count
        blnMerge = False
        For j = lngLeft To lngRight

            If Cells(i, j).MergeCells Then
                blnMerge = True
                Exit For
            End If

        Next
        If blnMerge = False Then

            Dim r As Range
            Dim s As Range
            
            Set r = Range(Cells(lngTop, lngLeft), Cells(i - 1, lngRight))

            Application.ScreenUpdating = False

            With ThisWorkbook.Worksheets("Work")

                .Cells.Clear

                r.Cut Destination:=.Range(r.Address)

                .Rows(r.Rows(2).Row).Delete Shift:=xlUp

                Set s = .Range(r.Address).Resize(r.Rows.Count - 1)

                s.Cut Destination:=Range(s.Address)
                
            End With

            Set s = Range(strSel)
            s.Resize(s.Rows.Count - 1).Select
            
            Application.ScreenUpdating = True

            Exit For

        End If
    Next
    Exit Sub
e:
    MsgBox "他の結合セルに影響するため実行できません。", vbOKOnly + vbExclamation, C_TITLE
End Sub
