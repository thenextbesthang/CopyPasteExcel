Attribute VB_Name = "Module1"
Option Explicit
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
Sub copy()
Attribute copy.VB_ProcData.VB_Invoke_Func = "d\n14"
    copyPaste
    updateDates

End Sub
Function copyPaste()

    Set sourceSheet = ActiveSheet
    


    'add new sheet
    With sourceSheet.Parent
        Set targetSheet = .Sheets.Add(After:=Sheets(Sheets.Count))
    End With
    
    'copy and set page break borders
    
    sourceSheet.Range("A1:T44").copy
    targetSheet.Range("a1").PasteSpecial xlPasteAll
    targetSheet.PageSetup.Orientation = xlLandscape
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    targetSheet.Columns("L").ColumnWidth = 10
 
  '  targetSheet.Range("C29:E29").Formula = "=7-sum(C26:C28)"
  '  targetSheet.Range("K29:M29").Formula = "=7-sum(K26:K28)"
    'copy signatures
    copyPasteSignatures sourceSheet, targetSheet
    
    
End Function

Function updateDates()
    Dim oldDate As Date
    Dim newDate As Date
    
    'first find the date where timesheet information has most recently been entered
    oldDate = sourceSheet.Range("Q10")
    targetSheet.Range("C10") = Format(DateSerial(Year(oldDate), Month(oldDate), Day(oldDate) + 1), "mm/dd/yyyy")
    ActiveWindow.Zoom = 100
    targetSheet.Name = Month(CDate(Cells(10, 3))) & "." & Day(CDate(Cells(10, 3))) & "-" & Month(CDate(Cells(10, 17))) & "." & Day(CDate(Cells(10, 17))) & "." & Year(CDate(Cells(10, 17)))

End Function

Function copyPasteSignatures(sourceSheet As Worksheet, targetSheet As Worksheet)
    Dim pic As Shape, rng As Range
    For Each pic In sourceSheet.Shapes
           If pic.Type = msoPicture Then
           pic.copy
          With targetSheet
                .Select
                .Range(pic.TopLeftCell.Address).Select
                .Paste
            End With
        Selection.Placement = xlMoveAndSize
        End If
    Next pic
End Function



