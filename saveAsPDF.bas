Attribute VB_Name = "Module2"
Sub saveAsPDF()
    Set FSO = CreateObject("Scripting.FileSystemObject")

       
        
        For Each wks In ThisWorkbook.Worksheets
       '  Set wks = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
         'MsgBox (wks.Name)
         
         
            myDate = wks.Range("P10").Value + 6

         
            Select Case Month(myDate)
                Case Is = 1
                    folderPath = "January"
                Case Is = 2
                    folderPath = "February"
                Case Is = 3
                    folderPath = "March"
                Case Is = 4
                    folderPath = "April"
                Case Is = 5
                    folderPath = "May"
                Case Is = 6
                    folderPath = "June"
                Case Is = 7
                    folderPath = "July"
                Case Is = 8
                    folderPath = "August"
                Case Is = 9
                    folderPath = "September"
                Case Is = 10
                    folderPath = "October"
                Case Is = 11
                    folderPath = "November"
                Case Is = 12
                    folderPath = "December"
            End Select
            
            folderPath = ThisWorkbook.Path & "\" & folderPath
            filePath = folderPath & "\" & wks.Range("D5").Value & "_" & Replace(myDate, "/", "_")
            
            
            
            If FSO.FolderExists(folderPath) = False Then MkDir folderPath
       
            
                wks.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=filePath, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False

            
        Next wks
End Sub
