Sub NewDate_Macro()

    If Confirmation_MsgBox = vbYes Then
    
        
        Dim startTime As Double
        startTime = Timer
        
        ScreenUpdating (False)
        
            ' Actions to execute
            AddNewDate "Task", 4 ' New column will be in row 4, next to "Task" cell

        ScreenUpdating (True)
        
        Dim EndTime As Double
        EndTime = Timer
        

        Dim totalTime As Double
        totalTime = EndTime - startTime
        
        
        Success_MsgBox totalTime
        

    Else
        
        Withdrawal_MsgBox
    
    End If

End Sub
Private Function Confirmation_MsgBox() As VbMsgBoxResult

    Confirmation_MsgBox = MsgBox("Would you like to run this macro?", _
                        vbYesNo + vbQuestion, _
                        "Confirm Run")
    
End Function
Private Sub Success_MsgBox(value As Double)

        MsgBox "Execution Time: " & Format(Int(value / 60), "00") & ":" & Format(value Mod 60, "00") & " (mm:ss)", _
        vbInformation, _
        "Success! :)"
    
End Sub
Private Sub Withdrawal_MsgBox()

    MsgBox "Macro was not run.", vbInformation, "Ending Process"
    
End Sub
Private Sub ScreenUpdating(status As Boolean)

        Application.ScreenUpdating = status

End Sub
Private Sub AddNewDate(headerToFind As String, rowToSearch As Integer)

    Dim ws As Worksheet
    Dim curDate As Date
    Dim newHeaderName As String
    Dim searchRange, foundHeader As Range
    Dim newColumnNum As Integer

    ' Set current Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set current date
    curDate = Date

    ' Set header name
    newHeaderName = Format(curDate, "dd-mmm-yy") ' Ex. 05-May-24

    ' Set row to search
    Set searchRange = ws.Rows(rowToSearch)  ' Use Set for object assignment

    ' Look for headerToFind
    Set foundHeader = searchRange.Find(What:=headerToFind, LookIn:=xlValues, LookAt:=xlPart)
    
    ' Add New Column
    If Not foundHeader Is Nothing Then

        ' Set where new column will be added
        newColumnNum = foundHeader.Column + 1

        ' Add New Column
        ws.Columns(newColumnNum).Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove

        ' Add Header name
        ws.Cells(rowToSearch, newColumnNum).value = newHeaderName
    Else
        MsgBox "Check sheet's structure", vbInformation, "Column Not Added"
    End If

End Sub


