Attribute VB_Name = "NewSheets"
Option Private Module
Option Explicit

Public Sub NewSheet()
    Dim sht As Worksheet
    Dim newSheetName As String
    'Get sheet name input from user
    newSheetName = InputBox("What would you like to call this sheet?", "New Sheet")
    
    If newSheetName = "" Then Exit Sub  'if user cancels
    
    If SheetExists(newSheetName) Then       ' if sheet with name exists
        MsgBox "The sheet " & """" & newSheetName & """" & " already exists." & vbNewLine _
        & "Please create a new sheet with another name.", vbExclamation
        Exit Sub
    End If
    
    Set sht = Sheets.Add(After:=Sheets(Worksheets.Count))
    sht.Name = newSheetName
    
    BuildSheet sht
End Sub

Private Function SheetExists(sheetToFind As String) As Boolean
    Dim sht As Worksheet
    
    SheetExists = False
    
    For Each sht In Worksheets
        If sheetToFind = sht.Name Then
            SheetExists = True
            Exit Function
        End If
    Next sht
End Function

Private Sub BuildSheet(sht As Worksheet)
    Dim rng As Range
    Dim i As Integer
    
    sht.Range("A1").value = "Name"
    sht.Range("B1").value = "Email Address"
    sht.Range("C1").value = "On Behalf Of"
    sht.Range("D1").value = "CC"
    sht.Range("E1").value = "BCC"
    sht.Range("F1").value = "Subject"
    sht.Range("G1").value = "Attachment(s)"
    sht.Range("G1").AddComment "Separate each attachment path with comma's."
    
    i = 1
    For Each rng In sht.Range("H1:S1")
        rng.value = "Merge_Field_" & i
        i = i + 1
    Next rng

    sht.Range("A:G").ColumnWidth = 27.86
    sht.Range("H:S").ColumnWidth = 16.43
    
    With sht.Range("1:1")
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Interior.Color = 14277081
        .RowHeight = 20.25
        .Borders(xlEdgeBottom).Weight = -4138
        .VerticalAlignment = xlVAlignCenter
    End With
    
    With sht
        .Range("A1:G1").Font.Bold = True
        
        .Range("C:C").EntireColumn.Hidden = True
        .Range("D:D").EntireColumn.Hidden = True
        .Range("E:E").EntireColumn.Hidden = True
        .Range("G:G").EntireColumn.Hidden = True
        
        .Range("H2").Select
    End With
    
    ActiveWindow.FreezePanes = True
End Sub
