Attribute VB_Name = "Attachments"
Option Private Module
Option Explicit

Public Sub AddAttachments(userSelection As Range)
    Dim rng As Range
    Dim attachmentsStr As String
    Dim i As Integer
    
    If Not IsValidSelection(userSelection) Then Exit Sub
    
    attachmentsStr = ToString(ChooseFiles)
    If attachmentsStr = "" Then Exit Sub
    
    For Each rng In userSelection
        If IsEmpty(rng) Then
            rng.value = attachmentsStr
        Else
            If Not rng.Row = 1 Then
                rng.value = rng.value & ", " & attachmentsStr
            End If
        End If
    Next rng
End Sub

Public Sub AddSequentialAttachment(userSelection As Range)
    Dim i As Integer
    Dim attachment As FileDialogSelectedItems
    
    If Not IsValidSelection(userSelection) Then Exit Sub
    
    Set attachment = ChooseFiles
    
    If userSelection(1).Row = 1 Then
        For i = 1 To attachment.Count
            userSelection(i + 1) = attachment(i)
        Next i
    Else
        For i = 1 To Attachments.Count
            userSelection(i) = attachment(i)
        Next i
    End If
End Sub

Private Function IsValidSelection(userSelection As Range) As Boolean
    IsValidSelection = True
    
    If Intersect(Selection, ActiveSheet.Range("G:G")) Is Nothing Then
        MsgBox "Please selected a range within the Attachment(s) column.", _
                vbExclamation
                
        IsValidSelection = False
        Exit Function
    End If
    
    If Selection.address = Selection.EntireColumn.address Then
        MsgBox "You have selected the entire column. Please select a " & _
                "limited range to prevent Excel from not responding.", _
                vbCritical, "Warning!"
                
        IsValidSelection = False
        Exit Function
    End If
End Function

Private Function ChooseFiles() As FileDialogSelectedItems
    Dim myFile As Object
    
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    
    With myFile
        .Title = "Choose files to attach"
        .AllowMultiSelect = True
        
        If .Show <> -1 Then Exit Function
        
        Set ChooseFiles = .SelectedItems
    End With
End Function

Private Function ToString(SelectedItems As FileDialogSelectedItems) As String
    Dim selectedItemsStr As String
    Dim i As Integer
    
    If SelectedItems Is Nothing Then Exit Function
    
    For i = 1 To SelectedItems.Count
        If i = 1 Then
            selectedItemsStr = SelectedItems(1)
        Else
            selectedItemsStr = selectedItemsStr & ", " & SelectedItems(i)
        End If
    Next i
    
    ToString = selectedItemsStr
End Function
