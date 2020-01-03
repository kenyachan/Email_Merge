Attribute VB_Name = "FormControls"
' Form Controls provides common Subs and Functions used by all forms in this Workbook.
'
' Subs
' ----
' OpenAboutForm(control As IRibbonControl)                  Opens the About page
'
' Functions
' ---------
' ChooseFile() As String                                    Opens a File Explorer dialog where the user can select a File
' ChooseFolder() As String                                  Opens a File Explorer dialog where the user can select a Folder
' FileOrDirExists(pathName As String) As Boolean            Tells you if a pathName exists or not
' ------------------------------------------------------------------------------------------------------------ '
Option Private Module

' These Subs are used by the custom ribbon to open their respective forms
'Public Sub OpenfrmAbout(control As IRibbonControl)
'    frmAbout.Show
'End Sub

'-------------------------------------------------------------------------------------------------------------- '
' Opens a File Explorer dialog which allows you to browse and select multiple file.
' Returns a String of the Full Path of the File
'Public Function ChooseFiles() As String
'    Dim myFile As Object
'    Dim selectedFiles As String
'    Dim i As Integer
'
'    Set myFile = Application.FileDialog(msoFileDialogOpen)
'
'    With myFile
'        .Title = "Choose your files"
'        .AllowMultiSelect = True
'
'        If .Show <> -1 Then Exit Function
'
'        For i = 1 To .SelectedItems.Count
'            If i = 1 Then
'                selectedFiles = .SelectedItems(1)
'            Else
'                selectedFiles = selectedFiles & ", " & .SelectedItems(i)
'            End If
'        Next i
'
'        ChooseFiles = selectedFiles
'    End With
'End Function

' Opens a File Explorer dialog which allows you to browse and select a file.
' Returns a String of the Full Path of the File

' Opens a File Explorer dialog which allows you to browse and select a folder.
Public Function ChooseFolder() As String
    Dim myFolder As Object
    
    Set myFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With myFolder
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then Exit Function
        
        ChooseFolder = .SelectedItems(1)
    End With
End Function

' A function that checks if pathName exists
' Return Values:
'       True if it does
'       False if it doesn't
Public Function FileOrDirExists(pathName As String) As Boolean
    Dim iTemp As Integer
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(pathName)
     
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
    
     'Resume error checking
    On Error GoTo 0
End Function

