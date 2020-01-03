Attribute VB_Name = "EmailMerge"
Option Private Module
Option Explicit

Public Sub Main()
    Dim Word As Object
    Dim templatePath As String
    Dim union As Range
    
    If Not TestOutlookIsOpen Then Exit Sub                  ' Check if outlook is open
    
    templatePath = ChooseTemplate                           ' Ask user to choose the template
    If templatePath = "" Then Exit Sub                      ' If user cancels selecting a tempalte -> stop here

On Error GoTo Handler
    Set Word = ExecuteMerge(templatePath)                   ' Execute a mail merge with the activesheet and template
    SaveBackup Word
    
    Select Case GetSelectionState
        Case "Send_Selection"
            Set union = Intersect(ActiveSheet.Range("A:A"), Selection.SpecialCells(xlCellTypeVisible))
            CreateEmails union, Word
            
        Case "Send_All"
            CreateEmails ActiveSheet.Range("A:A"), Word
    End Select
    
    Word.Quit SaveChanges:=wdDoNotSaveChanges               ' Quits MS Word without saving any changes
    Set Word = Nothing
    
    Exit Sub

Handler:
    MsgBox "Could not display all emails. The last email that was not displayed will be shown."
    Word.Visible = True
    Set Word = Nothing
End Sub

Private Sub CreateEmails(records As Range, msWord As Object)
    Dim rng As Range
    
    For Each rng In records
        If IsEmpty(rng) Then Exit For
                
        If Not rng.Row = 1 Then
            msWord.ActiveDocument.MailMerge.DataSource.ActiveRecord = rng.Row - 1
            msWord.ActiveDocument.Content.Copy

            CreateEmail _
                toAddress:=rng.Offset(0, 1), _
                subject:=rng.Offset(0, 5), _
                onBehalfOf:=rng.Offset(0, 2), _
                cc:=rng.Offset(0, 3), _
                bcc:=rng.Offset(0, 4), _
                Attachments:=rng.Offset(0, 6)
        End If
    Next rng
End Sub

Private Sub CreateEmail(toAddress As String, subject As String, _
Optional onBehalfOf As String, Optional cc As String, Optional bcc As String, _
Optional Attachments As String)
    Dim olApp, olMail As Object
    Dim attchments() As String
    Dim i As Integer

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .To = toAddress
        .subject = subject
        .GetInspector.WordEditor.Content.Paste

        If Not IsMissing(onBehalfOf) Then .SentOnBehalfOfName = onBehalfOf
        If Not IsMissing(cc) Then .cc = cc
        If Not IsMissing(bcc) Then .bcc = bcc
        If Not IsMissing(Attachments) Then
            attchments = Split(Attachments, ",")

            For i = LBound(attchments) To UBound(attchments)
                .Attachments.Add (Trim(attchments(i)))
            Next i
        End If

        .Display
    End With

    Set olMail = Nothing
    Set olApp = Nothing
End Sub

Private Function ExecuteMerge(templatePath As String) As Object
    Dim Word As Object
    Dim wdTemplate As Object
    
    Set Word = CreateObject("Word.Application")
    Set wdTemplate = Word.Documents.Add(templatePath)      ' Opens the Email Template in Word
    
    With wdTemplate.MailMerge
        .OpenDataSource _
                name:=GetSetting("Data_Source"), _
                AddToRecentFiles:=False, _
                Revert:=False, _
                Connection:="Data Source=" & GetSetting("Data_Source") & ";Mode=Read", _
                SQLStatement:="SELECT * FROM `" & ActiveSheet.name & "$`"
        .ViewMailMergeFieldCodes = False                    ' "Preview Results"
        .Execute                                            ' Execute the merge; if saving backup, you need to execute the merge
    End With
    
    Set ExecuteMerge = Word                                 ' Return the instance of word with merged document as active
End Function

Private Sub SaveBackup(wordApp As Object)
    Dim outFileName As String
    Dim outFile As String
    
    If GetSetting("Save_Merge_File") = True Then
        outFileName = GetSetting("Merged_File_Name") & ".docx"     ' Out File of Mail Mail Merged template for back up
        outFile = GetSetting("Merged_File_Path") & "\" & outFileName                     ' Full path of output file
        
        wordApp.ActiveDocument.SaveAs2 outFile                     ' Save a copy to the backup location
    End If
    
    wordApp.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges   ' Close the merged document
End Sub

Private Function ChooseTemplate() As String
    Dim myFile As Object
         
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    
    With myFile
        .Title = "Choose a template"
        .AllowMultiSelect = False
        .Filters.Clear
        '.Filters.Add "Word Documents", "*.docx;*.doc"
        
        If .Show <> -1 Then Exit Function
        
        ChooseTemplate = .SelectedItems(1)
    End With
End Function

' Used for determining whether or not to send only to the selected rows or to everyone
Private Function GetSelectionState() As String
    With Selection
        If .address = .EntireRow.address Then
            GetSelectionState = "Send_Selection"
        Else
            GetSelectionState = "Send_All"
        End If
    End With
End Function
Private Function TestOutlookIsOpen() As Boolean
    Dim olApp As Object
    
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    
    If olApp Is Nothing Then
        TestOutlookIsOpen = False
        MsgBox "Please open Outlook and try again."
    Else
        TestOutlookIsOpen = True
    End If
End Function
