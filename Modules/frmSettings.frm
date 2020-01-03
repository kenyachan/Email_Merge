VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10575
   OleObjectBlob   =   "frmSettings.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload frmSettings
End Sub

Private Sub btnSelectFolder_Click()
    Dim newPath As String
    
    newPath = ChooseFolder
    
    If newPath <> "" Then
        txtMergedFilePath.Text = newPath
        txtMergedFilePath_AfterUpdate
    End If
End Sub

Private Sub ckBoxBackup_Click()
    Select Case ckBoxBackup.value
        Case True
            txtMergedFilePath.Enabled = True
            txtMergedFileName.Enabled = True
            lblMergedFilePath.Enabled = True
            lblMergedFileName.Enabled = True
            lblExtension.Enabled = True
            btnSelectFolder.Enabled = True
            UpdateSetting "Save_Merge_File", True
        Case False
            txtMergedFilePath.Enabled = False
            txtMergedFileName.Enabled = False
            lblMergedFilePath.Enabled = False
            lblMergedFileName.Enabled = False
            lblExtension.Enabled = False
            btnSelectFolder.Enabled = False
            UpdateSetting "Save_Merge_File", False
    End Select
End Sub

Private Sub txtMergedFileName_AfterUpdate()
    UpdateSetting "Merged_File_Name", txtMergedFileName.Text
End Sub

Private Sub txtMergedFilePath_AfterUpdate()
    UpdateSetting "Merged_File_Path", txtMergedFilePath.Text
End Sub

Private Sub UserForm_Initialize()
    ckBoxBackup.value = GetSetting("Save_Merge_File")
    ckBoxBackup_Click                                   ' So the merge path and names are disabled when save merge file is false
    txtMergedFilePath = GetSetting("Merged_File_Path")
    txtMergedFileName = GetSetting("Merged_File_Name")
End Sub
