VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AboutText_Click()

End Sub

Private Sub CloseBtn_Click()
    Unload frmAbout
End Sub

Private Sub CopyRightText1_Click()

End Sub

Private Sub UserForm_Initialize()
    VersionText.Caption = "Version " & GetSetting("Version") & " (Build " & GetSetting("Build") & ")"
End Sub
