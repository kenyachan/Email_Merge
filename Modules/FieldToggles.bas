Attribute VB_Name = "FieldToggles"
Option Private Module
Option Explicit

'Public Sub ToggleEmailField(field As String)
'    Select Case field
'        Case "On_Behalf_Of"
'            UpdateSetting field, HideColumn(3)
'        Case "CC"
'            UpdateSetting field, HideColumn(4)
'        Case "BCC"
'            UpdateSetting field, HideColumn(5)
'        Case "Attachments"
'            UpdateSetting field, HideColumn(7)
'    End Select
'
'End Sub

Public Sub ToggleEmailField(field As String)
    Select Case field
        Case "On_Behalf_Of"
            HideColumn (3)
        Case "CC"
            HideColumn (4)
        Case "BCC"
            HideColumn (5)
        Case "Attachments"
            HideColumn (7)
    End Select

End Sub

Private Function HideColumn(columnNumber As Integer) As Boolean
    With ActiveSheet.Columns(columnNumber).EntireColumn
        If .Hidden Then
            .Hidden = False
            HideColumn = True
        Else
            .Hidden = True
            HideColumn = False
        End If
    End With
End Function
