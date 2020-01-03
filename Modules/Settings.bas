Attribute VB_Name = "Settings"
' The Settings Module provides all the Subs and Functions for adding, storing, retrieving and removing settings
'
' Subs
' ----
' OpenSettingsForm(control As IRibbonControl)               Opens the Settings page
' ListSettings()                                            Lists all the setting in Immediate Window
' UpdateSetting(setting As String, setting As String)       Saves the setting into the setting
'
' Functions
' ---------
' RemoveSettingKey(setting As String) As Boolean            Removes 'setting' and associated setting
' AddSettingKey(setting As String) As Boolean               Adds 'setting'
' GetSetting(setting As String) As String                   Gets the setting associated with 'setting'

Option Private Module

Public Sub ListCommands()
    Debug.Print "----------------------------------------------------"
    Debug.Print "ListSettings"
    Debug.Print "UpdateSettings(setting As String, value As String)"
    Debug.Print "RemoveSetting(setting As String)"
    Debug.Print "AddSetting(setting As String)"
    Debug.Print "GetSetting(setting As String)"
End Sub

' For use in Immediate Window to see what settings there are
Public Sub ListSettings()
    Dim rng As Range
    
    For Each rng In Worksheets("Settings").Range("A:A")
        If IsEmpty(rng) Then Exit For
        
        Debug.Print rng.value & " - " & rng.Offset(0, 1).value
    Next rng
End Sub

' Updates the setting value
Public Sub UpdateSetting(setting As String, value As String)
    Dim rng As Range
    
    Worksheets("Settings").Unprotect
    
    For Each rng In Worksheets("Settings").Range("A:A")
        If IsEmpty(rng.value) Then Exit For
        
        If rng.value = setting Then
            rng.Offset(0, 1) = value
            Exit For
        End If
    Next rng
    
    Worksheets("Settings").Protect
End Sub

' Removes the setting and associated value
Public Function RemoveSetting(setting As String) As Boolean
    Dim rng As Range
    Dim confirmation As Integer
    
    Worksheets("Settings").Unprotect
    
    For Each rng In Worksheets("Settings").Range("A:A")
        If IsEmpty(rng) Then
            RemoveSetting = False
            MsgBox """" & setting & """" & " does not exist in the list of settings." & _
            vbNewLine & vbNewLine & "Nothing was removed.", vbExclamation
            Exit For
        End If
        
        If rng = setting Then
            confirmation = MsgBox("Warning!" & vbNewLine & vbNewLine & _
                    "You are about to remove a setting and it's associated values." & _
                    "This may cause errors and should only be done if you know what you're doing." & _
                    vbNewLine & vbNewLine & "Are you sure you want to remove " & """" & setting & """" & "?", _
                    vbYesNo + vbCritical)
            Select Case confirmation
                Case vbYes
                    rng.EntireRow.Delete
                    'rng.Offset(0, 1).Clear
                    RemoveSetting = True
                    Exit For
                Case vbNo
                    Exit For
            End Select
        End If
    Next rng
    Worksheets("Settings").Protect
End Function

' Adds the setting to the list of settings (values need to be "updated" separately)
Public Function AddSetting(setting As String) As Boolean
    Dim rng As Range
    
    Worksheets("Settings").Unprotect
    
    For Each rng In Worksheets("Settings").Range("A:A")
        If rng.value = setting Then
            AddSetting = False
            Exit For
        End If
    
        If IsEmpty(rng) Then
            rng.value = setting
            AddSetting = True
            Exit For
        End If
    Next rng

    Worksheets("Settings").Protect
End Function

' Gets the value of setting
Public Function GetSetting(setting As String) As String
    Dim rng As Range
    
    For Each rng In Worksheets("Settings").Range("A:A")
        If IsEmpty(rng) Then Exit For
        
        If rng.value = setting Then
            GetSetting = rng.Offset(0, 1).value
            Exit For
        End If
    Next rng
End Function
