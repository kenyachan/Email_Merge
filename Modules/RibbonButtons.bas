Attribute VB_Name = "RibbonButtons"
Option Private Module
Option Explicit

Private Sub CreateEmails(control As IRibbonControl)
    EmailMerge.Main
End Sub

Private Sub OpenfrmAbout(control As IRibbonControl)
    frmAbout.Show
End Sub

Private Sub OpenfrmSettings(control As IRibbonControl)
    frmSettings.Show
End Sub

Private Sub AddSequentialAttachments(control As IRibbonControl)
    Attachments.AddSequentialAttachment Selection
End Sub

Private Sub AddAttachments(control As IRibbonControl)
    Attachments.AddAttachments Selection
End Sub

Private Sub ToggleOnBehalfOf(control As IRibbonControl)
    FieldToggles.ToggleEmailField ("On_Behalf_Of")
End Sub

Private Sub ToggleCC(control As IRibbonControl)
    FieldToggles.ToggleEmailField ("CC")
End Sub

Private Sub ToggleBCC(control As IRibbonControl)
    FieldToggles.ToggleEmailField ("BCC")
End Sub

Private Sub ToggleAttachments(control As IRibbonControl)
    FieldToggles.ToggleEmailField ("Attachments")
End Sub

Private Sub CreateNewSheet(control As IRibbonControl)
    NewSheets.NewSheet
End Sub



