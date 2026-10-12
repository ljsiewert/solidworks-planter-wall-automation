Option Explicit

' Code-behind source: paste into a new UserForm named UserForm_QuoteReport.
Private reportBox As MSForms.TextBox
Private WithEvents closeButton As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Me.Caption = "Quote Quantities - No Model Updates"
    Me.Width = 600
    Me.Height = 520
    Set reportBox = Me.Controls.Add("Forms.TextBox.1", "txtQuoteReport", True)
    With reportBox
        .Left = 12
        .Top = 12
        .Width = 564
        .Height = 430
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
    End With
    Set closeButton = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With closeButton
        .Caption = "Close"
        .Left = 486
        .Top = 454
        .Width = 90
        .Height = 26
        .Default = True
        .Cancel = True
    End With
End Sub

Public Sub SetReport(ByVal text As String)
    reportBox.Text = text
End Sub

Private Sub closeButton_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
