Option Explicit

' Code-behind: paste into a new UserForm named UserForm_AutomationProgress.
Private stageLabel As MSForms.Label
Private countLabel As MSForms.Label
Private elapsedLabel As MSForms.Label
Private estimateLabel As MSForms.Label
Private barFill As MSForms.Label

Private Sub UserForm_Initialize()
    Dim track As MSForms.Label
    Dim note As MSForms.Label
    Me.Caption = "Planter automation progress"
    Me.Width = 460
    Me.Height = 215
    Set stageLabel = AddLabel("lblStage", 16, 16, 420, 32)
    Set countLabel = AddLabel("lblCount", 16, 52, 420, 18)
    Set track = AddLabel("lblTrack", 16, 76, 420, 16)
    track.BackColor = RGB(220, 220, 220)
    Set barFill = AddLabel("lblFill", 16, 76, 0, 16)
    barFill.BackColor = RGB(45, 125, 200)
    Set elapsedLabel = AddLabel("lblElapsed", 16, 103, 420, 18)
    Set estimateLabel = AddLabel("lblEstimate", 16, 127, 420, 18)
    Set note = AddLabel("lblNote", 16, 153, 420, 36)
    note.Caption = "Updates between stages. SolidWorks/Excel calls may hold the display." & vbCrLf & _
                   "Progress is milestones, not time or individual files."
End Sub

Public Sub UpdateProgress(ByVal stage As String, ByVal completed As Long, _
                          ByVal total As Long, ByVal elapsed As Double, _
                          ByVal remaining As Double)
    stageLabel.Caption = stage
    countLabel.Caption = CStr(completed) & " of " & CStr(total) & _
                         " stages completed (" & CStr(Int(100 * completed / total)) & "%)"
    barFill.Width = 420 * completed / total
    elapsedLabel.Caption = "Elapsed at last update: " & UIProgressTime(elapsed)
    If remaining < 0 Then
        estimateLabel.Caption = "Estimated remaining: unavailable until stages have timing history"
    Else
        estimateLabel.Caption = "Estimated remaining: ~" & UIProgressTime(remaining) & " (previous run timings)"
    End If
    Me.Repaint
End Sub

Private Function AddLabel(ByVal name As String, ByVal left As Single, ByVal top As Single, _
                          ByVal width As Single, ByVal height As Single) As MSForms.Label
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", name, True)
    label.Left = left
    label.Top = top
    label.Width = width
    label.Height = height
    Set AddLabel = label
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "The workflow is still running. This window is not a cancellation control.", _
               vbInformation, "Automation Progress"
    End If
End Sub
