VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "UIProgressSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SETTINGS_APP As String = "PlanterWallAutomation"
Private Const SETTINGS_SECTION As String = "StageSecondsV1"
Private form As UserForm_AutomationProgress
Private stages As Collection
Private durations() As Double
Private current As Long
Private started As Date
Private stageStarted As Date

Public Sub BeginRun(ByVal plan As Collection)
    Set stages = plan
    ReDim durations(1 To stages.Count)
    started = Now
    stageStarted = started
    current = 0
    Set form = New UserForm_AutomationProgress
    form.Show vbModeless
End Sub

Public Sub NextStage()
    Dim remaining As Double
    Dim i As Long
    Dim saved As String
    If current > 0 Then
        durations(current) = DateDiff("s", stageStarted, Now)
        UITraceEnd "Stage: " & CStr(stages(current))
    End If
    current = current + 1
    If current > stages.Count Then
        Err.Raise vbObjectError + 2300, "UIProgressSession", "Progress plan has fewer stages than the workflow."
    End If
    stageStarted = Now
    UITraceBegin "Stage: " & CStr(stages(current))
    remaining = 0
    For i = current To stages.Count
        saved = GetSetting(SETTINGS_APP, SETTINGS_SECTION, CStr(stages(i)), "")
        If Len(saved) = 0 Then
            remaining = -1
            Exit For
        End If
        If Not IsNumeric(saved) Then
            Err.Raise vbObjectError + 2300, "UIProgressSession", "Invalid saved timing for stage: " & CStr(stages(i))
        End If
        If CDbl(saved) < 0 Then
            Err.Raise vbObjectError + 2300, "UIProgressSession", "Negative saved stage timing."
        End If
        remaining = remaining + CDbl(saved)
    Next i
    form.UpdateProgress CStr(stages(current)), current - 1, stages.Count, _
                        DateDiff("s", started, Now), remaining
    DoEvents
End Sub

Public Sub RestoringAfterFailure()
    If form Is Nothing Then Exit Sub
    If current < 1 Then Exit Sub
    On Error GoTo Failed
    form.UpdateProgress "Workflow stopped; closing table editor and releasing protection...", _
                        current - 1, stages.Count, DateDiff("s", started, Now), -1
    DoEvents
    Exit Sub
Failed:
    MsgBox "Could not update the progress window while handling a workflow error." & _
           vbCrLf & Err.Description, vbExclamation, "Progress Display"
End Sub

Public Sub Complete()
    If current <> stages.Count Then
        Err.Raise vbObjectError + 2300, "UIProgressSession", "Progress completed before all planned stages ran."
    End If
    durations(current) = DateDiff("s", stageStarted, Now)
    UITraceEnd "Stage: " & CStr(stages(current))
    form.UpdateProgress "Workflow returned; review export dialogs/logs for results.", _
                        stages.Count, stages.Count, DateDiff("s", started, Now), 0
    SaveTimings
    CloseWindow
End Sub

Private Sub SaveTimings()
    Dim i As Long
    On Error GoTo Failed
    For i = 1 To stages.Count
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, CStr(stages(i)), CStr(durations(i))
    Next i
    Exit Sub
Failed:
    MsgBox "Automation finished, but progress timing history could not be saved." & _
           vbCrLf & Err.Description, vbExclamation, "Progress Timing"
End Sub

Public Sub CloseWindow()
    If Not form Is Nothing Then Unload form
    Set form = Nothing
End Sub
