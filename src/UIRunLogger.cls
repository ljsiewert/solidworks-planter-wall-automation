VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "UIRunLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private filePath As String
Private fso As Object
Private spans As Object
Private totals As Object
Private started As Double
Private enabled As Boolean

Public Property Get Path() As String
    Path = filePath
End Property

Public Sub BeginRun(ByVal folderPath As String)
    Dim stream As Object
    Dim candidate As String
    Dim suffix As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set spans = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    candidate = fso.BuildPath(folderPath, "Run_" & Format$(Now, "yyyy-mm-dd_HH-nn-ss"))
    filePath = candidate & ".log"
    Do While fso.FileExists(filePath)
        suffix = suffix + 1
        filePath = candidate & "_" & CStr(suffix) & ".log"
    Loop
    ' Creation failures stop before model edits. Each later append closes its stream.
    Set stream = fso.CreateTextFile(filePath, False, True)
    stream.WriteLine "Planter automation timing log v1 (UTF-16)"
    stream.WriteLine "Timestamp" & vbTab & "ElapsedSeconds" & vbTab & "Event" & vbTab & _
                     "Operation" & vbTab & "DurationSeconds" & vbTab & "Details"
    stream.Close
    started = UIRunClock()
    enabled = True
    Record "RUN_START", "RunWithUI", "", "Elapsed time includes dialogs/user waits; nested durations overlap."
End Sub

Public Sub BeginSpan(ByVal name As String)
    If Not enabled Then Exit Sub
    spans(name) = UIRunClock()
    Record "BEGIN", name, "", ""
End Sub

Public Sub EndSpan(ByVal name As String)
    Dim seconds As Double
    If Not enabled Then Exit Sub
    If Not spans.Exists(name) Then
        Err.Raise vbObjectError + 2500, "UIRunLogger", "Timing operation was not started: " & name
    End If
    seconds = UIRunClock() - CDbl(spans(name))
    spans.Remove name
    If totals.Exists(name) Then
        totals(name) = CDbl(totals(name)) + seconds
    Else
        totals.Add name, seconds
    End If
    Record "END", name, UIInchesText(seconds), ""
End Sub

Public Sub Detail(ByVal name As String, ByVal text As String)
    Record "DETAIL", name, "", text
End Sub

Public Sub Finish(ByVal result As String, ByVal message As String)
    Dim key As Variant
    If Not enabled Then Exit Sub
    For Each key In spans.Keys
        Record "INCOMPLETE", CStr(key), UIInchesText(UIRunClock() - CDbl(spans(key))), result
    Next key
    For Each key In totals.Keys
        Record "SUMMARY", CStr(key), UIInchesText(CDbl(totals(key))), "Nested operations overlap; do not sum all rows."
    Next key
    Record "RUN_" & result, "RunWithUI", UIInchesText(UIRunClock() - started), message
    enabled = False
End Sub

Private Sub Record(ByVal eventName As String, ByVal name As String, _
                   ByVal seconds As String, ByVal text As String)
    Dim stream As Object
    Dim failure As String
    If Not enabled Then Exit Sub
    On Error GoTo Failed
    Set stream = fso.OpenTextFile(filePath, 8, False, -1)
    stream.WriteLine Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbTab & _
        UIInchesText(UIRunClock() - started) & vbTab & eventName & vbTab & _
        CleanField(name) & vbTab & seconds & vbTab & CleanField(text)
    stream.Close
    Exit Sub
Failed:
    failure = Err.Description
    enabled = False
    MsgBox "Runtime logging stopped; the model workflow will continue." & vbCrLf & _
           filePath & vbCrLf & failure, vbExclamation, "Run Log Error"
    If Not stream Is Nothing Then CloseAfterError stream
End Sub

Private Sub CloseAfterError(ByVal stream As Object)
    On Error GoTo Failed
    stream.Close
    Exit Sub
Failed:
    MsgBox "Could not close the run log stream: " & Err.Description, vbExclamation, "Run Log Error"
End Sub

Private Function CleanField(ByVal text As String) As String
    CleanField = Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
End Function
