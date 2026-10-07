VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "UIReferenceSaveGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const GUARD_ERROR As Long = vbObjectError + 2400
Private app As SldWorks.SldWorks
Private paths As Collection
Private readOnlyStates As Collection
Private originalDiscard As Boolean
Private preferenceChanged As Boolean
Private appliedCount As Long

Public Sub BeginGuard(ByVal application As SldWorks.SldWorks, ByVal template As SldWorks.ModelDoc2)
    Dim dependencies As Variant
    Dim references As Object
    Dim documents As Variant
    Dim document As SldWorks.ModelDoc2
    Dim i As Long
    Dim path As String

    Set app = application
    Set paths = New Collection
    Set readOnlyStates = New Collection
    Set references = CreateObject("Scripting.Dictionary")
    references.CompareMode = vbTextCompare
    dependencies = template.Extension.GetDependencies(True, True, False, False, False)
    If Not IsArray(dependencies) Then
        Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", "Cannot discover template references; automatic discard is not safe."
    End If
    If (UBound(dependencies) - LBound(dependencies) + 1) Mod 2 <> 0 Then
        Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", "Unexpected template dependency list."
    End If
    For i = LBound(dependencies) To UBound(dependencies) Step 2
        path = CStr(dependencies(i + 1))
        If Len(path) > 0 Then references(path) = True
    Next i
    documents = app.GetDocuments
    If Not IsArray(documents) Then
        Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", "Cannot inspect open documents before automatic discard."
    End If

    ' Validate every affected loaded reference before changing any document state.
    For i = LBound(documents) To UBound(documents)
        Set document = documents(i)
        path = document.GetPathName
        If references.Exists(path) And StrComp(path, template.GetPathName, vbTextCompare) <> 0 Then
            If document.GetSaveFlag <> False Then
                Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", _
                    "This template reference already has unsaved changes:" & vbCrLf & path & _
                    vbCrLf & "Save or discard those changes manually before running UI mode."
            End If
            paths.Add path
            readOnlyStates.Add CBool(document.IsOpenedReadOnly)
        End If
    Next i
    originalDiscard = CBool(app.GetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefNoPromptOrSave))
    preferenceChanged = True
    app.SetUserPreferenceToggle swUserPreferenceToggle_e.swExtRefNoPromptOrSave, True
    If app.GetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefNoPromptOrSave) = False Then
        Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", "Could not enable discard of read-only reference changes."
    End If
    For i = 1 To paths.Count
        appliedCount = i
        Set document = app.GetOpenDocumentByName(CStr(paths(i)))
        If Not document Is Nothing Then
            If document.SetReadOnlyState(True) = False Then
                Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", "Cannot protect template reference: " & CStr(paths(i))
            End If
            If document.IsOpenedReadOnly = False Then
                Err.Raise GUARD_ERROR, "UIReferenceSaveGuard", "Template reference did not become read-only: " & CStr(paths(i))
            End If
        End If
    Next i
End Sub

Public Sub Restore()
    Dim i As Long
    Dim issues As String
    If app Is Nothing Then Exit Sub
    ' Each restoration handles its own error so one failure cannot skip the others.
    For i = 1 To appliedCount
        issues = issues & RestoreDocument(CStr(paths(i)), CBool(readOnlyStates(i)))
    Next i
    issues = issues & RestorePreference()
    If Len(issues) > 0 Then
        Err.Raise GUARD_ERROR, "UIReferenceSaveGuard.Restore", _
            "Could not fully restore reference-save protection:" & vbCrLf & issues
    End If
    appliedCount = 0
End Sub

Private Function RestoreDocument(ByVal path As String, ByVal readOnly As Boolean) As String
    Dim document As SldWorks.ModelDoc2
    On Error GoTo Failed
    Set document = app.GetOpenDocumentByName(path)
    ' Suppression may have unloaded the document; its in-memory flag then no longer exists.
    If document Is Nothing Then Exit Function
    If CBool(document.IsOpenedReadOnly) <> readOnly Then
        If document.SetReadOnlyState(readOnly) = False Then
            Err.Raise GUARD_ERROR, , "SolidWorks rejected the original read-only state."
        End If
    End If
    If CBool(document.IsOpenedReadOnly) <> readOnly Then
        Err.Raise GUARD_ERROR, , "Original read-only state was not restored."
    End If
    Exit Function
Failed:
    RestoreDocument = path & ": " & Err.Description & vbCrLf
End Function

Private Function RestorePreference() As String
    If Not preferenceChanged Then Exit Function
    On Error GoTo Failed
    app.SetUserPreferenceToggle swUserPreferenceToggle_e.swExtRefNoPromptOrSave, originalDiscard
    If CBool(app.GetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefNoPromptOrSave)) <> originalDiscard Then
        Err.Raise GUARD_ERROR, , "Original discard preference was not restored."
    End If
    preferenceChanged = False
    Exit Function
Failed:
    RestorePreference = "External References discard preference: " & Err.Description & vbCrLf
End Function
