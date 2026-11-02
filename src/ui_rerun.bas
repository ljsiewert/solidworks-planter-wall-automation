Attribute VB_Name = "ui_rerun"
Option Explicit

' Import into the SAME VBA project as packngo1 and userform_automation.
' Rerun from a packed output assembly into a new Pack and Go folder, keeping custom cuts.
Private Const RERUN_ERROR As Long = vbObjectError + 2500

' Silent: any failure simply leaves the UI in normal template mode.
Public Sub UIDetectRerunSource()
    Dim model As SldWorks.ModelDoc2
    Dim manager As SldWorks.EquationMgr
    Dim names As Variant
    Dim values(0 To 10) As String
    Dim i As Long

    rerunPath = ""
    rerunInputs = Empty
    On Error GoTo NotOutput
    Set model = Application.SldWorks.ActiveDoc
    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub
    If Len(model.GetPathName) = 0 Then Exit Sub
    If StrComp(packngo1.GetFileName(model.GetPathName), "planter_assembly.SLDASM", vbTextCompare) = 0 Then Exit Sub
    If model.GetDesignTable Is Nothing Then Exit Sub
    Set manager = model.GetEquationMgr
    names = Array("WALL_COUNT", "SIDE", "RETURN_TYPE", "OVERALL_LENGTH", "OVERALL_WIDTH", "OVERALL_HEIGHT", _
        "OVERALL_MATERIAL", "OVERALL_THICKNESS", "CONTINUATION", "DEPTH_OVERRIDE", "DEPTH_OVERRIDE_DIMENSION")
    For i = 0 To 10
        values(i) = UIGlobalValue(manager, CStr(names(i)))
        If values(i) = "ERROR" Then Exit Sub
    Next i
    rerunInputs = UIRerunFormInputs(values)
    rerunPath = model.GetPathName
    Exit Sub
NotOutput:
    rerunPath = ""
    rerunInputs = Empty
End Sub

' Model globals -> form values: walls, side, return, L, W, H, material, thickness, mounts, override, base width.
Public Function UIRerunFormInputs(ByVal values As Variant) As Variant
    Dim walls As Long
    Dim sideLabel As String
    Dim returnLabel As String
    Dim materialLabel As String
    Dim thickness As Double
    Dim baseWidth As Double
    walls = CLng(CDbl(values(0)))
    sideLabel = "N/A"
    If walls = 2 Then
        Select Case CLng(CDbl(values(1)))
            Case 1: sideLabel = "Right"
            Case 2: sideLabel = "Left"
        End Select
    End If
    Select Case CLng(CDbl(values(2)))
        Case 1: returnLabel = "Single return 1 inch"
        Case 2: returnLabel = "Single return 2 inches"
        Case 3: returnLabel = "Double return"
        Case Else: returnLabel = "No return"
    End Select
    If CLng(CDbl(values(6))) = 2 Then materialLabel = "Borcon Weathering Steel" Else materialLabel = "Mild Steel"
    thickness = CDbl(values(7))
    If CDbl(values(9)) = 1 Then baseWidth = Round(CDbl(values(10)) / 2 - thickness, 4)
    UIRerunFormInputs = Array(CStr(walls), sideLabel, returnLabel, CDbl(values(3)), CDbl(values(4)), CDbl(values(5)), _
        materialLabel, IIf(Abs(thickness - 0.1875) < 0.0001, "3/16 inch", "1/4 inch"), _
        CDbl(values(8)) = 1, CDbl(values(9)) = 1, baseWidth)
End Function

' Removes a trailing custom suffix (-C1, -C2, ...) from a file base name.
Public Function UIRerunStripSuffix(ByVal base As String) As String
    Dim dashPos As Long
    Dim tail As String
    UIRerunStripSuffix = base
    dashPos = InStrRev(base, "-")
    If dashPos = 0 Then Exit Function
    tail = Mid$(base, dashPos + 1)
    If Len(tail) < 2 Then Exit Function
    If UCase(Left$(tail, 1)) <> "C" Then Exit Function
    If Not UIRerunAllDigits(Mid$(tail, 2)) Then Exit Function
    UIRerunStripSuffix = Left$(base, dashPos - 1)
End Function

Public Function UIRerunAllDigits(ByVal text As String) As Boolean
    Dim i As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        If Mid$(text, i, 1) < "0" Or Mid$(text, i, 1) > "9" Then Exit Function
    Next i
    UIRerunAllDigits = True
End Function

' "C9175" -> "C", "S5950" -> "S", "SL5950" -> "SL"; anything else -> "".
Public Function UIRerunLengthCode(ByVal segment As String) As String
    segment = UCase(segment)
    Select Case Left$(segment, 2)
        Case "SL", "SR", "SC"
            If UIRerunAllDigits(Mid$(segment, 3)) Then
                UIRerunLengthCode = Left$(segment, 2)
                Exit Function
            End If
    End Select
    Select Case Left$(segment, 1)
        Case "C", "S"
            If UIRerunAllDigits(Mid$(segment, 2)) Then UIRerunLengthCode = Left$(segment, 1)
    End Select
End Function

' Length code of parts(index); also accepts a negative length ("S-37") that Split broke into "S" and "37".
Public Function UIRerunCodeAt(ByVal parts As Variant, ByVal index As Long) As String
    Dim code As String
    If index > UBound(parts) Then Exit Function
    code = UIRerunLengthCode(CStr(parts(index)))
    If Len(code) = 0 And index < UBound(parts) Then
        Select Case UCase(CStr(parts(index)))
            Case "C", "S", "SL", "SR", "SC"
                If UIRerunAllDigits(CStr(parts(index + 1))) Then code = UCase(CStr(parts(index)))
        End Select
    End If
    UIRerunCodeAt = code
End Function

' Maps a packed output file name back to the template file it came from (the names packngo1.GetNewBaseName
' understands), or "" when the file is not a recognized planter file (it then keeps its name).
Public Function UIRerunTemplateFile(ByVal fileName As String, ByVal masterBase As String) As String
    Dim dotPos As Long
    Dim ext As String
    Dim base As String
    Dim parts As Variant
    Dim role As String
    Dim isPart As Boolean
    Dim isAssembly As Boolean
    Dim partRole As Boolean

    dotPos = InStrRev(fileName, ".")
    If dotPos < 2 Then Exit Function
    ext = LCase(Mid$(fileName, dotPos))
    Select Case ext
        Case ".sldprt": isPart = True
        Case ".sldasm": isAssembly = True
        Case ".slddrw"
        Case Else: Exit Function
    End Select
    base = Left$(fileName, dotPos - 1)
    If StrComp(base, masterBase, vbTextCompare) = 0 Or _
       StrComp(UIRerunStripSuffix(base), UIRerunStripSuffix(masterBase), vbTextCompare) = 0 Then
        If Not isPart Then UIRerunTemplateFile = "planter_assembly" & ext
        Exit Function
    End If
    parts = Split(UCase(UIRerunStripSuffix(base)), "-")
    partRole = True
    Select Case CStr(parts(UBound(parts)))
        Case "CG": role = "centergusset"
        Case "MG": role = "mountinggusset"
        Case "LW": role = "leftsidewall"
        Case "RW": role = "rightsidewall"
        Case "BP"
            Select Case UIRerunCodeAt(parts, 0)
                Case "C": role = "baseplatecenter"
                Case "S": role = "baseplateside"
            End Select
        Case "W"
            Select Case UIRerunCodeAt(parts, 1)
                Case "C": role = "planterwallcenter"
                Case "S": role = "planterwallside"
                Case "SL": role = "planterwallsideleft"
                Case "SR": role = "planterwallsideright"
                Case "SC": role = "planterwallsidecenter"
            End Select
        Case Else
            partRole = False
            Select Case UIRerunCodeAt(parts, 1)
                Case "C": role = "planter_center_asm"
                Case "S": role = "planter_side_asm"
                Case "SL": role = "planter_side_asm_left"
                Case "SR": role = "planter_side_asm_right"
                Case "SC": role = "planter_side_asm_center"
            End Select
    End Select
    If Len(role) = 0 Then Exit Function
    If partRole And isAssembly Then Exit Function
    If Not partRole And isPart Then Exit Function
    UIRerunTemplateFile = role & ext
End Function

' A rerun always creates a new custom folder, including when dimensions are unchanged.
Public Sub UIRerunBuildNames(ByVal oldMaster As String, ByVal parentRoot As String)
    Dim fso As Object
    Dim serial As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    packngo1.customSuffix = ""
    If Not packngo1.Step3_BuildNames() Then
        Err.Raise RERUN_ERROR, "UIRerunBuildNames", "Could not build the new part numbers."
    End If
    UIApplyActualDimensionNames
    serial = 1
    Do While fso.FolderExists(parentRoot & packngo1.masterName & "-C" & CStr(serial)) Or _
             StrComp(oldMaster, packngo1.masterName & "-C" & CStr(serial), vbTextCompare) = 0
        serial = serial + 1
    Loop
    UIAppendNameSuffix "-C" & CStr(serial)
End Sub

' Planter model files in the folder. Dir$ skips hidden files; "~$" lock files that SolidWorks keeps
' for open documents are also excluded explicitly. Returns the count; names(0 .. count-1) are filled.
Private Function UIRerunListModelFiles(ByVal folder As String, ByRef names() As String) As Long
    Dim name As String
    Dim count As Long
    ReDim names(0 To 15)
    name = Dir$(folder & "*.*", vbNormal Or vbReadOnly Or vbArchive)
    Do While Len(name) > 0
        If UIRerunIsPlanterFileName(name) Then
            If count > UBound(names) Then ReDim Preserve names(0 To 2 * count)
            names(count) = name
            count = count + 1
        End If
        name = Dir$()
    Loop
    UIRerunListModelFiles = count
End Function

Public Function UIRerunIsPlanterFileName(ByVal fileName As String) As Boolean
    If Left$(fileName, 2) = "~$" Then Exit Function
    If InStrRev(fileName, ".") = 0 Then Exit Function
    Select Case LCase(Mid$(fileName, InStrRev(fileName, ".") + 1))
        Case "sldprt", "sldasm", "slddrw": UIRerunIsPlanterFileName = True
    End Select
End Function

' Role of a planter document base name (template name without extension), or "" when it is not a planter file.
Public Function UIRerunNameRole(ByVal base As String) As String
    Dim role As String
    Dim parts As Variant
    role = UIRerunTemplateFile(base & ".SLDASM", "")
    If Len(role) = 0 Then role = UIRerunTemplateFile(base & ".SLDPRT", "")
    If Len(role) = 0 Then
        parts = Split(UCase(UIRerunStripSuffix(base)), "-")
        If UBound(parts) = 5 Then
            If UIRerunCodeAt(parts, 2) = "C" And UIRerunCodeAt(parts, 3) = "S" Then role = "planter_assembly.sldasm"
        End If
    End If
    If Len(role) > 0 Then role = Left$(role, InStrRev(role, ".") - 1)
    UIRerunNameRole = role
End Function

' Current file base name for the planter role of name, or "" when name is not a planter name or its shape
' (number of dash-separated fields) differs from the current file for that role.
Public Function UIRerunCurrentBase(ByVal name As String, ByVal roles As Variant, ByVal bases As Variant, _
                                   ByVal count As Long) As String
    Dim role As String
    Dim k As Long
    role = UIRerunNameRole(name)
    If Len(role) = 0 Then Exit Function
    For k = 0 To count - 1
        If StrComp(roles(k), role, vbTextCompare) = 0 Then
            If UBound(Split(UIRerunStripSuffix(name), "-")) = UBound(Split(UIRerunStripSuffix(bases(k)), "-")) Then
                UIRerunCurrentBase = bases(k)
            End If
            Exit Function
        End If
    Next k
End Function

' Equations link to other documents by name ("X@1-C4200-1400-22-C1<1>.Assembly", "1-SL3600-1400-22-C1<3>.Assembly").
' Renaming files does not update those names, so any planter name that is not current is replaced by the
' current name for the same role. roles/bases are parallel arrays of the current files (count entries).
Public Function UIRerunRepairEquation(ByVal text As String, ByVal roles As Variant, ByVal bases As Variant, _
                                      ByVal count As Long) As String
    Dim result As String
    Dim position As Long
    Dim finish As Long
    Dim token As String
    Dim ch As String
    Dim replacement As String
    Dim dashPos As Long
    Dim current As String

    position = 1
    Do While position <= Len(text)
        ch = Mid$(text, position, 1)
        result = result & ch
        finish = 0
        If ch = """" Or ch = "@" Then
            finish = position + 1
            Do While finish <= Len(text)
                ch = Mid$(text, finish, 1)
                If ch = """" Or ch = "@" Or ch = "<" Or ch = "." Or ch = " " Then Exit Do
                finish = finish + 1
            Loop
            If finish > Len(text) Or finish = position + 1 Then
                finish = 0
            ElseIf Mid$(text, finish, 1) <> "<" And Mid$(text, finish, 1) <> "." Then
                finish = 0
            End If
        End If
        If finish > 0 Then
            token = Mid$(text, position + 1, finish - position - 1)
            replacement = token
            ' Component references carry an instance number ("1-C4200-1400-22-C1-1"); keep it.
            dashPos = InStrRev(token, "-")
            current = ""
            If dashPos > 1 Then
                If UIRerunAllDigits(Mid$(token, dashPos + 1)) Then
                    current = UIRerunCurrentBase(Left$(token, dashPos - 1), roles, bases, count)
                    If Len(current) > 0 Then
                        If StrComp(Left$(token, dashPos - 1), current, vbTextCompare) <> 0 Then replacement = current & Mid$(token, dashPos)
                    End If
                End If
            End If
            If Len(current) = 0 Then
                current = UIRerunCurrentBase(token, roles, bases, count)
                If Len(current) > 0 Then
                    If StrComp(token, current, vbTextCompare) <> 0 Then replacement = current
                End If
            End If
            result = result & replacement
            position = finish
        Else
            position = position + 1
        End If
    Loop
    UIRerunRepairEquation = result
End Function

' SolidWorks may reformat spacing when it stores an equation.
Public Function UIRerunSqueeze(ByVal text As String) As String
    UIRerunSqueeze = Replace(text, " ", "")
End Function

' Repairs name-based equation links in every planter part/assembly in the folder; returns the number of
' equations changed. Documents that were not loaded (suppressed components) are opened invisibly and closed.
Private Function UIRerunRepairLinks(ByVal app As SldWorks.SldWorks, ByVal folder As String) As Long
    Dim names() As String
    Dim roles() As String
    Dim bases() As String
    Dim fileCount As Long
    Dim roleCount As Long
    Dim changed As Long
    Dim i As Long
    Dim k As Long
    Dim ext As String
    Dim base As String
    Dim role As String
    Dim docType As Long
    Dim document As SldWorks.ModelDoc2
    Dim manager As SldWorks.EquationMgr
    Dim opened As Boolean
    Dim dirty As Boolean
    Dim oldText As String
    Dim newText As String
    Dim readBack As String
    Dim errors As Long
    Dim warnings As Long

    UITraceBegin "Repair equation links"
    fileCount = UIRerunListModelFiles(folder, names)
    ReDim roles(0 To fileCount)
    ReDim bases(0 To fileCount)
    For i = 0 To fileCount - 1
        ext = LCase$(Mid$(names(i), InStrRev(names(i), ".")))
        If ext <> ".slddrw" Then
            base = Left$(names(i), InStrRev(names(i), ".") - 1)
            role = UIRerunNameRole(base)
            If Len(role) > 0 Then
                roles(roleCount) = role
                bases(roleCount) = base
                roleCount = roleCount + 1
            End If
        End If
    Next i
    For i = 0 To fileCount - 1
        ext = LCase$(Mid$(names(i), InStrRev(names(i), ".")))
        If ext <> ".slddrw" Then
            If ext = ".sldasm" Then docType = swDocASSEMBLY Else docType = swDocPART
            Set document = app.GetOpenDocumentByName(folder & names(i))
            opened = False
            If document Is Nothing Then
                app.DocumentVisible False, docType
                Set document = app.OpenDoc6(folder & names(i), docType, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)
                app.DocumentVisible True, docType
                If document Is Nothing Then
                    Err.Raise RERUN_ERROR, "UIRerunRepairLinks", "Cannot open " & names(i) & " to repair its equations. SolidWorks error: " & CStr(errors)
                End If
                opened = True
            End If
            Set manager = document.GetEquationMgr
            dirty = False
            For k = 0 To manager.GetCount - 1
                oldText = manager.Equation(k)
                newText = UIRerunRepairEquation(oldText, roles, bases, roleCount)
                If StrComp(oldText, newText, vbBinaryCompare) <> 0 Then
                    manager.Equation(k) = newText
                    readBack = manager.Equation(k)
                    If StrComp(UIRerunSqueeze(readBack), UIRerunSqueeze(newText), vbTextCompare) <> 0 Then
                        UITraceDetail "Equation repair rejected", names(i) & ": old=" & oldText & " | attempted=" & newText & " | returned=" & readBack & " | status=" & CStr(manager.Status)
                        manager.Equation(k) = oldText
                        If opened Then app.CloseDoc folder & names(i)
                        Err.Raise RERUN_ERROR, "UIRerunRepairLinks", "SolidWorks rejected the repaired equation in " & names(i) & "." & vbCrLf & vbCrLf & _
                            "Old: " & oldText & vbCrLf & "Attempted: " & newText & vbCrLf & "SolidWorks returned: " & readBack
                    End If
                    UITraceDetail "Repaired equation", names(i) & ": " & oldText & " -> " & newText
                    dirty = True
                    changed = changed + 1
                End If
            Next k
            If dirty Then
                If document.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, errors, warnings) = False Then
                    Err.Raise RERUN_ERROR, "UIRerunRepairLinks", "Cannot save " & names(i) & " after repairing its equations. SolidWorks error: " & CStr(errors)
                End If
            End If
            If opened Then app.CloseDoc folder & names(i)
        End If
    Next i
    UITraceDetail "Repaired equation links", CStr(changed) & " equation(s) in " & folder
    UITraceEnd "Repair equation links"
    UIRerunRepairLinks = changed
End Function

Private Function UIRerunIsInFolder(ByVal path As String, ByVal folder As String) As Boolean
    UIRerunIsInFolder = (StrComp(Left$(path, InStrRev(path, "\")), folder, vbTextCompare) = 0)
End Function

' Unsaved edits in the output (for example custom cuts) must reach disk before files are closed and renamed.
Private Sub UIRerunSaveOpenWork(ByVal folder As String)
    Dim documents As Variant
    Dim document As SldWorks.ModelDoc2
    Dim dirty As New Collection
    Dim list As String
    Dim i As Long
    Dim errors As Long
    Dim warnings As Long

    documents = packngo1.swApp.GetDocuments
    If Not IsArray(documents) Then Exit Sub
    For i = LBound(documents) To UBound(documents)
        Set document = documents(i)
        If UIRerunIsInFolder(document.GetPathName, folder) Then
            If document.GetSaveFlag Then
                dirty.Add document
                list = list & packngo1.GetFileName(document.GetPathName) & vbCrLf
            End If
        End If
    Next i
    If dirty.Count = 0 Then Exit Sub
    If MsgBox("These planter files have unsaved changes:" & vbCrLf & vbCrLf & list & vbCrLf & _
              "Save them now and continue? Choose No to stop without changing anything.", _
              vbQuestion Or vbYesNo, "Save Planter Changes") <> vbYes Then
        Err.Raise RERUN_ERROR, "UIRerunSaveOpenWork", "Stopped so unsaved planter changes are not lost. Nothing was changed."
    End If
    For Each document In dirty
        If document.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, errors, warnings) = False Then
            Err.Raise RERUN_ERROR, "UIRerunSaveOpenWork", "Could not save " & document.GetPathName & ". SolidWorks error: " & CStr(errors)
        End If
    Next document
End Sub

Private Function UIRerunProgressPlan(ByVal modelChanges As Boolean) As Collection
    Dim plan As New Collection
    If modelChanges Then
        plan.Add "Checking export macros and unsaved planter files"
        plan.Add "Preparing a local working copy"
        plan.Add "Updating the local design table and rebuilding"
        plan.Add "Planning Pack and Go copies and new part numbers"
        plan.Add "Copying the output into a new folder"
        plan.Add "Closing the local working copy"
        plan.Add "Updating, saving and reopening the new output"
    Else
        plan.Add "Checking export macro files"
    End If
    If exportDXF Then plan.Add "Running DXF component exporter"
    If exportSTEP Then plan.Add "Running STEP component exporter"
    If exportPDFAssemblies Then plan.Add "Running assembly PDF exporter"
    If exportPDFComponents Then plan.Add "Running component PDF exporter"
    Set UIRerunProgressPlan = plan
End Function

' SolidWorks' What's Wrong list (document: feature, error code) for the assembly and every loaded
' subassembly/part, for rebuild failure messages; the full list goes to the run log.
Private Function UIRerunWhatsWrong(ByVal model As SldWorks.ModelDoc2) As String
    Dim assembly As SldWorks.AssemblyDoc
    Dim components As Variant
    Dim component As SldWorks.Component2
    Dim document As SldWorks.ModelDoc2
    Dim seen As Object
    Dim text As String
    Dim lines As Long
    Dim i As Long
    On Error GoTo Unavailable
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    UIRerunWhatsWrongOf model, text, lines
    seen(model.GetPathName) = True
    Set assembly = model
    components = assembly.GetComponents(False)
    If IsArray(components) Then
        For i = LBound(components) To UBound(components)
            Set component = components(i)
            Set document = component.GetModelDoc2
            If Not document Is Nothing Then
                If Not seen.Exists(document.GetPathName) Then
                    seen(document.GetPathName) = True
                    UIRerunWhatsWrongOf document, text, lines
                End If
            End If
        Next i
    End If
    If lines = 0 Then
        UIRerunWhatsWrong = vbCrLf & "(SolidWorks reported no What's Wrong items; check the tree for red/yellow markers.)"
        Exit Function
    End If
    UITraceDetail "What's Wrong", Replace(text, vbCrLf, "; ")
    If lines > 15 Then
        UIRerunWhatsWrong = vbCrLf & "What's Wrong (" & CStr(lines) & " items, full list in the run log):" & _
                            Left$(text, UIRerunNthLineEnd(text, 15))
    Else
        UIRerunWhatsWrong = vbCrLf & "What's Wrong:" & text
    End If
    Exit Function
Unavailable:
    UIRerunWhatsWrong = vbCrLf & "(What's Wrong list unavailable: " & Err.Description & ")"
End Function

Private Sub UIRerunWhatsWrongOf(ByVal document As SldWorks.ModelDoc2, ByRef text As String, ByRef lines As Long)
    Dim features As Variant
    Dim codes As Variant
    Dim warnings As Variant
    Dim feature As SldWorks.Feature
    Dim manager As SldWorks.EquationMgr
    Dim equationText As String
    Dim equationStatus As Long
    Dim i As Long
    ' GetWhatsWrong can return no items even when inherited globals fail to evaluate.
    Set manager = document.GetEquationMgr
    For i = 0 To manager.GetCount - 1
        equationText = manager.Equation(i)
        equationStatus = manager.Status
        If equationStatus < 0 Then
            text = text & vbCrLf & "  " & packngo1.GetFileName(document.GetPathName) & _
                   ": equation " & CStr(i + 1) & " evaluation failed (status " & CStr(equationStatus) & "): " & equationText
            lines = lines + 1
        End If
    Next i
    If document.Extension.GetWhatsWrongCount = 0 Then Exit Sub
    If document.Extension.GetWhatsWrong(features, codes, warnings) = False Then Exit Sub
    If Not IsArray(features) Then Exit Sub
    For i = LBound(features) To UBound(features)
        Set feature = features(i)
        text = text & vbCrLf & "  " & packngo1.GetFileName(document.GetPathName) & ": " & feature.Name & _
               " (" & feature.GetTypeName2 & ") " & IIf(warnings(i), "warning ", "error ") & CStr(codes(i))
        lines = lines + 1
    Next i
End Sub

Private Function UIRerunNthLineEnd(ByVal text As String, ByVal n As Long) As Long
    Dim position As Long
    Dim k As Long
    position = 1
    For k = 1 To n
        position = InStr(position + 1, text, vbCrLf)
        If position = 0 Then
            UIRerunNthLineEnd = Len(text)
            Exit Function
        End If
    Next k
    UIRerunNthLineEnd = position - 1
End Function

Private Sub UIRerunExports(ByVal model As SldWorks.ModelDoc2, ByVal progress As UIProgressSession)
    If exportDXF Then
        progress.NextStage
        UIRunExport model, "dxfcomponents.swp", "RunDXFComponentDrawings"
    End If
    If exportSTEP Then
        progress.NextStage
        UIRunExport model, "stepcomponents.swp", "RunSTEPComponentExport"
    End If
    If exportPDFAssemblies Then
        progress.NextStage
        UIRunExport model, "pdfassemblies.swp", "RunPDFAssemblyDrawings"
    End If
    If exportPDFComponents Then
        progress.NextStage
        UIRunExport model, "pdfcomponents.swp", "RunPDFComponentDrawings"
    End If
End Sub

' Runs only the selected exporters on the open output assembly; nothing is changed or renamed.
Public Sub UIRunExportsOnly()
    Dim model As SldWorks.ModelDoc2
    Dim progress As UIProgressSession
    Dim logPath As String
    Dim failureMessage As String

    On Error GoTo Failed
    logPath = UIStartRunLog()
    Set model = packngo1.swModel
    UITraceDetail "Exports only", model.GetPathName
    UITraceDetail "Exports", "DXF=" & CStr(exportDXF) & "; STEP=" & CStr(exportSTEP) & _
        "; PDF assemblies=" & CStr(exportPDFAssemblies) & "; PDF components=" & CStr(exportPDFComponents)
    Set progress = New UIProgressSession
    progress.BeginRun UIRerunProgressPlan(False)
    progress.NextStage
    UICheckExportFiles
    packngo1.destFolder = Left$(model.GetPathName, InStrRev(model.GetPathName, "\"))
    UIRerunExports model, progress
    progress.Complete
    UIEndRunLog "FINISHED", "Exports only returned; review exporter status/logs for results."
    packngo1.swApp.Frame.SetStatusBarText "Selected exports returned for " & model.GetPathName & _
        ". Review exporter status/logs. Run log: " & logPath
    Exit Sub
Failed:
    failureMessage = Err.Description
    If Not progress Is Nothing Then progress.CloseWindow
    UITraceDetail "Workflow failure", failureMessage
    UIEndRunLog "FAILED", failureMessage
    MsgBox failureMessage & vbCrLf & vbCrLf & "No planter files were changed." & vbCrLf & "Run log: " & logPath, _
           vbCritical, "Exports Stopped"
End Sub

' Only loaded documents inside the specified output folder are closed.
Private Sub UIRerunCloseOutput(ByVal app As SldWorks.SldWorks, ByVal masterPath As String)
    Dim documents As Variant
    Dim paths As New Collection
    Dim document As SldWorks.ModelDoc2
    Dim path As Variant
    Dim folder As String
    Dim i As Long
    folder = Left$(masterPath, InStrRev(masterPath, "\"))
    documents = app.GetDocuments
    If IsArray(documents) Then
        For i = LBound(documents) To UBound(documents)
            Set document = documents(i)
            If UIRerunIsInFolder(document.GetPathName, folder) Then paths.Add document.GetPathName
        Next i
    End If
    Set document = Nothing
    documents = Empty
    app.CloseDoc masterPath
    For Each path In paths
        If Not app.GetOpenDocumentByName(CStr(path)) Is Nothing Then app.CloseDoc CStr(path)
        If Not app.GetOpenDocumentByName(CStr(path)) Is Nothing Then
            Err.Raise RERUN_ERROR, "UIRerunCloseOutput", "Could not close " & CStr(path) & "."
        End If
    Next path
End Sub

Private Sub UIRerunCheckCopiedModel(ByVal model As SldWorks.ModelDoc2, ByVal folder As String)
    Dim assembly As SldWorks.AssemblyDoc
    Dim components As Variant
    Dim component As SldWorks.Component2
    Dim path As String
    Dim i As Long
    If Not UIRerunIsInFolder(model.GetPathName, folder) Then
        Err.Raise RERUN_ERROR, "UIRerunCheckCopiedModel", "The active assembly is not in the new output folder."
    End If
    Set assembly = model
    components = assembly.GetComponents(False)
    If Not IsArray(components) Then Exit Sub
    For i = LBound(components) To UBound(components)
        Set component = components(i)
        path = component.GetPathName
        If Len(path) > 0 Then
            If Not UIRerunIsInFolder(path, folder) Then
                Err.Raise RERUN_ERROR, "UIRerunCheckCopiedModel", "The copied assembly still uses a component outside its folder:" & vbCrLf & path
            End If
        End If
    Next i
End Sub

Private Function UIRerunCreateStage(ByVal app As SldWorks.SldWorks, ByVal sourceFolder As String, _
                                    ByVal masterFile As String, ByRef stageFolder As String, _
                                    ByRef stagedFiles As Collection) As String
    Dim fso As Object
    Dim names() As String
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim dependencies As Variant
    Dim referencePath As String
    Dim target As String
    Dim stageRoot As String
    Dim stagedPath As String
    Dim verified As Boolean
    Dim k As Long
    Dim updatedDependencies As Variant
    Dim redirectReportedSuccess As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stagedFiles = New Collection
    count = UIRerunListModelFiles(sourceFolder, names)
    If count = 0 Then Err.Raise RERUN_ERROR, "UIRerunCreateStage", "The source output contains no model files."
    stageRoot = Environ$("LOCALAPPDATA")
    If Len(stageRoot) = 0 Then
        Err.Raise RERUN_ERROR, "UIRerunCreateStage", "LOCALAPPDATA is not configured for local staging."
    End If
    If Not fso.FolderExists(stageRoot) Then
        Err.Raise RERUN_ERROR, "UIRerunCreateStage", "The local staging root does not exist:" & vbCrLf & stageRoot
    End If
    If fso.GetDrive(fso.GetDriveName(stageRoot)).DriveType <> 2 Then
        Err.Raise RERUN_ERROR, "UIRerunCreateStage", "The staging root must be on a local fixed drive:" & vbCrLf & stageRoot
    End If
    stageFolder = fso.BuildPath(stageRoot, "PlanterRerun-" & fso.GetTempName) & "\"
    fso.CreateFolder Left$(stageFolder, Len(stageFolder) - 1)
    UITraceDetail "Local staging folder", stageFolder
    For i = 0 To count - 1
        target = stageFolder & names(i)
        stagedFiles.Add target
        fso.CopyFile sourceFolder & names(i), target, False
        fso.GetFile(target).Attributes = fso.GetFile(target).Attributes And Not vbReadOnly
    Next i
    ' Paths change, but document names must not: filesystem renames break equation context.
    For i = 0 To count - 1
        stagedPath = stageFolder & names(i)
        If Not app.GetOpenDocumentByName(stagedPath) Is Nothing Then
            Err.Raise RERUN_ERROR, "UIRerunCreateStage", "The staged document must be closed before redirecting references:" & vbCrLf & stagedPath
        End If
        dependencies = app.GetDocumentDependencies2(stagedPath, False, False, False)
        If IsArray(dependencies) Then
            For j = LBound(dependencies) To UBound(dependencies) - 1 Step 2
                referencePath = CStr(dependencies(j + 1))
                If UIRerunIsInFolder(referencePath, sourceFolder) Then
                    target = stageFolder & packngo1.GetFileName(referencePath)
                    If Not fso.FileExists(target) Then
                        Err.Raise RERUN_ERROR, "UIRerunCreateStage", "A source reference is missing from the local copy:" & vbCrLf & referencePath
                    End If
                    UITraceDetail "Redirect staged reference", stagedPath & " | " & referencePath & " -> " & target
                    redirectReportedSuccess = app.ReplaceReferencedDocument(stagedPath, referencePath, target)
                    UITraceDetail "Reference replacement return", stagedPath & ": " & CStr(redirectReportedSuccess)
                    verified = False
                    updatedDependencies = app.GetDocumentDependencies2(stagedPath, False, False, False)
                    If IsArray(updatedDependencies) Then
                        For k = LBound(updatedDependencies) To UBound(updatedDependencies) - 1 Step 2
                            If StrComp(CStr(updatedDependencies(k + 1)), referencePath, vbTextCompare) = 0 Then
                                Err.Raise RERUN_ERROR, "UIRerunCreateStage", "The staged document still references the original file:" & vbCrLf & _
                                    stagedPath & vbCrLf & referencePath
                            End If
                            If StrComp(CStr(updatedDependencies(k + 1)), target, vbTextCompare) = 0 Then verified = True
                        Next k
                    End If
                    If Not verified Then
                        Err.Raise RERUN_ERROR, "UIRerunCreateStage", "Cannot verify the redirected staged reference:" & vbCrLf & _
                            stagedPath & vbCrLf & "From: " & referencePath & vbCrLf & "To: " & target & _
                            vbCrLf & "API return: " & CStr(redirectReportedSuccess)
                    End If
                    If redirectReportedSuccess = False Then
                        UITraceDetail "Reference replacement warning", "API returned False, but saved dependencies verified the replacement: " & _
                            stagedPath & " | " & referencePath & " -> " & target
                    End If
                ElseIf Not UIRerunIsInFolder(referencePath, stageFolder) Then
                    UITraceDetail "Existing external staging reference", names(i) & ": " & referencePath
                End If
            Next j
        End If
    Next i
    UIRerunCreateStage = stageFolder & masterFile
    If Not fso.FileExists(UIRerunCreateStage) Then
        Err.Raise RERUN_ERROR, "UIRerunCreateStage", "The source master assembly was not copied."
    End If
End Function

Private Function UIRerunCleanStage(ByVal stageFolder As String, ByVal stagedFiles As Collection) As String
    Dim fso As Object
    Dim path As Variant
    On Error GoTo Failed
    If Len(stageFolder) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(stageFolder) Then Exit Function
    ' Delete only the exact files created by this run; RmDir refuses nonempty folders.
    If Not stagedFiles Is Nothing Then
        For Each path In stagedFiles
            If Not UIRerunIsInFolder(CStr(path), stageFolder) Then
                Err.Raise RERUN_ERROR, "UIRerunCleanStage", "Cleanup path is outside the staging folder."
            End If
            If fso.FileExists(CStr(path)) Then Kill CStr(path)
        Next path
    End If
    RmDir Left$(stageFolder, Len(stageFolder) - 1)
    Exit Function
Failed:
    UIRerunCleanStage = "Local staging cleanup failed: " & Err.Description & vbCrLf & stageFolder
    UITraceDetail "Staging cleanup warning", UIRerunCleanStage
End Function

Public Sub UIRunRerun()
    Dim app As SldWorks.SldWorks
    Dim source As SldWorks.ModelDoc2
    Dim packed As SldWorks.ModelDoc2
    Dim pack As SldWorks.PackAndGo
    Dim tableSession As UIDesignTableSession
    Dim progress As UIProgressSession
    Dim fso As Object
    Dim usedNames As Object
    Dim includedPaths As Object
    Dim expected As Variant
    Dim copiedExpected As Variant
    Dim documents As Variant
    Dim statuses As Variant
    Dim oldNames() As String
    Dim newNames() As String
    Dim plan As String
    Dim count As Long
    Dim i As Long
    Dim role As String
    Dim sourcePath As String
    Dim sourceFolder As String
    Dim stageFolder As String
    Dim stagedPath As String
    Dim stagedFiles As Collection
    Dim cleanupWarning As String
    Dim parentRoot As String
    Dim oldMaster As String
    Dim finalFolder As String
    Dim finalPath As String
    Dim configurationName As String
    Dim stage As String
    Dim state As String
    Dim logPath As String
    Dim failureMessage As String
    Dim openWarnings As String
    Dim errors As Long
    Dim warnings As Long
    Dim inputIndex As Long

    On Error GoTo Failed
    stage = "checking"
    logPath = UIStartRunLog()
    Set app = packngo1.swApp
    Set source = packngo1.swModel
    Set fso = CreateObject("Scripting.FileSystemObject")
    sourcePath = source.GetPathName
    sourceFolder = Left$(sourcePath, InStrRev(sourcePath, "\"))
    parentRoot = UIOutputRoot(outputFolder)
    oldMaster = fso.GetBaseName(sourcePath)
    configurationName = source.ConfigurationManager.ActiveConfiguration.Name
    UITraceDetail "Rerun source", sourcePath
    UITraceDetail "SolidWorks revision", app.RevisionNumber
    For inputIndex = 0 To 9
        UITraceDetail "Input " & UITableInputAddress(inputIndex), CStr(designTableInputs(inputIndex))
    Next inputIndex
    UITraceDetail "Exports", "DXF=" & CStr(exportDXF) & "; STEP=" & CStr(exportSTEP) & _
        "; PDF assemblies=" & CStr(exportPDFAssemblies) & "; PDF components=" & CStr(exportPDFComponents)
    Set progress = New UIProgressSession
    progress.BeginRun UIRerunProgressPlan(True)

    progress.NextStage
    UICheckExportFiles
    UIRerunSaveOpenWork sourceFolder

    progress.NextStage
    stage = "staging"
    UITraceBegin "Prepare local rerun copy"
    Set source = Nothing
    Set packngo1.swModel = Nothing
    Set packngo1.swEqMgr = Nothing
    Set packngo1.swPackAndGo = Nothing
    UIRerunCloseOutput app, sourcePath
    stagedPath = UIRerunCreateStage(app, sourceFolder, packngo1.GetFileName(sourcePath), stageFolder, stagedFiles)
    Set source = app.OpenDoc6(stagedPath, swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, _
                             configurationName, errors, warnings)
    If source Is Nothing Or errors <> 0 Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot open the local working assembly. SolidWorks error: " & CStr(errors)
    End If
    If warnings <> 0 Then UITraceDetail "Local copy open warnings", CStr(warnings)
    UIRerunCheckCopiedModel source, stageFolder
    UITraceEnd "Prepare local rerun copy"

    progress.NextStage
    stage = "applying"
    Set tableSession = New UIDesignTableSession
    tableSession.Apply source, designTableInputs, expected
    UIReadCalculatedGlobals source, expected
    UIVerifyDesignTableModel source, expected
    UIRerunBuildNames oldMaster, parentRoot
    finalFolder = parentRoot & packngo1.masterName & "\"
    finalPath = finalFolder & packngo1.masterName & ".SLDASM"
    If StrComp(finalFolder, sourceFolder, vbTextCompare) = 0 Or fso.FolderExists(finalFolder) Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "The new output folder already exists:" & vbCrLf & finalFolder
    End If
    UITraceDetail "New part number", packngo1.masterName

    progress.NextStage
    UITraceBegin "Prepare rerun Pack and Go"
    UITraceBegin "Local GetPackAndGo"
    Set pack = source.Extension.GetPackAndGo
    UITraceEnd "Local GetPackAndGo"
    If pack Is Nothing Then Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot get Pack and Go for the output assembly."
    UITraceBegin "Local Pack and Go drawing collection"
    pack.IncludeDrawings = True
    UITraceEnd "Local Pack and Go drawing collection"
    pack.IncludeSuppressed = True
    pack.IncludeSimulationResults = False
    pack.FlattenToSingleFolder = True
    UITraceBegin "Local Pack and Go GetDocumentNames"
    If pack.GetDocumentNames(documents) = False Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot list the output's Pack and Go documents."
    End If
    UITraceEnd "Local Pack and Go GetDocumentNames"
    If Not IsArray(documents) Then Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go returned no document list."
    Set includedPaths = CreateObject("Scripting.Dictionary")
    includedPaths.CompareMode = vbTextCompare
    For i = LBound(documents) To UBound(documents)
        includedPaths(CStr(documents(i))) = True
    Next i
    ' AddExternalDocuments copies extras outside the renamed native list; do not use it to add drawings.
    count = UIRerunListModelFiles(stageFolder, oldNames)
    For i = 0 To count - 1
        If Not includedPaths.Exists(stageFolder & oldNames(i)) Then
            Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go did not include an existing output document:" & vbCrLf & _
                stageFolder & oldNames(i) & vbCrLf & "Resolve its assembly/drawing reference before rerunning; it will not be silently omitted."
        End If
    Next i
    count = UBound(documents) - LBound(documents) + 1
    ReDim newNames(LBound(documents) To UBound(documents))
    Set usedNames = CreateObject("Scripting.Dictionary")
    usedNames.CompareMode = vbTextCompare
    plan = ""
    For i = LBound(documents) To UBound(documents)
        role = UIRerunTemplateFile(packngo1.GetFileName(CStr(documents(i))), oldMaster)
        If Len(role) = 0 Then
            newNames(i) = finalFolder & packngo1.GetFileName(CStr(documents(i)))
        Else
            newNames(i) = finalFolder & packngo1.GetNewBaseName(role)
        End If
        If usedNames.Exists(newNames(i)) Then
            Err.Raise RERUN_ERROR, "UIRunRerun", "Two planter files would both be named " & newNames(i) & ":" & vbCrLf & _
                usedNames(newNames(i)) & vbCrLf & CStr(documents(i))
        End If
        usedNames(newNames(i)) = CStr(documents(i))
        plan = plan & CStr(documents(i)) & " -> " & newNames(i) & "; "
    Next i
    If Not usedNames.Exists(finalPath) Then Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go did not include the master assembly."
    UITraceDetail "Pack and Go plan", plan
    UITraceEnd "Prepare rerun Pack and Go"

    progress.NextStage
    stage = "copying"
    fso.CreateFolder Left$(finalFolder, Len(finalFolder) - 1)
    If pack.SetSaveToName(True, finalFolder) = False Then Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot set the new output folder."
    If pack.SetDocumentSaveToNames(newNames) = False Then Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot set the new output file names."
    UITraceBegin "Save rerun Pack and Go"
    statuses = source.Extension.SavePackAndGo(pack)
    If Not IsArray(statuses) Then Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go returned no save statuses."
    If UBound(statuses) - LBound(statuses) + 1 <> count Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go returned an incomplete save-status list."
    End If
    For i = LBound(statuses) To UBound(statuses)
        If CLng(statuses(i)) <> 0 Then
            Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go failed for " & newNames(LBound(newNames) + i - LBound(statuses)) & _
                ". Save status: " & CStr(statuses(i))
        End If
    Next i
    For i = LBound(newNames) To UBound(newNames)
        If Not fso.FileExists(newNames(i)) Then Err.Raise RERUN_ERROR, "UIRunRerun", "Pack and Go did not create " & newNames(i)
    Next i
    UITraceEnd "Save rerun Pack and Go"

    progress.NextStage
    stage = "copied"
    Set tableSession = Nothing
    Set pack = Nothing
    Set source = Nothing
    Set packngo1.swModel = Nothing
    Set packngo1.swEqMgr = Nothing
    UIRerunCloseOutput app, stagedPath
    cleanupWarning = UIRerunCleanStage(stageFolder, stagedFiles)

    progress.NextStage
    Set packed = app.OpenDoc6(finalPath, swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, _
                              configurationName, errors, warnings)
    If packed Is Nothing Or errors <> 0 Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot open the new assembly. SolidWorks error: " & CStr(errors)
    End If
    If warnings <> 0 Then openWarnings = "SolidWorks reported open warnings (" & CStr(warnings) & "); review the assembly." & vbCrLf
    UITraceDetail "Opened copied assembly", finalPath & "; warnings=" & CStr(warnings)
    UIRerunCheckCopiedModel packed, finalFolder
    Set tableSession = New UIDesignTableSession
    tableSession.Apply packed, designTableInputs, copiedExpected
    UIReadCalculatedGlobals packed, copiedExpected
    UIVerifyDesignTableModel packed, expected
    If packed.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent Or swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, errors, warnings) = False Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot save the new assembly and references. SolidWorks error: " & CStr(errors)
    End If
    Set tableSession = Nothing
    Set packed = Nothing
    UIRerunCloseOutput app, finalPath
    Set packed = app.OpenDoc6(finalPath, swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, _
                              configurationName, errors, warnings)
    If packed Is Nothing Or errors <> 0 Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "Cannot reopen the saved new assembly. SolidWorks error: " & CStr(errors)
    End If
    If warnings <> 0 Then openWarnings = openWarnings & "SolidWorks reported reopen warnings (" & CStr(warnings) & "); review the assembly." & vbCrLf
    UIRerunCheckCopiedModel packed, finalFolder
    If packed.ForceRebuild3(False) = False Then
        Err.Raise RERUN_ERROR, "UIRunRerun", "The new assembly failed to rebuild after reopening." & UIRerunWhatsWrong(packed)
    End If
    UIVerifyDesignTableModel packed, expected
    stage = "verified"
    Set packngo1.swModel = packed
    Set packngo1.swEqMgr = packed.GetEquationMgr
    packngo1.destFolder = finalFolder
    UIRerunExports packed, progress
    progress.Complete
    UIEndRunLog "FINISHED", "New output: " & finalPath & "; " & CStr(count) & _
        " file(s) copied with custom edits kept. Original output kept. Review exporter status/logs for results." & _
        vbCrLf & openWarnings & cleanupWarning
    app.Frame.SetStatusBarText "New planter output created: " & finalPath & _
        ". Original output kept. Review exporter status/logs. Run log: " & logPath
    If Len(openWarnings & cleanupWarning) > 0 Then
        MsgBox openWarnings & cleanupWarning & vbCrLf & "New output: " & finalPath & vbCrLf & _
            "Run log: " & logPath, vbExclamation, "Rerun Warnings"
    End If
    Exit Sub

Failed:
    failureMessage = Err.Description
    If Not packed Is Nothing Then failureMessage = failureMessage & UIRerunWhatsWrong(packed)
    UITraceDetail "Workflow failure", failureMessage
    If Not progress Is Nothing Then progress.RestoringAfterFailure
    If Not tableSession Is Nothing Then
        If UICloseTableEditor(tableSession) = False Then
            failureMessage = failureMessage & vbCrLf & "Design table editor cleanup also failed."
        End If
    End If
    If Not progress Is Nothing Then progress.CloseWindow
    Select Case stage
        Case "checking"
            state = "No new model inputs were saved. Any approved saves of existing edits were kept."
        Case "applying"
            state = "Only the local working copy was updated. The original output was kept:" & vbCrLf & sourcePath & vbCrLf & _
                    "Local working copy: " & stagedPath
        Case "staging"
            state = "The original output was kept. Local working folder (possibly incomplete):" & vbCrLf & stageFolder
        Case "copying"
            state = "The original files were not renamed or updated with new inputs." & vbCrLf & _
                    "The new folder may be incomplete; do not use it:" & vbCrLf & finalFolder
        Case "copied"
            state = "The original output was kept. The new output has not passed verification; do not use it yet:" & vbCrLf & finalPath
        Case Else
            state = "The new output passed model verification; an export or later operation failed:" & vbCrLf & finalPath
    End Select
    UIEndRunLog "FAILED", failureMessage & vbCrLf & state
    MsgBox failureMessage & vbCrLf & vbCrLf & state & vbCrLf & "Run log: " & logPath, _
           vbCritical, "Rerun Stopped"
End Sub
