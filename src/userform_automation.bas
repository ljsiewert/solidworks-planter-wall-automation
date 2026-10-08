Attribute VB_Name = "userform_automation"
Option Explicit

' Import into the SAME VBA project as packngo1, not the exporter projects.
Public quoteOnly As Boolean
Public exportPDFAssemblies As Boolean
Public exportPDFComponents As Boolean
Public exportDXF As Boolean
Public exportSTEP As Boolean
Public outputFolder As String
Public designTableInputs As Variant

' These are referenced, but not declared, by the supplied packngo1 source.
Public filteredExcludedFiles As Boolean
Public excludedFiles As Variant

Private Const UI_ERROR As Long = vbObjectError + 2100
Private uiRunning As Boolean
Private runLogger As UIRunLogger

' Change this on other machines; logs stay local and are ignored by Git.
Public Const UI_LOG_FOLDER As String = _
    "C:\Users\lsiewert\StudioProjects\solidworks-planter-wall-automation\run-logs"

Public Function UIRunClock() As Double
    Dim day As Date
    Dim seconds As Double
    Do
        day = Date
        seconds = CDbl(Timer)
    Loop While day <> Date
    UIRunClock = CDbl(day) * 86400# + seconds
End Function

Public Sub UITraceBegin(ByVal name As String)
    If Not runLogger Is Nothing Then runLogger.BeginSpan name
End Sub

Public Sub UITraceEnd(ByVal name As String)
    If Not runLogger Is Nothing Then runLogger.EndSpan name
End Sub

Public Sub UITraceDetail(ByVal name As String, ByVal text As String)
    If Not runLogger Is Nothing Then runLogger.Detail name, text
End Sub

Public Sub RunWithUI()
    Dim form As UserForm_AutomationUI

    If uiRunning Then
        MsgBox "The UI workflow is already running.", vbExclamation, "Planter Automation"
        Exit Sub
    End If

    On Error GoTo Failed
    uiRunning = True
    If Not packngo1.Step1_Initialize() Then GoTo Finished
    UIRequireCleanTemplate packngo1.swModel
    If Not packngo1.Step2_ReadGlobals() Then GoTo Finished
    If StrComp(packngo1.GetFileName(packngo1.swModel.GetPathName), _
               "planter_assembly.SLDASM", vbTextCompare) <> 0 Then
        Err.Raise UI_ERROR, "RunWithUI", "Start UI mode from the original planter_assembly.SLDASM template."
    End If
    packngo1.destFolder = ""

    Set form = New UserForm_AutomationUI
    form.Show vbModal
    If form.Accepted Then
        Unload form
        Set form = Nothing
        UIRunAutomation
    End If

Finished:
    If Not form Is Nothing Then Unload form
    uiRunning = False
    Exit Sub

Failed:
    MsgBox Err.Description, vbCritical, "Planter Automation"
    Resume Finished
End Sub

Public Function UIParseInches(ByVal text As String) As Double
    Dim i As Long
    Dim character As String
    Dim decimalCount As Long
    Dim digitCount As Long
    Dim cleaned As String

    cleaned = Trim$(text)
    For i = 1 To Len(cleaned)
        character = Mid$(cleaned, i, 1)
        If character >= "0" And character <= "9" Then
            digitCount = digitCount + 1
        ElseIf character = "." Then
            decimalCount = decimalCount + 1
        Else
            Err.Raise UI_ERROR, "UIParseInches", _
                "Enter inches as a positive decimal, for example 100.25 (no units or fractions)."
        End If
    Next i
    If digitCount = 0 Or decimalCount > 1 Then
        Err.Raise UI_ERROR, "UIParseInches", "Enter a positive decimal measurement in inches."
    End If

    UIParseInches = Val(cleaned)
    If UIParseInches <= 0 Then
        Err.Raise UI_ERROR, "UIParseInches", "Measurements must be greater than 0 inches."
    End If
End Function

Public Function UIInchesText(ByVal value As Double) As String
    ' Str uses a period regardless of the Windows decimal separator.
    UIInchesText = Trim$(Str$(value))
End Function

Public Function UIProgressTime(ByVal seconds As Double) As String
    Dim minutes As Double
    minutes = Int(seconds / 60)
    UIProgressTime = CStr(minutes) & "m " & CStr(Int(seconds - minutes * 60)) & "s"
End Function

Private Function UIProgressPlan() As Collection
    Dim plan As New Collection
    plan.Add "Checking export macro files"
    plan.Add "Updating embedded design table and rebuilding template"
    plan.Add "Verifying calculated lengths and output destination"
    plan.Add "Collecting Pack-and-Go documents and rebuilding references"
    plan.Add "Creating output folder and applying part-number names"
    plan.Add "Saving Pack-and-Go files"
    plan.Add "Closing original template without saving and releasing protection"
    plan.Add "Opening packed assembly"
    plan.Add "Rebuilding, verifying and saving packed assembly"
    If exportDXF Then plan.Add "Running DXF component exporter"
    If exportSTEP Then plan.Add "Running STEP component exporter"
    If exportPDFAssemblies Then plan.Add "Running assembly PDF exporter"
    If exportPDFComponents Then plan.Add "Running component PDF exporter"
    Set UIProgressPlan = plan
End Function

Public Sub UIRequireCleanTemplate(ByVal model As SldWorks.ModelDoc2)
    If model.GetSaveFlag <> False Then
        Err.Raise UI_ERROR, "UIRequireCleanTemplate", _
            "The original template already has unsaved changes." & vbCrLf & _
            "Save or discard them manually, then reopen a clean template before running UI mode."
    End If
End Sub

Public Sub UIValidateDimensions(ByVal length As Double, ByVal width As Double, ByVal height As Double)
    If height < 6 Or height > 48 Then
        Err.Raise UI_ERROR, "UIValidateDimensions", "Height must be between 6 and 48 inches, inclusive."
    End If
    If length < 6 Or width < 6 Or length <= height Or width <= height Then
        Err.Raise UI_ERROR, "UIValidateDimensions", "Length and width must each be greater than height and at least 6 inches."
    End If
End Sub

Public Function UIWallTableLabel(ByVal walls As Long, ByVal orientation As String) As String
    Select Case walls
        Case 1: UIWallTableLabel = "ONE WALL"
        Case 2
            Select Case orientation
                Case "Left": UIWallTableLabel = "TWO WALLS (CENTER AND LEFT SIDE)"
                Case "Right": UIWallTableLabel = "TWO WALLS (CENTER AND RIGHT SIDE)"
                Case Else: Err.Raise UI_ERROR, "UIWallTableLabel", "Choose Left or Right."
            End Select
        Case 3: UIWallTableLabel = "THREE WALLS"
        Case 4: UIWallTableLabel = "FOUR WALLS"
        Case Else: Err.Raise UI_ERROR, "UIWallTableLabel", "Choose 1 to 4 walls."
    End Select
End Function

Public Function UIReturnTableLabel(ByVal label As String) As String
    Select Case label
        Case "No return": UIReturnTableLabel = "Default"
        Case "Single return 1 inch": UIReturnTableLabel = "single_return1"
        Case "Single return 2 inches": UIReturnTableLabel = "single_return2"
        Case "Double return": UIReturnTableLabel = "double_return"
        Case Else: Err.Raise UI_ERROR, "UIReturnTableLabel", "Choose a documented return type."
    End Select
End Function

Public Function UIReadTableNumber(ByVal value As Variant, ByVal address As String) As Double
    Dim text As String
    If IsError(value) Then
        Err.Raise UI_ERROR, "UIReadTableNumber", "Excel formula error at " & address
    End If
    text = Trim$(CStr(value))
    If Left$(text, 1) = "=" Then text = Mid$(text, 2)
    If Len(text) = 0 Then
        Err.Raise UI_ERROR, "UIReadTableNumber", "Empty calculated value at " & address
    End If
    ' Table formulas produce strings such as "=113"; these are not equations to write.
    If text = "0" Then
        UIReadTableNumber = 0
    ElseIf VarType(value) <> vbString And IsNumeric(value) Then
        UIReadTableNumber = CDbl(value)
    Else
        UIReadTableNumber = UIParseInches(text)
    End If
End Function

Public Function UIOutputRoot(ByVal selectedFolder As String) As String
    Dim fso As Object
    Dim root As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    root = Trim$(selectedFolder)
    If Len(root) = 0 Then
        Err.Raise UI_ERROR, "UIOutputRoot", "Choose an existing output folder."
    End If
    If Not fso.FolderExists(root) Then
        Err.Raise UI_ERROR, "UIOutputRoot", "Output folder does not exist: " & root
    End If
    root = fso.GetAbsolutePathName(root)
    If Right$(root, 1) <> "\" Then root = root & "\"
    If InStr(1, "\" & root, "\PACK_TEST\", vbTextCompare) = 0 Then
        root = root & "PACK_TEST\"
    End If
    UIOutputRoot = root
End Function

Public Sub UIValidateSelections(ByVal walls As Long, ByVal orientation As String, _
                                ByVal returnLabel As String, ByVal materialLabel As String, _
                                ByVal gaugeLabel As String)
    If walls < 1 Or walls > 4 Then
        Err.Raise UI_ERROR, "UIValidateSelections", "The naming specification supports 1 to 4 walls."
    End If
    If walls = 2 Then
        If orientation <> "Left" And orientation <> "Right" Then
            Err.Raise UI_ERROR, "UIValidateSelections", "Choose Left or Right for a two-wall model."
        End If
    ElseIf orientation <> "N/A" Then
        Err.Raise UI_ERROR, "UIValidateSelections", "Side orientation is N/A unless wall count is 2."
    End If
    Select Case returnLabel
        Case "No return", "Single return 1 inch", "Single return 2 inches", "Double return"
        Case Else
            Err.Raise UI_ERROR, "UIValidateSelections", "Choose a documented return type."
    End Select
    Select Case materialLabel
        Case "Mild Steel", "Borcon Weathering Steel"
        Case Else
            Err.Raise UI_ERROR, "UIValidateSelections", "Choose a documented material."
    End Select
    Select Case gaugeLabel
        Case "3/16 inch", "1/4 inch"
        Case Else
            Err.Raise UI_ERROR, "UIValidateSelections", "Choose a documented thickness."
    End Select
End Sub

Public Sub UIMapSelections(ByVal walls As Long, ByVal orientation As String, _
                          ByVal returnLabel As String, ByVal materialLabel As String, _
                          ByVal gaugeLabel As String)
    UIValidateSelections walls, orientation, returnLabel, materialLabel, gaugeLabel
    packngo1.wallCount = CStr(walls)
    Select Case orientation
        Case "N/A": packngo1.side = "0"
        Case "Right": packngo1.side = "1"
        Case "Left": packngo1.side = "2"
    End Select
    Select Case returnLabel
        Case "No return": packngo1.returnType = "0"
        Case "Single return 1 inch": packngo1.returnType = "1"
        Case "Single return 2 inches": packngo1.returnType = "2"
        Case "Double return": packngo1.returnType = "3"
    End Select
    Select Case materialLabel
        Case "Mild Steel": packngo1.overallMaterial = "1"
        Case "Borcon Weathering Steel": packngo1.overallMaterial = "2"
    End Select
    Select Case gaugeLabel
        Case "3/16 inch": packngo1.thicknessLabel = "1"
        Case "1/4 inch": packngo1.thicknessLabel = "2"
    End Select
End Sub

Public Sub UICheckDestination()
    Dim fso As Object
    Dim candidate As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not packngo1.Step3_BuildNames() Then
        Err.Raise UI_ERROR, "UICheckDestination", "Could not build the output part numbers."
    End If
    candidate = UIOutputRoot(outputFolder) & packngo1.masterName & "\"
    If fso.FolderExists(candidate) Then
        Err.Raise UI_ERROR, "UICheckDestination", _
            "This output already exists. Choose another parent folder; UI mode never overwrites:" & _
            vbCrLf & candidate
    End If
End Sub

Public Sub UICheckExportFiles()
    If exportDXF Then UIRequireMacro "dxfcomponents.swp"
    If exportSTEP Then UIRequireMacro "stepcomponents.swp"
    If exportPDFAssemblies Then UIRequireMacro "pdfassemblies.swp"
    If exportPDFComponents Then UIRequireMacro "pdfcomponents.swp"
End Sub

Private Sub UIRequireMacro(ByVal name As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(packngo1.MACRO_FOLDER & name) Then
        Err.Raise UI_ERROR, "UIRequireMacro", _
            "Selected export macro is missing: " & packngo1.MACRO_FOLDER & name
    End If
End Sub

Public Sub UIVerifyDesignTableModel(ByVal model As SldWorks.ModelDoc2, ByVal expected As Variant)
    Dim names As Variant
    Dim manager As SldWorks.EquationMgr
    Dim units As Variant
    Dim i As Long
    Dim actual As String
    names = Array("WALL_COUNT", "SIDE", "RETURN_TYPE", "OVERALL_LENGTH", _
                  "OVERALL_WIDTH", "OVERALL_HEIGHT", "CENTER_PLANTER_LENGTH", _
                  "SIDE_PLANTER_LENGTH", "OVERALL_MATERIAL", "OVERALL_THICKNESS")
    units = model.GetUnits
    If units(0) <> swLengthUnit_e.swINCHES Then
        Err.Raise UI_ERROR, "UIVerifyModel", "The supplied design table requires document length units of inches."
    End If
    Set manager = model.GetEquationMgr
    For i = LBound(names) To UBound(names)
        actual = packngo1.GetGlobalVariable(manager, CStr(names(i)))
        If actual = "ERROR" Then
            Err.Raise UI_ERROR, "UIVerifyModel", "Cannot read model global: " & CStr(names(i))
        End If
        If Abs(CDbl(actual) - CDbl(expected(i))) > 0.000001 * (1 + Abs(CDbl(expected(i)))) Then
            Err.Raise UI_ERROR, "UIVerifyModel", _
                "Design table did not update " & CStr(names(i)) & "." & vbCrLf & _
                "Expected: " & CStr(expected(i)) & "; model: " & actual
        End If
    Next i
End Sub

Private Sub UIReadCalculatedGlobals(ByVal model As SldWorks.ModelDoc2, ByVal expected As Variant)
    Dim requested As Variant
    Dim i As Long
    requested = Array(packngo1.wallCount, packngo1.side, packngo1.returnType, _
        packngo1.overallLength, packngo1.overallWidth, packngo1.overallHeight)
    For i = 0 To 5
        If Abs(CDbl(expected(i)) - CDbl(requested(i))) > 0.000001 Then
            Err.Raise UI_ERROR, "UIReadCalculatedGlobals", "The table calculation does not match the submitted input at index " & CStr(i)
        End If
    Next i
    If CDbl(expected(8)) <> CDbl(packngo1.overallMaterial) Then
        Err.Raise UI_ERROR, "UIReadCalculatedGlobals", "The table material calculation does not match the selected material."
    End If
    If CDbl(expected(9)) <> IIf(packngo1.thicknessLabel = "1", 0.1875, 0.25) Then
        Err.Raise UI_ERROR, "UIReadCalculatedGlobals", "The table thickness calculation does not match the selected thickness."
    End If
    UIVerifyDesignTableModel model, expected
    Set packngo1.swEqMgr = model.GetEquationMgr
    ' R8 is the naming code; K3 drives physical thickness, not THICKNESS_LABEL.
    packngo1.centerPlanterLength = CStr(expected(6))
    packngo1.sidePlanterLength = CStr(expected(7))
    packngo1.thicknessLabel = CStr(IIf(CDbl(expected(9)) = 0.1875, 1, 2))
End Sub

Private Sub UICreateDestination()
    Dim fso As Object
    Dim root As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    root = UIOutputRoot(outputFolder)
    If Not fso.FolderExists(root) Then fso.CreateFolder root
    packngo1.packTestFolder = root
    packngo1.destFolder = root & packngo1.masterName & "\"
    If fso.FolderExists(packngo1.destFolder) Then
        Err.Raise UI_ERROR, "UICreateDestination", "Output folder was created by another run; refusing to overwrite it."
    End If
    fso.CreateFolder packngo1.destFolder
    If packngo1.swPackAndGo.SetSaveToName(True, packngo1.destFolder) = False Then
        Err.Raise UI_ERROR, "UICreateDestination", "SolidWorks rejected the output folder."
    End If
    packngo1.stepFolder = packngo1.destFolder & "STEP\"
    packngo1.pdfAssembliesFolder = packngo1.destFolder & "PDF_Assemblies\"
    packngo1.pdfComponentsFolder = packngo1.destFolder & "PDF_Components\"
End Sub

Private Sub UICheckSaveStatuses(ByVal statuses As Variant)
    Dim i As Long
    If Not IsArray(statuses) Then
        Err.Raise UI_ERROR, "UICheckSaveStatuses", "SolidWorks returned no Pack-and-Go save status array."
    End If
    If UBound(statuses) - LBound(statuses) + 1 <> packngo1.docCount Then
        Err.Raise UI_ERROR, "UICheckSaveStatuses", "SolidWorks returned an incomplete Pack-and-Go status array."
    End If
    For i = LBound(statuses) To UBound(statuses)
        If CLng(statuses(i)) <> 0 Then
            Err.Raise UI_ERROR, "UICheckSaveStatuses", _
                "Pack-and-Go failed at document " & CStr(i + 1) & _
                "; save status " & CStr(statuses(i)) & "."
        End If
    Next i
End Sub

Private Sub UIRunExport(ByVal packed As SldWorks.ModelDoc2, _
                       ByVal file As String, ByVal procedure As String)
    Dim activated As SldWorks.ModelDoc2
    Dim errors As Long

    Set activated = packngo1.swApp.ActivateDoc3(packed.GetPathName, False, _
                        swRebuildOnActivation_e.swDontRebuildActiveDoc, errors)
    If activated Is Nothing Or errors <> 0 Then
        Err.Raise UI_ERROR, "UIRunExport", "Cannot activate the packed assembly for " & procedure
    End If
    If packngo1.RunExternalMacro(packngo1.MACRO_FOLDER & file, procedure) = False Then
        Err.Raise UI_ERROR, "UIRunExport", "Export macro invocation failed: " & procedure
    End If
End Sub

Private Sub UICollectPackAndGoDocuments()
    packngo1.usedManualFallback = False
    Set packngo1.swPackAndGo = Nothing
    ' Mirrors the manual Step4 sequence, with separate timing and API checks.
    UITraceBegin "Collection ForceRebuild3"
    If packngo1.swModel.ForceRebuild3(False) = False Then
        Err.Raise UI_ERROR, "UICollectPackAndGoDocuments", "The template failed to rebuild before document collection."
    End If
    UITraceEnd "Collection ForceRebuild3"
    UITraceBegin "GetPackAndGo"
    Set packngo1.swPackAndGo = packngo1.swModel.Extension.GetPackAndGo
    If packngo1.swPackAndGo Is Nothing Then
        Err.Raise UI_ERROR, "UICollectPackAndGoDocuments", "Could not obtain the native Pack-and-Go object."
    End If
    UITraceEnd "GetPackAndGo"
    UITraceBegin "Pack-and-Go IncludeDrawings"
    packngo1.swPackAndGo.IncludeDrawings = True
    UITraceEnd "Pack-and-Go IncludeDrawings"
    packngo1.swPackAndGo.IncludeSimulationResults = False
    packngo1.swPackAndGo.FlattenToSingleFolder = True
    UITraceBegin "Pack-and-Go GetDocumentNames"
    If packngo1.swPackAndGo.GetDocumentNames(packngo1.docNames) = False Then
        Err.Raise UI_ERROR, "UICollectPackAndGoDocuments", _
            "Native Pack-and-Go document collection failed. UI mode cannot safely rename the manual fallback list."
    End If
    packngo1.docCount = UBound(packngo1.docNames) + 1
    UITraceEnd "Pack-and-Go GetDocumentNames"
    UITraceDetail "Pack-and-Go options", "Drawings=True; SimulationResults=False; FlattenToSingleFolder=True"
End Sub

Private Sub UITraceSourceFiles()
    Dim fso As Object
    Dim i As Long
    Dim bytes As Double
    Dim failed As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    UITraceBegin "Source metadata inventory"
    For i = 0 To UBound(packngo1.docNames)
        If UITraceSourceFile(fso, CStr(packngo1.docNames(i)), bytes) = False Then failed = failed + 1
    Next i
    UITraceEnd "Source metadata inventory"
    UITraceDetail "Source inventory totals", "Documents=" & CStr(packngo1.docCount) & _
        "; Bytes=" & UIInchesText(bytes) & "; Metadata failures=" & CStr(failed)
    If failed > 0 Then
        MsgBox CStr(failed) & " source-file metadata lookups failed; see the run log." & vbCrLf & _
            "This diagnostic does not determine whether SolidWorks can pack these documents.", _
            vbExclamation, "Source Access Diagnostics"
    End If
End Sub

Private Function UITraceSourceFile(ByVal fso As Object, ByVal sourcePath As String, _
                                  ByRef totalBytes As Double) As Boolean
    Dim file As Object
    Dim started As Double
    Dim bytes As Double
    Dim failure As String
    started = UIRunClock()
    On Error GoTo Failed
    Set file = fso.GetFile(sourcePath)
    bytes = CDbl(file.Size)
    totalBytes = totalBytes + bytes
    UITraceDetail "Source file", sourcePath & "; Bytes=" & UIInchesText(bytes) & _
        "; MetadataSeconds=" & UIInchesText(UIRunClock() - started)
    UITraceSourceFile = True
    Exit Function
Failed:
    failure = Err.Description
    UITraceDetail "Source metadata failure", sourcePath & "; Seconds=" & _
        UIInchesText(UIRunClock() - started) & "; Error=" & failure
End Function

Private Sub UIRunAutomation()
    Dim template As SldWorks.ModelDoc2
    Dim tableSession As UIDesignTableSession
    Dim packed As SldWorks.ModelDoc2
    Dim expected As Variant
    Dim statuses As Variant
    Dim errors As Long
    Dim warnings As Long
    Dim packedPath As String
    Dim configurationName As String
    Dim failureMessage As String
    Dim fso As Object
    Dim progress As UIProgressSession
    Dim saveGuard As UIReferenceSaveGuard
    Dim templatePath As String
    Dim stillOpen As SldWorks.ModelDoc2
    Dim logPath As String
    Dim inputIndex As Long

    On Error GoTo Failed
    Set runLogger = New UIRunLogger
    runLogger.BeginRun UI_LOG_FOLDER
    logPath = runLogger.Path
    Set progress = New UIProgressSession
    progress.BeginRun UIProgressPlan()
    Set template = packngo1.swModel
    UIRequireCleanTemplate template
    templatePath = template.GetPathName
    UITraceDetail "Template", templatePath
    UITraceDetail "SolidWorks revision", packngo1.swApp.RevisionNumber
    UITraceDetail "Output parent", outputFolder
    For inputIndex = 0 To 6
        UITraceDetail "Input N" & CStr(inputIndex + 3), CStr(designTableInputs(inputIndex))
    Next inputIndex
    UITraceDetail "Exports", "Quote=" & CStr(quoteOnly) & "; DXF=" & CStr(exportDXF) & _
        "; STEP=" & CStr(exportSTEP) & "; PDF assemblies=" & CStr(exportPDFAssemblies) & _
        "; PDF components=" & CStr(exportPDFComponents)
    configurationName = template.ConfigurationManager.ActiveConfiguration.Name
    UITraceDetail "Configuration", configurationName
    Set fso = CreateObject("Scripting.FileSystemObject")
    progress.NextStage
    UICheckExportFiles
    progress.NextStage
    Set tableSession = New UIDesignTableSession
    Set saveGuard = New UIReferenceSaveGuard
    UITraceBegin "Reference protection preflight"
    saveGuard.BeginGuard packngo1.swApp, template
    UITraceEnd "Reference protection preflight"
    tableSession.SetSaveGuard saveGuard
    tableSession.Apply template, designTableInputs, expected
    progress.NextStage
    UIReadCalculatedGlobals template, expected
    UICheckDestination
    UITraceDetail "Calculated planter lengths", "Center=" & packngo1.centerPlanterLength & _
        "; side=" & packngo1.sidePlanterLength

    progress.NextStage
    UICollectPackAndGoDocuments
    UITraceDetail "Pack-and-Go document count", CStr(packngo1.docCount)
    UITraceSourceFiles
    UIVerifyDesignTableModel template, expected
    progress.NextStage
    UICreateDestination
    If Not packngo1.Step6_BuildRenamedList() Then
        Err.Raise UI_ERROR, "UIRunAutomation", "Could not build the renamed document list."
    End If
    If Not packngo1.Step7_ApplyRenamedList() Then
        Err.Raise UI_ERROR, "UIRunAutomation", "Could not apply the renamed document list."
    End If
    progress.NextStage
    UITraceBegin "SavePackAndGo API"
    statuses = template.Extension.SavePackAndGo(packngo1.swPackAndGo)
    UITraceEnd "SavePackAndGo API"
    UICheckSaveStatuses statuses
    UITraceDetail "Pack-and-Go statuses", "All returned document statuses passed validation."
    UITraceDetail "Pack-and-Go destination", packngo1.destFolder
    UITraceDetail "Template inputs", "No original design-table input restoration will be performed."
    packedPath = packngo1.destFolder & packngo1.masterName & ".SLDASM"
    If Not fso.FileExists(packedPath) Then
        Err.Raise UI_ERROR, "UIRunAutomation", _
            "Expected main assembly was not created. Use the original planter_assembly.SLDASM template."
    End If

    progress.NextStage
    saveGuard.ProtectLoadedReferences
    Set tableSession = Nothing
    Set template = Nothing
    Set packngo1.swEqMgr = Nothing
    Set packngo1.swPackAndGo = Nothing
    Set packngo1.swModel = Nothing
    packngo1.swApp.CloseDoc templatePath
    Set stillOpen = packngo1.swApp.GetOpenDocumentByName(templatePath)
    If Not stillOpen Is Nothing Then
        Err.Raise UI_ERROR, "UIRunAutomation", _
            "The original template remains loaded after CloseDoc. Check whether another document references it."
    End If
    saveGuard.Restore
    progress.NextStage
    Set packed = packngo1.swApp.OpenDoc6(packedPath, swDocASSEMBLY, _
                     swOpenDocOptions_e.swOpenDocOptions_Silent, configurationName, errors, warnings)
    If packed Is Nothing Or errors <> 0 Then
        Err.Raise UI_ERROR, "UIRunAutomation", "Cannot open the packed assembly. SolidWorks error: " & CStr(errors)
    End If
    If warnings <> 0 Then
        Err.Raise UI_ERROR, "UIRunAutomation", "Packed assembly opened with warnings: " & CStr(warnings) & ". Review it before exporting."
    End If
    UITraceDetail "Opened packed assembly", packed.GetPathName
    progress.NextStage
    UITraceBegin "Packed assembly ForceRebuild3"
    If packed.ForceRebuild3(False) = False Then
        Err.Raise UI_ERROR, "UIRunAutomation", "The packed assembly failed to rebuild."
    End If
    UITraceEnd "Packed assembly ForceRebuild3"
    UIVerifyDesignTableModel packed, expected
    UITraceBegin "Packed assembly Save3"
    If packed.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, errors, warnings) = False Then
        Err.Raise UI_ERROR, "UIRunAutomation", "Cannot save the rebuilt packed assembly. SolidWorks error: " & CStr(errors)
    End If
    UITraceEnd "Packed assembly Save3"
    If errors <> 0 Or warnings <> 0 Then
        Err.Raise UI_ERROR, "UIRunAutomation", "Packed assembly save reported errors/warnings: " & CStr(errors) & "/" & CStr(warnings)
    End If
    Set packngo1.swModel = packed
    Set packngo1.swEqMgr = packed.GetEquationMgr
    If exportDXF Then
        progress.NextStage
        UIRunExport packed, "dxfcomponents.swp", "RunDXFComponentDrawings"
    End If
    If exportSTEP Then
        progress.NextStage
        UIRunExport packed, "stepcomponents.swp", "RunSTEPComponentExport"
    End If
    If exportPDFAssemblies Then
        progress.NextStage
        UIRunExport packed, "pdfassemblies.swp", "RunPDFAssemblyDrawings"
    End If
    If exportPDFComponents Then
        progress.NextStage
        UIRunExport packed, "pdfcomponents.swp", "RunPDFComponentDrawings"
    End If
    progress.Complete
    runLogger.Finish "FINISHED", "Workflow returned; export success must be checked in exporter dialogs/logs."
    Set runLogger = Nothing

    MsgBox "Pack-and-Go saved to:" & vbCrLf & packngo1.destFolder & vbCrLf & vbCrLf & _
           "Selected export macros have returned. Review their dialogs and logs for export results." & _
           vbCrLf & "The original template was closed without saving." & _
           vbCrLf & "Run log: " & logPath, _
           vbInformation, "UI Workflow Finished"
    Exit Sub

Failed:
    failureMessage = Err.Description
    UITraceDetail "Workflow failure", failureMessage
    UITraceBegin "Failure cleanup"
    If Not progress Is Nothing Then progress.RestoringAfterFailure
    If Not tableSession Is Nothing Then
        If UICloseTableEditor(tableSession) = False Then
            failureMessage = failureMessage & vbCrLf & "Design table editor cleanup also failed."
        End If
    End If
    If Not saveGuard Is Nothing Then
        If UIRestoreSaveGuard(saveGuard) = False Then
            failureMessage = failureMessage & vbCrLf & "Reference-save protection restoration also failed."
        End If
    End If
    If Not progress Is Nothing Then progress.CloseWindow
    If Not runLogger Is Nothing Then
        UITraceEnd "Failure cleanup"
        runLogger.Finish "FAILED", failureMessage
    End If
    Set runLogger = Nothing
    MsgBox failureMessage & vbCrLf & vbCrLf & _
           "Any partial output has been kept for inspection. No existing output was deleted." & _
           vbCrLf & "No template restoration was attempted. If the original is still open, close it without saving." & _
           vbCrLf & "Output: " & packngo1.destFolder & vbCrLf & "Run log: " & logPath, _
           vbCritical, "UI Workflow Stopped"
End Sub

Private Function UIRestoreSaveGuard(ByVal guard As UIReferenceSaveGuard) As Boolean
    On Error GoTo Failed
    guard.Restore
    UIRestoreSaveGuard = True
    Exit Function
Failed:
    MsgBox Err.Description & vbCrLf & _
           "Review document read-only states and System Options > External References.", _
           vbCritical, "Reference Save Settings"
End Function

Private Function UICloseTableEditor(ByVal session As UIDesignTableSession) As Boolean
    On Error GoTo Failed
    session.CloseEditor
    UICloseTableEditor = True
    Exit Function
Failed:
    MsgBox "Could not close the design table editor. Close it manually and do not save the template." & vbCrLf & _
           Err.Description, vbCritical, "Design Table Cleanup"
End Function
