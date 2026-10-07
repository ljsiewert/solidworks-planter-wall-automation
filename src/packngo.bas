Option Explicit

'============================================================
' >>> DEPLOY TARGET <<<
' File:   Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\packngo.swp
' Module: packngo1
' Action: Tools > Macro > Edit on packngo.swp, select all code in module
'         packngo1, delete, paste this file's content in, save.
' Changed: Step8_ConfirmAndSave now rebuilds the newly opened packed
'         assembly and calls the new PromptAndRunExports sub (asks
'         quote-only / DXF / STEP / PDF, then runs the right macros).
' Changed 2: Pre-created PDF output folders are now split into
'         PDF_Assemblies\ and PDF_Components\ (previously both shared
'         one PDF\ folder) to match the same split in pdfassemblies.bas
'         and pdfcomponents.bas.
' Changed 3: RunExternalMacro now also treats a 0 error code
'         (swRunMacroDone) as success even when RunMacro2's boolean
'         return value is False - that return value is known to be
'         unreliable and was causing "Failed to run macro" popups for
'         macros that actually completed successfully.
'============================================================

'============================================================
' PACK AND GO - MAIN WORKFLOW
'
' Steps 1-8: Initialize, read globals, build names, set
' destination, and run Pack and Go
'
' RUN:
'   RunPackAndGo       - Execute full workflow
'
' Individual steps can be run independently for testing.
'============================================================

'============================================================
' SHARED STATE
' (Public so utilities can access them)
'============================================================

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2
Public swEqMgr As SldWorks.EquationMgr
Public swPackAndGo As SldWorks.PackAndGo

Public wallCount As String
Public side As String
Public returnType As String

Public overallLength As String
Public overallWidth As String
Public overallHeight As String

Public centerPlanterLength As String
Public sidePlanterLength As String

Public thicknessLabel As String
Public overallMaterial As String

Public masterName As String
Public centerName As String
Public centerHardwareName As String
Public centerGussetName As String
Public mountingGussetName As String
Public sideName As String
Public sideLeftName As String
Public sideRightName As String
Public sideCenterName As String
Public sideWallName As String
Public sharedSideName As String

Public packTestFolder As String
Public destFolder As String
Public stepFolder As String
Public pdfAssembliesFolder As String
Public pdfComponentsFolder As String

Public docNames As Variant
Public docCount As Long
Public usedManualFallback As Boolean

Public renamedSummary As String

'============================================================
' EXTERNAL EXPORT MACRO CONFIG
'
' These macros live on the network drive as separate standalone
' .swp files (not modules in this project), so they are invoked
' via swApp.RunMacro2. Verified against the live files on the
' network drive:
'   packngo.swp, pdfassemblies.swp, pdfcomponents.swp
' Each .swp's VBA module is auto-named "<basefilename>1" (e.g.
' pdfassemblies.swp -> module "pdfassemblies1"), which is how
' SolidWorks names the default module when a macro is created.
' RunExternalMacro derives the module name from the filename
' using that convention - update DeriveModuleName if a macro
' file's module was ever renamed manually.
'============================================================

Public Const MACRO_FOLDER As String = _
    "Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\"



'============================================================
' MAIN ENTRY POINT
'============================================================

Sub RunPackAndGo()

    If Not Step1_Initialize() Then
        Exit Sub
    End If
    If Not Step2_ReadGlobals() Then
        Exit Sub
    End If
    If Not Step3_BuildNames() Then
        Exit Sub
    End If
    If Not Step4_GetDocumentNames() Then
        Exit Sub
    End If
    If Not Step5_SetDestination() Then
        Exit Sub
    End If
    If Not Step6_BuildRenamedList() Then
        Exit Sub
    End If
    If Not Step7_ApplyRenamedList() Then
        Exit Sub
    End If

    Step8_ConfirmAndSave

End Sub


'============================================================
' STEP 1: GET THE ACTIVE DOCUMENT AND VALIDATE IT
'============================================================

Function Step1_Initialize() As Boolean

    Step1_Initialize = False

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No SOLIDWORKS document is open.", _
               vbExclamation, "Step 1"
        Exit Function
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "The active document must be an assembly.", _
               vbExclamation, "Step 1"
        Exit Function
    End If

    If swModel.GetPathName = "" Then
        MsgBox "The assembly has not been saved yet.", _
               vbExclamation, "Step 1"
        Exit Function
    End If

    packTestFolder = _
        "Z:\Engineering\Levi Stuff\Border Concepts\Planter_Walls\PACK_TEST\"

    Step1_Initialize = True

End Function


'============================================================
' STEP 2: READ GLOBAL VARIABLES FROM EQUATION MANAGER
'============================================================

Function Step2_ReadGlobals() As Boolean

    Step2_ReadGlobals = False

    Set swEqMgr = swModel.GetEquationMgr

    wallCount = GetGlobalVariable(swEqMgr, "WALL_COUNT")
    side = GetGlobalVariable(swEqMgr, "SIDE")
    returnType = GetGlobalVariable(swEqMgr, "RETURN_TYPE")

    overallLength = GetGlobalVariable(swEqMgr, "OVERALL_LENGTH")
    overallWidth = GetGlobalVariable(swEqMgr, "OVERALL_WIDTH")
    overallHeight = GetGlobalVariable(swEqMgr, "OVERALL_HEIGHT")

    centerPlanterLength = _
        GetGlobalVariable(swEqMgr, "CENTER_PLANTER_LENGTH")
    sidePlanterLength = _
        GetGlobalVariable(swEqMgr, "SIDE_PLANTER_LENGTH")

    thicknessLabel = GetGlobalVariable(swEqMgr, "THICKNESS_LABEL")
    overallMaterial = GetGlobalVariable(swEqMgr, "OVERALL_MATERIAL")

    If wallCount = "ERROR" Or side = "ERROR" Or returnType = "ERROR" Then
        Exit Function
    End If

    If overallLength = "ERROR" Or overallWidth = "ERROR" Or overallHeight = "ERROR" Then
        Exit Function
    End If

    If centerPlanterLength = "ERROR" Or sidePlanterLength = "ERROR" Then
        Exit Function
    End If

    If thicknessLabel = "ERROR" Or overallMaterial = "ERROR" Then
        Exit Function
    End If

    Step2_ReadGlobals = True

End Function


'============================================================
' STEP 3: BUILD PART NUMBERS AND COMPONENT NAMES
'============================================================

Function Step3_BuildNames() As Boolean

    Dim pnOverallLength As String
    Dim pnOverallWidth As String
    Dim pnCenterLength As String
    Dim pnSideLength As String
    Dim pnHeight As String

    Step3_BuildNames = False

    pnOverallLength = FormatPNDimension(CDbl(overallLength))
    pnOverallWidth = FormatPNDimension(CDbl(overallWidth))
    pnCenterLength = FormatPNDimension(CDbl(centerPlanterLength))
    pnSideLength = FormatPNDimension(CDbl(sidePlanterLength))
    pnHeight = FormatPNDimension(CDbl(overallHeight))

    masterName = _
        wallCount & side & _
        "-" & returnType & _
        "-C" & pnOverallLength & _
        "-S" & pnOverallWidth & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    centerName = _
        returnType & _
        "-C" & pnCenterLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    centerHardwareName = _
        "C" & pnCenterLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    centerGussetName = _
        pnHeight & _
        "-" & thicknessLabel & overallMaterial & _
        "-CG"

    mountingGussetName = _
        pnHeight & _
        "-" & thicknessLabel & overallMaterial & _
        "-MG"

    sideName = _
        returnType & _
        "-S" & pnSideLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    sideLeftName = _
        returnType & _
        "-SL" & pnSideLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    sideRightName = _
        returnType & _
        "-SR" & pnSideLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    sideCenterName = _
        returnType & _
        "-SC" & pnSideLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    sideWallName = _
        pnHeight & _
        "-" & thicknessLabel & overallMaterial

    sharedSideName = _
        "S" & pnSideLength & _
        "-" & pnHeight & _
        "-" & thicknessLabel & overallMaterial

    Step3_BuildNames = True

End Function


'============================================================
' STEP 4: GET PACK AND GO DOCUMENT LIST
'
' Tries the native PackAndGo.GetDocumentNames call first.
' If it fails (common with suppressed components), falls back
' to manually walking the assembly tree and building the list.
'============================================================

Function Step4_GetDocumentNames() As Boolean

    Step4_GetDocumentNames = False
    usedManualFallback = False

    If swModel Is Nothing Then
        MsgBox "Run Step 1 first (swModel is not set).", vbExclamation, "Step 4"
        Exit Function
    End If

    ' Force rebuild
    swModel.ForceRebuild3 False

    If swPackAndGo Is Nothing Then
        Set swPackAndGo = swModel.Extension.GetPackAndGo
        If swPackAndGo Is Nothing Then
            MsgBox "Could not get Pack and Go object.", vbCritical, "Step 4"
            Exit Function
        End If
    End If

    swPackAndGo.IncludeDrawings = True
    swPackAndGo.IncludeSimulationResults = False
    swPackAndGo.FlattenToSingleFolder = True

    ' Try native call first
    If swPackAndGo.GetDocumentNames(docNames) Then
        docCount = UBound(docNames) + 1
        usedManualFallback = False
    Else
        ' Fallback to manual walk
        usedManualFallback = True

        If Not GetPackAndGoDocuments_Manual() Then
            MsgBox "Pack and Go native call failed, and manual fallback found no components." & vbCrLf & vbCrLf & _
                   "Verify all components are loaded (not all suppressed).", _
                   vbCritical, "Step 4"
            Exit Function
        End If

        docCount = UBound(docNames) + 1
    End If

    Step4_GetDocumentNames = True

End Function



'============================================================
' STEP 5: SET DESTINATION FOLDER
'============================================================

Function Step5_SetDestination() As Boolean

    Dim result As Boolean
    Dim folderExists As Boolean
    Dim userChoice As VbMsgBoxResult

    Step5_SetDestination = False

    destFolder = packTestFolder & masterName & "\"

    ' Check if folder already exists
    folderExists = (Dir(destFolder, vbDirectory) <> "")
    
    If folderExists Then
        ' Prompt user to overwrite or cancel
        userChoice = MsgBox("Destination folder already exists:" & vbCrLf & vbCrLf & _
                           destFolder & vbCrLf & vbCrLf & _
                           "Would you like to overwrite it or cancel?", _
                           vbQuestion + vbYesNo, _
                           "Folder Exists")
        
        If userChoice = vbNo Then
            MsgBox "Pack and Go cancelled.", vbInformation, "Cancelled"
            Step5_SetDestination = False
            Exit Function
        End If
        
        ' User chose to overwrite - clear the folder contents
        On Error Resume Next
        ClearFolderContents destFolder
        On Error GoTo 0
    Else
        On Error GoTo FolderError

        MkDir destFolder

        On Error GoTo 0
    End If

    result = swPackAndGo.SetSaveToName(True, destFolder)

    stepFolder = destFolder & "STEP\"
    pdfAssembliesFolder = destFolder & "PDF_Assemblies\"
    pdfComponentsFolder = destFolder & "PDF_Components\"

    If Dir(stepFolder, vbDirectory) = "" Then

        On Error GoTo FolderError

        MkDir stepFolder

        On Error GoTo 0

    End If

    If Dir(pdfAssembliesFolder, vbDirectory) = "" Then

        On Error GoTo FolderError

        MkDir pdfAssembliesFolder

        On Error GoTo 0

    End If

    If Dir(pdfComponentsFolder, vbDirectory) = "" Then

        On Error GoTo FolderError

        MkDir pdfComponentsFolder

        On Error GoTo 0

    End If

    Step5_SetDestination = True

    Exit Function

FolderError:
    MsgBox "Could not create destination folder:" & vbCrLf & destFolder, _
           vbCritical, "Step 5"

End Function


'============================================================
' STEP 6: BUILD RENAMED DOCUMENT LIST
'============================================================

Function Step6_BuildRenamedList() As Boolean

    Dim i As Long
    Dim currentFile As String
    Dim newName As String
    Dim skippedCount As Long

    Step6_BuildRenamedList = False

    renamedSummary = "RENAMED FILES:" & vbCrLf & vbCrLf

    ' If we used the manual fallback, don't try to rename
    If usedManualFallback Then
        renamedSummary = renamedSummary & _
            "(Using manual fallback - files will keep original names)" & vbCrLf
        Step6_BuildRenamedList = True
        Exit Function
    End If

    ' Rename all files based on the naming scheme
    For i = 0 To UBound(docNames)

        currentFile = GetFileName(CStr(docNames(i)))
        newName = GetNewBaseName(currentFile)

        If newName = "" Then
            ' Keep original filename for unmapped files
            renamedSummary = renamedSummary & _
                currentFile & " -> (original name)" & vbCrLf
            skippedCount = skippedCount + 1
            newName = currentFile
        Else
            renamedSummary = renamedSummary & _
                currentFile & " -> " & newName & vbCrLf
        End If

        ' Store ONLY the filename
        docNames(i) = newName

    Next i

    If skippedCount > 0 Then
        renamedSummary = renamedSummary & vbCrLf & _
            "(" & skippedCount & " file(s) kept with original names)"
    End If

    Step6_BuildRenamedList = True

End Function


'============================================================
' STEP 7: APPLY RENAMED DOCUMENT LIST
'============================================================

Function Step7_ApplyRenamedList() As Boolean

    Dim result As Boolean

    Step7_ApplyRenamedList = False

    ' If we used the manual fallback, skip applying renames
    If usedManualFallback Then
        Step7_ApplyRenamedList = True
        Exit Function
    End If

    result = swPackAndGo.SetDocumentSaveToNames(docNames)

    If result Then
        Step7_ApplyRenamedList = True
    Else
        MsgBox "SolidWorks rejected the renamed document list.", _
               vbCritical, "Step 7"
        Exit Function
    End If

End Function


'============================================================
' STEP 8: CONFIRM AND SAVE
'============================================================

Sub Step8_ConfirmAndSave()

    Dim confirmMsg As String

    confirmMsg = renamedSummary & vbCrLf & vbCrLf & _
                 "Ready to save Pack and Go?" & vbCrLf & vbCrLf & _
                 "Destination:" & vbCrLf & destFolder

    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "Step 8: Confirm") = vbYes Then

        If swPackAndGo Is Nothing Then
            MsgBox "Pack and Go object is null.", vbCritical, "Step 8"
            Exit Sub
        End If

        On Error GoTo SaveError

        swModel.Extension.SavePackAndGo swPackAndGo

        MsgBox "Pack and Go completed!" & vbCrLf & vbCrLf & _
               "Destination:" & vbCrLf & destFolder & vbCrLf & vbCrLf & _
               "Main assembly:" & vbCrLf & masterName & ".SLDASM", _
               vbInformation, "Pack and Go Complete"

        ' Close the original assembly WITHOUT saving to preserve the template
        swApp.CloseDoc swModel.GetPathName

        ' Open the newly created packed assembly
        Dim packedAssemblyPath As String
        Dim swPackedAssembly As SldWorks.ModelDoc2
        packedAssemblyPath = destFolder & masterName & ".SLDASM"
        Set swPackedAssembly = swApp.OpenDoc6(packedAssemblyPath, swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)

        ' Rebuild the newly opened packed assembly so all references/views resolve cleanly
        If Not swPackedAssembly Is Nothing Then
            swPackedAssembly.ForceRebuild3 False
        End If

        ' Ask what the user needs and kick off the relevant export macros
        PromptAndRunExports

        Exit Sub

SaveError:
        MsgBox "Error saving Pack and Go:" & vbCrLf & Err.Description, _
               vbCritical, "Step 8 Error"

    End If

End Sub


'============================================================
' STEP 9: DELETE EXCLUDED FILES
'============================================================

Sub Step9_DeleteExcludedFiles()

    Dim i As Long
    Dim excludedFile As String
    Dim filePath As String
    Dim fso As Object

    If Not filteredExcludedFiles Then
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Delete each excluded file from destination folder
    For i = LBound(excludedFiles) To UBound(excludedFiles)
        excludedFile = CStr(excludedFiles(i))
        filePath = destFolder & excludedFile

        On Error Resume Next

        If fso.FileExists(filePath) Then
            fso.DeleteFile filePath, True
        End If

        On Error GoTo 0
    Next i

End Sub


'============================================================
' POST-PACK-AND-GO: PROMPT FOR EXPORT TYPES AND RUN THEM
'
' Runs on the newly opened/rebuilt packed assembly.
' If this is just for a quote, only the assembly PDF drawings
' are created. Otherwise, the user is asked individually about
' DXF, STEP, and PDF exports. DXF and STEP are component-only
' (no assemblies); PDF covers both assemblies and components.
'============================================================

Sub PromptAndRunExports()

    Dim isQuoteOnly As VbMsgBoxResult
    Dim wantDXF As VbMsgBoxResult
    Dim wantSTEP As VbMsgBoxResult
    Dim wantPDF As VbMsgBoxResult

    isQuoteOnly = MsgBox( _
        "Is this Pack and Go just for a quote?" & vbCrLf & vbCrLf & _
        "(If yes, only the assembly PDF drawings will be created - " & _
        "no DXF, STEP, or component PDFs.)", _
        vbQuestion + vbYesNo, "Quote Only?")

    If isQuoteOnly = vbYes Then
        RunExternalMacro MACRO_FOLDER & "pdfassemblies.swp", "RunPDFAssemblyDrawings"
        Exit Sub
    End If

    wantDXF = MsgBox("Do you need DXF files?", vbQuestion + vbYesNo, "DXF Files?")
    wantSTEP = MsgBox("Do you need STEP files?", vbQuestion + vbYesNo, "STEP Files?")
    wantPDF = MsgBox("Do you need PDF files?", vbQuestion + vbYesNo, "PDF Files?")

    If wantDXF <> vbYes And wantSTEP <> vbYes And wantPDF <> vbYes Then
        MsgBox "No export types selected. Nothing was exported.", vbInformation, "Nothing To Do"
        Exit Sub
    End If

    ' DXF and STEP are COMPONENTS ONLY (no assembly DXF/STEP)
    If wantDXF = vbYes Then
        RunExternalMacro MACRO_FOLDER & "dxfcomponents.swp", "RunDXFComponentDrawings"
    End If

    If wantSTEP = vbYes Then
        RunExternalMacro MACRO_FOLDER & "stepcomponents.swp", "RunSTEPComponentExport"
    End If

    ' PDF covers both the main/sub-assembly drawings and the component drawings
    If wantPDF = vbYes Then
        RunExternalMacro MACRO_FOLDER & "pdfassemblies.swp", "RunPDFAssemblyDrawings"
        RunExternalMacro MACRO_FOLDER & "pdfcomponents.swp", "RunPDFComponentDrawings"
    End If

End Sub


'============================================================
' HELPER: DERIVE THE VBA MODULE NAME FROM THE MACRO FILENAME
'
' SolidWorks names the default module "<basefilename>1" when a
' macro is created (verified against the live macros on the
' network drive: pdfassemblies.swp -> "pdfassemblies1", etc.)
'============================================================

Function DeriveModuleName(macroFile As String) As String

    Dim fileName As String
    Dim baseName As String

    fileName = Mid(macroFile, InStrRev(macroFile, "\") + 1)

    If InStrRev(fileName, ".") > 0 Then
        baseName = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        baseName = fileName
    End If

    DeriveModuleName = baseName & "1"

End Function


'============================================================
' HELPER: RUN AN EXTERNAL STANDALONE MACRO FILE
'============================================================

Function RunExternalMacro(macroFile As String, procName As String) As Boolean

    Dim runErr As Long
    Dim ok As Boolean
    Dim moduleName As String

    RunExternalMacro = False

    If Dir(macroFile) = "" Then
        MsgBox "Macro file not found:" & vbCrLf & macroFile, vbCritical, "Macro Not Found"
        Exit Function
    End If

    moduleName = DeriveModuleName(macroFile)

    ok = swApp.RunMacro2(macroFile, moduleName, procName, _
                          swRunMacroOption_e.swRunMacroUnloadAfterRun, runErr)

    ' RunMacro2's boolean return value is unreliable in some SolidWorks
    ' versions - it can report False even though the macro ran cleanly.
    ' The Error out-parameter is the trustworthy signal: 0 means
    ' swRunMacroDone (no error), so treat that as success too.
    If Not ok And runErr = 0 Then
        ok = True
    End If

    If Not ok Then
        MsgBox "Failed to run macro:" & vbCrLf & procName & vbCrLf & macroFile & vbCrLf & vbCrLf & _
               "Error code: " & runErr & vbCrLf & vbCrLf & _
               "Tried module name: " & moduleName & vbCrLf & _
               "If that macro's module was renamed manually, update DeriveModuleName.", _
               vbCritical, "Macro Error"
    End If

    RunExternalMacro = ok

End Function


'============================================================
' HELPER FUNCTION: CLEAR FOLDER CONTENTS
'
' Deletes all files and subdirectories in the folder
' Used when user chooses to overwrite an existing folder
'============================================================

Private Sub ClearFolderContents(folderPath As String)

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        Exit Sub
    End If

    Set folder = fso.GetFolder(folderPath)

    ' Delete all files in the folder
    On Error Resume Next
    For Each file In folder.Files
        file.Delete True
    Next file

    ' Delete all subfolders in the folder
    For Each subFolder In folder.SubFolders
        subFolder.Delete True
    Next subFolder
    On Error GoTo 0

End Sub


'============================================================
' UTILITY: DEBUG - SHOW CURRENT STATE
'============================================================

Sub Debug_ShowState()

    Dim msg As String

    msg = "swModel: " & IIf(swModel Is Nothing, "(not set)", swModel.GetPathName) & vbCrLf
    msg = msg & "destFolder: " & destFolder & vbCrLf
    msg = msg & "masterName: " & masterName & vbCrLf
    msg = msg & "docCount: " & docCount & vbCrLf
    msg = msg & "swPackAndGo: " & IIf(swPackAndGo Is Nothing, "(not set)", "(set)")

    MsgBox msg, vbInformation, "Current State"

End Sub


'============================================================
' UTILITY FUNCTIONS
'============================================================

Function GetGlobalVariable( _
    eqMgr As SldWorks.EquationMgr, _
    variableName As String) As String

    Dim i As Integer
    Dim equationText As String
    Dim value As Double

    For i = 0 To eqMgr.GetCount - 1

        If eqMgr.GlobalVariable(i) Then

            equationText = eqMgr.Equation(i)

            If InStr(1, equationText, """" & variableName & """", vbTextCompare) > 0 Then

                value = eqMgr.value(i)
                GetGlobalVariable = CStr(value)
                Exit Function

            End If

        End If

    Next i

    MsgBox "Could not find global variable:" & vbCrLf & vbCrLf & variableName, _
           vbCritical, "Global Variable Missing"

    GetGlobalVariable = "ERROR"

End Function


Function FormatPNDimension(value As Double) As String

    FormatPNDimension = CStr(Int(value * 100 + 0.5))

End Function


Function GetFileName(fullPath As String) As String

    Dim lastBackslash As Integer

    lastBackslash = InStrRev(fullPath, "\")

    If lastBackslash > 0 Then
        GetFileName = Mid$(fullPath, lastBackslash + 1)
    Else
        GetFileName = fullPath
    End If

End Function


Function GetNewBaseName(currentFile As String) As String

    Select Case LCase$(currentFile)

        '======================================================
        ' 1. TOP-LEVEL ASSEMBLY
        '======================================================

        Case "planter_assembly.sldasm"
            GetNewBaseName = masterName & ".SLDASM"

        Case "planter_assembly.slddrw"
            GetNewBaseName = masterName & ".SLDDRW"


        '======================================================
        ' 2. CENTER PLANTER
        '======================================================

        Case "planter_center_asm.sldasm"
            GetNewBaseName = centerName & ".SLDASM"

        Case "planter_center_asm.slddrw"
            GetNewBaseName = centerName & ".SLDDRW"

        Case "planterwallcenter.sldprt"
            GetNewBaseName = centerName & "-W.SLDPRT"

        Case "planterwallcenter.slddrw"
            GetNewBaseName = centerName & "-W.SLDDRW"

        Case "baseplatecenter.sldprt"
            GetNewBaseName = centerHardwareName & "-BP.SLDPRT"

        Case "baseplatecenter.slddrw"
            GetNewBaseName = centerHardwareName & "-BP.SLDDRW"

        Case "centergusset.sldprt"
            GetNewBaseName = centerGussetName & ".SLDPRT"

        Case "centergusset.slddrw"
            GetNewBaseName = centerGussetName & ".SLDDRW"

        Case "mountinggusset.sldprt"
            GetNewBaseName = mountingGussetName & ".SLDPRT"

        Case "mountinggusset.slddrw"
            GetNewBaseName = mountingGussetName & ".SLDDRW"


        '======================================================
        ' 3. SIDE ASSEMBLIES
        '======================================================

        Case "planter_side_asm.sldasm"
            GetNewBaseName = sideName & ".SLDASM"

        Case "planter_side_asm.slddrw"
            GetNewBaseName = sideName & ".SLDDRW"

        Case "planter_side_asm_left.sldasm"
            GetNewBaseName = sideLeftName & ".SLDASM"

        Case "planter_side_asm_left.slddrw"
            GetNewBaseName = sideLeftName & ".SLDDRW"

        Case "planter_side_asm_right.sldasm"
            GetNewBaseName = sideRightName & ".SLDASM"

        Case "planter_side_asm_right.slddrw"
            GetNewBaseName = sideRightName & ".SLDDRW"

        Case "planter_side_asm_center.sldasm"
            GetNewBaseName = sideCenterName & ".SLDASM"

        Case "planter_side_asm_center.slddrw"
            GetNewBaseName = sideCenterName & ".SLDDRW"


        '======================================================
        ' 4. SIDE WALLS
        '======================================================

        Case "planterwallside.sldprt"
            GetNewBaseName = sideName & "-W.SLDPRT"

        Case "planterwallside.slddrw"
            GetNewBaseName = sideName & "-W.SLDDRW"

        Case "planterwallsideleft.sldprt"
            GetNewBaseName = sideLeftName & "-W.SLDPRT"

        Case "planterwallsideleft.slddrw"
            GetNewBaseName = sideLeftName & "-W.SLDDRW"

        Case "planterwallsideright.sldprt"
            GetNewBaseName = sideRightName & "-W.SLDPRT"

        Case "planterwallsideright.slddrw"
            GetNewBaseName = sideRightName & "-W.SLDDRW"

        Case "planterwallsidecenter.sldprt"
            GetNewBaseName = sideCenterName & "-W.SLDPRT"

        Case "planterwallsidecenter.slddrw"
            GetNewBaseName = sideCenterName & "-W.SLDDRW"


        '======================================================
        ' 5. SIDE BASE PLATE
        '======================================================

        Case "baseplateside.sldprt"
            GetNewBaseName = sharedSideName & "-BP.SLDPRT"

        Case "baseplateside.slddrw"
            GetNewBaseName = sharedSideName & "-BP.SLDDRW"


        '======================================================
        ' 6. RETURN WALLS
        '======================================================

        Case "leftsidewall.sldprt"
            GetNewBaseName = returnType & "-" & sideWallName & "-LW.SLDPRT"

        Case "leftsidewall.slddrw"
            GetNewBaseName = returnType & "-" & sideWallName & "-LW.SLDDRW"

        Case "rightsidewall.sldprt"
            GetNewBaseName = returnType & "-" & sideWallName & "-RW.SLDPRT"

        Case "rightsidewall.slddrw"
            GetNewBaseName = returnType & "-" & sideWallName & "-RW.SLDDRW"


        '======================================================
        ' NO MATCH - UNMAPPED FILE
        '======================================================

        Case Else
            GetNewBaseName = ""

    End Select

End Function


'============================================================
' MANUAL FALLBACK: WALK THE ASSEMBLY TREE
'
' Used when Pack and Go's native GetDocumentNames fails.
' Recursively collects all component paths and their drawings.
'============================================================

Private Function GetPackAndGoDocuments_Manual() As Boolean

    Dim paths As Collection
    Dim swConf As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    Dim topPath As String
    Dim i As Long
    Dim n As Long
    Dim arr() As String

    GetPackAndGoDocuments_Manual = False

    Set paths = New Collection

    topPath = swModel.GetPathName
    AddUniquePath paths, topPath
    AddDrawingIfExists paths, topPath

    Set swConf = swModel.GetActiveConfiguration

    If Not swConf Is Nothing Then
        Set swRootComp = swConf.GetRootComponent3(True)
        If Not swRootComp Is Nothing Then
            CollectComponentPaths swRootComp, paths
        End If
    End If

    n = paths.Count
    If n = 0 Then Exit Function

    ReDim arr(0 To n - 1)
    For i = 1 To n
        arr(i - 1) = paths(i)
    Next i

    docNames = arr
    docCount = n

    GetPackAndGoDocuments_Manual = True

End Function


'============================================================
' RECURSIVELY COLLECT COMPONENT FILE PATHS + DRAWINGS
'============================================================

Private Sub CollectComponentPaths(swComp As SldWorks.Component2, paths As Collection)

    Dim vChildren As Variant
    Dim swChild As SldWorks.Component2
    Dim i As Long
    Dim p As String

    p = swComp.GetPathName

    If p <> "" Then
        AddUniquePath paths, p
        AddDrawingIfExists paths, p
    End If

    vChildren = swComp.GetChildren

    If Not IsEmpty(vChildren) Then
        For i = 0 To UBound(vChildren)
            Set swChild = vChildren(i)
            CollectComponentPaths swChild, paths
        Next i
    End If

End Sub


'============================================================
' COLLECTION HELPERS
'============================================================

Private Sub AddUniquePath(paths As Collection, p As String)

    If p = "" Then Exit Sub

    On Error Resume Next
    paths.Add p, p
    On Error GoTo 0

End Sub


Private Sub AddDrawingIfExists(paths As Collection, partOrAsmPath As String)

    Dim drawingPath As String
    Dim dotPos As Long

    dotPos = InStrRev(partOrAsmPath, ".")
    If dotPos = 0 Then Exit Sub

    drawingPath = Left$(partOrAsmPath, dotPos) & "SLDDRW"

    If Dir(drawingPath) <> "" Then
        AddUniquePath paths, drawingPath
    End If


End Sub


