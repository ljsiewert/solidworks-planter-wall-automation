Option Explicit

'============================================================
' >>> DEPLOY TARGET <<<
' File:   Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\dxfcomponents.swp
' Module: dxfcomponents1 (confirmed - matches what you already created)
' Action: REPLACE EXISTING CONTENT. Tools > Macro > Edit on
'         dxfcomponents.swp, select all code in module dxfcomponents1,
'         delete, paste this file's content in, save.
' Changed: COMPLETE REWRITE. Previously exported a plain 2D drawing view
'         of each component to DXF. Laser cutting needs the FLAT PATTERN
'         instead, so this now works directly on the sheet-metal PART
'         documents already loaded in the open assembly (same approach
'         as stepcomponents.bas) and exports each part's flat pattern
'         straight to DXF via IPartDoc.ExportFlatPatternView - no drawing
'         file involved at all. Non-sheet-metal parts are skipped
'         automatically (no flat pattern feature to export).
' Changed 2: Added ESC-key cancel support. New CancelRequested function
'         polls for the ESC key between each part; pressing it prompts
'         to confirm, then stops cleanly and reports how many were
'         completed before cancelling.
'============================================================

'============================================================
' PACK AND GO - COMPONENT FLAT PATTERN DXF EXPORT
'
' Exports every active (non-suppressed) sheet-metal PART
' component in the open assembly to DXF as its FLAT PATTERN
' (not a 2D drawing view). This is what the laser cutter needs.
'
' Non-sheet-metal parts (no Flat-Pattern feature) are skipped.
' Components only - no assembly DXF.
'
' Works directly on the part documents already loaded in memory
' as part of the open assembly (same technique as
' stepcomponents.bas). The Flat-Pattern feature is temporarily
' unsuppressed if needed to export, then restored to whatever
' suppression state it was in before - so the assembly is left
' unchanged.
'
' Run from the main assembly after Pack and Go completes.
'============================================================

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2

Public destFolder As String
Public dxfFolder As String

' Windows API for ESC-key cancel polling (see CancelRequested below)
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If


'============================================================
' CHECK FOR USER CANCEL REQUEST (press ESC between items)
'
' Call this once per loop iteration. DoEvents lets Windows
' deliver the keypress to this process even though SolidWorks
' is busy running the macro. If ESC is currently held down,
' confirm with the user before actually cancelling.
'============================================================

Function CancelRequested() As Boolean

    CancelRequested = False

    DoEvents

    If (GetAsyncKeyState(vbKeyEscape) And &H8000) <> 0 Then
        If MsgBox("Cancel the DXF export now?" & vbCrLf & vbCrLf & _
                  "Parts already exported will be kept." & vbCrLf & _
                  "Click No to keep going.", _
                  vbQuestion + vbYesNo, "Cancel Export?") = vbYes Then
            CancelRequested = True
        End If
    End If

End Function

'============================================================
' MAIN ENTRY POINT
'============================================================

Sub RunDXFComponentDrawings()

    Dim assemblyPath As String
    Dim activeParts As collection
    Dim i As Long
    Dim successCount As Long
    Dim failureCount As Long
    Dim skippedCount As Long
    Dim exportResult As Integer

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "Please run this from the packed assembly (from PACK_TEST folder).", vbExclamation
        Exit Sub
    End If

    ' Get the folder path from the active assembly
    assemblyPath = swModel.GetPathName
    If assemblyPath = "" Then
        MsgBox "Assembly has not been saved.", vbExclamation
        Exit Sub
    End If

    destFolder = Left(assemblyPath, InStrRev(assemblyPath, "\") - 1) & "\"
    dxfFolder = destFolder & "DXF\"

    ' Verify this is a Pack and Go folder
    If InStr(1, destFolder, "PACK_TEST", vbTextCompare) = 0 Then
        MsgBox "This assembly does not appear to be from a Pack and Go output folder." & vbCrLf & vbCrLf & _
               "Please open the assembly from:" & vbCrLf & _
               "Z:\...\PACK_TEST\<part_number>\", vbExclamation
        Exit Sub
    End If

    ' Create DXF folder if needed
    If Dir(dxfFolder, vbDirectory) = "" Then
        MkDir dxfFolder
    Else
        If HasExistingDXFs(dxfFolder) Then
            If Not PromptOverwriteDXFs() Then
                Exit Sub
            End If
            ClearDXFFolder dxfFolder
        End If
    End If

    ' Recursively collect unique active (non-suppressed) part components
    Set activeParts = New collection
    Dim swConfig As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2

    Set swConfig = swModel.GetActiveConfiguration
    Set swRootComp = swConfig.GetRootComponent3(True)

    If swRootComp Is Nothing Then
        MsgBox "Could not get the root assembly component.", vbExclamation
        Exit Sub
    End If

    CollectActivePartComponents swRootComp, activeParts

    If activeParts.count = 0 Then
        MsgBox "No active part components found in the assembly.", vbExclamation
        Exit Sub
    End If

    MsgBox "Found " & activeParts.count & " active part components." & vbCrLf & _
           "Exporting flat patterns to DXF..." & vbCrLf & vbCrLf & _
           "Folder: " & dxfFolder & vbCrLf & vbCrLf & _
           "Non-sheet-metal parts will be skipped automatically." & vbCrLf & vbCrLf & _
           "Press ESC at any time to cancel after the current part finishes." & vbCrLf & vbCrLf & _
           "This may take a few minutes.", vbInformation, "Starting DXF Export"

    successCount = 0
    failureCount = 0
    skippedCount = 0

    Dim wasCancelled As Boolean
    wasCancelled = False

    For i = 1 To activeParts.count
        If CancelRequested() Then
            wasCancelled = True
            Exit For
        End If

        exportResult = ExportComponentFlatPatternToDXF(activeParts(i))
        Select Case exportResult
            Case 1
                successCount = successCount + 1
            Case 0
                failureCount = failureCount + 1
            Case -1
                skippedCount = skippedCount + 1
        End Select
    Next i

    If wasCancelled Then
        MsgBox "DXF export cancelled by user." & vbCrLf & vbCrLf & _
               "Successfully exported before cancelling: " & successCount & " of " & activeParts.count & vbCrLf & _
               "Skipped (not sheet metal): " & skippedCount & vbCrLf & _
               "Failed: " & failureCount & vbCrLf & vbCrLf & _
               "Files saved to:" & vbCrLf & dxfFolder, vbExclamation, "Cancelled"
    Else
        MsgBox "DXF flat pattern export complete!" & vbCrLf & vbCrLf & _
               "Successfully exported: " & successCount & vbCrLf & _
               "Skipped (not sheet metal): " & skippedCount & vbCrLf & _
               "Failed: " & failureCount & vbCrLf & vbCrLf & _
               "Files saved to:" & vbCrLf & dxfFolder, vbInformation, "Complete"
    End If

End Sub


'============================================================
' RECURSIVELY COLLECT UNIQUE ACTIVE (NON-SUPPRESSED) PART
' COMPONENTS (SKIPS SUB-ASSEMBLIES - PARTS ONLY)
'============================================================

Sub CollectActivePartComponents(swComp As SldWorks.Component2, ByRef collectionOut As collection)

    Dim children() As Object
    Dim i As Long
    Dim childComp As SldWorks.Component2
    Dim modelDoc As SldWorks.ModelDoc2
    Dim filePath As String
    Dim fileName As String

    ' Skip suppressed components entirely (and their children)
    If swComp.IsSuppressed Then
        Exit Sub
    End If

    Set modelDoc = swComp.GetModelDoc2

    If Not modelDoc Is Nothing Then
        filePath = modelDoc.GetPathName
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

        ' Only PART files (.SLDPRT) - skip assemblies (.SLDASM)
        If LCase(Right(fileName, 7)) = ".sldprt" Then
            If Not AlreadyCollected(collectionOut, filePath) Then
                collectionOut.Add swComp, filePath
            End If
        End If
    End If

    children = swComp.GetChildren

    If Not IsEmpty(children) Then
        For i = LBound(children) To UBound(children)
            If Not children(i) Is Nothing Then
                Set childComp = children(i)
                CollectActivePartComponents childComp, collectionOut
            End If
        Next i
    End If

End Sub


'============================================================
' CHECK IF A FILE PATH HAS ALREADY BEEN COLLECTED
' (avoids exporting the same part twice when used multiple
' times in the assembly)
'============================================================

Function AlreadyCollected(col As collection, keyName As String) As Boolean

    Dim dummy As Object

    On Error GoTo NotFound
    Set dummy = col.Item(keyName)
    AlreadyCollected = True
    Exit Function

NotFound:
    AlreadyCollected = False

End Function


'============================================================
' EXPORT A SINGLE COMPONENT'S FLAT PATTERN TO DXF
'
' Returns: 1 = success, 0 = failure, -1 = skipped (not sheet metal)
'
' Finds the part's "Flat-Pattern" feature, temporarily
' unsuppresses it if needed, calls ExportFlatPatternView to
' write the DXF directly (no drawing file involved), then
' restores the feature's original suppression state so the
' assembly/part is left exactly as it was.
'============================================================

Function ExportComponentFlatPatternToDXF(swComp As SldWorks.Component2) As Integer

    Const swDxfDwgFileFormat_FlatPattern_DXF As Long = 1

    Dim modelDoc As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc
    Dim swFlatFeat As SldWorks.Feature
    Dim filePath As String
    Dim fileName As String
    Dim exportPath As String
    Dim wasSuppressed As Boolean
    Dim exportOk As Boolean

    ExportComponentFlatPatternToDXF = 0

    On Error GoTo ErrorHandler

    Set modelDoc = swComp.GetModelDoc2
    If modelDoc Is Nothing Then
        Exit Function
    End If

    Set swFlatFeat = FindFlatPatternFeature(modelDoc)

    ' Not a sheet metal part (no flat pattern feature) - skip it
    If swFlatFeat Is Nothing Then
        ExportComponentFlatPatternToDXF = -1
        Exit Function
    End If

    filePath = modelDoc.GetPathName
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)

    exportPath = dxfFolder & fileName & ".DXF"

    Set swPart = modelDoc

    ' Remember current suppression state, unsuppress to export if needed
    wasSuppressed = swFlatFeat.IsSuppressed

    If wasSuppressed Then
        swFlatFeat.SetSuppression2 swFeatureSuppressionAction_e.swUnSuppressFeature, _
                                    swInConfigurationOpts_e.swThisConfiguration, Nothing
        modelDoc.EditRebuild3
    End If

    exportOk = swPart.ExportFlatPatternView(exportPath, swDxfDwgFileFormat_FlatPattern_DXF)

    ' Restore original suppression state so the part is left unchanged
    If wasSuppressed Then
        swFlatFeat.SetSuppression2 swFeatureSuppressionAction_e.swSuppressFeature, _
                                    swInConfigurationOpts_e.swThisConfiguration, Nothing
        modelDoc.EditRebuild3
    End If

    If exportOk Then
        ExportComponentFlatPatternToDXF = 1
    Else
        ExportComponentFlatPatternToDXF = 0
    End If

    Exit Function

ErrorHandler:
    ExportComponentFlatPatternToDXF = 0

End Function


'============================================================
' FIND THE "FLAT-PATTERN" FEATURE ON A PART (IF ANY)
' Returns Nothing if the part is not sheet metal.
'============================================================

Function FindFlatPatternFeature(modelDoc As SldWorks.ModelDoc2) As SldWorks.Feature

    Dim swFeat As SldWorks.Feature

    Set FindFlatPatternFeature = Nothing

    Set swFeat = modelDoc.FirstFeature

    Do While Not swFeat Is Nothing
        If InStr(1, swFeat.GetTypeName2, "FlatPattern", vbTextCompare) > 0 Then
            Set FindFlatPatternFeature = swFeat
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

End Function


'============================================================
' CHECK IF DXF FOLDER HAS EXISTING DXF FILES
'============================================================

Function HasExistingDXFs(folderPath As String) As Boolean

    Dim fileName As String

    HasExistingDXFs = False

    fileName = Dir(folderPath & "*.DXF")

    If fileName <> "" Then
        HasExistingDXFs = True
    End If

End Function


'============================================================
' PROMPT USER TO OVERWRITE DXF FILES
' Returns True if user chooses OVERWRITE, False if CANCEL
'============================================================

Function PromptOverwriteDXFs() As Boolean

    Dim userChoice As VbMsgBoxResult

    PromptOverwriteDXFs = False

    userChoice = MsgBox("DXF files already exist in the DXF folder." & vbCrLf & vbCrLf & _
                       "Would you like to overwrite them?" & vbCrLf & vbCrLf & _
                       "Click YES to overwrite and continue." & vbCrLf & _
                       "Click NO to cancel this operation.", _
                       vbQuestion + vbYesNo, _
                       "Existing DXF Files Found")

    If userChoice = vbYes Then
        PromptOverwriteDXFs = True
    End If

End Function


'============================================================
' CLEAR ALL DXF FILES FROM FOLDER
'============================================================

Sub ClearDXFFolder(folderPath As String)

    Dim fileName As String
    Dim filePath As String

    On Error Resume Next

    fileName = Dir(folderPath & "*.DXF")

    Do While fileName <> ""
        filePath = folderPath & fileName
        Kill filePath
        fileName = Dir
    Loop

    On Error GoTo 0

End Sub
