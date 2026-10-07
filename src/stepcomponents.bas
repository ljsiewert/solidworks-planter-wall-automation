Option Explicit

'============================================================
' >>> DEPLOY TARGET <<<
' File:   Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\stepcomponents.swp
' Module: stepcomponents1
' Action: UPDATE NEEDED. Tools > Macro > Edit on stepcomponents.swp,
'         select all code in module stepcomponents1, delete, paste this
'         file's content in, save.
' Purpose: Exports active, non-suppressed PART components (not
'         assemblies) directly from the open assembly to STEP, using
'         SaveAs with swSaveAsOptions_Silent so it doesn't disturb the
'         assembly's live references to those parts.
' Changed: Fixed a compile error - swSaveAsOptions_SaveCopy is not a
'         valid constant, removed it (swSaveAsOptions_Silent alone is
'         correct).
' Changed 2: Added ESC-key cancel support. New CancelRequested function
'         polls for the ESC key between each part; pressing it prompts
'         to confirm, then stops cleanly and reports how many were
'         completed before cancelling.
'============================================================

'============================================================
' PACK AND GO - COMPONENT STEP EXPORT
'
' Exports every active (non-suppressed) PART component in the
' open assembly to STEP. Components only - no assembly STEP.
'
' Unlike the PDF/DXF component exporters, this does NOT work
' off drawing files - it works directly on the part documents
' already loaded in memory as part of the open assembly, and
' exports a COPY (swSaveAsOptions_SaveCopy) so the assembly's
' references to those parts are left untouched.
'
' Run from the main assembly after Pack and Go completes.
'============================================================

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2

Public destFolder As String
Public stepFolder As String

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
        If MsgBox("Cancel the STEP export now?" & vbCrLf & vbCrLf & _
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

Sub RunSTEPComponentExport()

    Dim assemblyPath As String
    Dim activeParts As collection
    Dim i As Long
    Dim successCount As Long
    Dim failureCount As Long

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
    stepFolder = destFolder & "STEP\"

    ' Verify this is a Pack and Go folder
    If InStr(1, destFolder, "PACK_TEST", vbTextCompare) = 0 Then
        MsgBox "This assembly does not appear to be from a Pack and Go output folder." & vbCrLf & vbCrLf & _
               "Please open the assembly from:" & vbCrLf & _
               "Z:\...\PACK_TEST\<part_number>\", vbExclamation
        Exit Sub
    End If

    ' Create STEP folder if needed
    If Dir(stepFolder, vbDirectory) = "" Then
        MkDir stepFolder
    Else
        If HasExistingSTEPs(stepFolder) Then
            If Not PromptOverwriteSTEPs() Then
                Exit Sub
            End If
            ClearSTEPFolder stepFolder
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
           "Exporting to STEP..." & vbCrLf & vbCrLf & _
           "Folder: " & stepFolder & vbCrLf & vbCrLf & _
           "Press ESC at any time to cancel after the current part finishes." & vbCrLf & vbCrLf & _
           "This may take a few minutes.", vbInformation, "Starting STEP Export"

    successCount = 0
    failureCount = 0

    Dim wasCancelled As Boolean
    wasCancelled = False

    For i = 1 To activeParts.count
        If CancelRequested() Then
            wasCancelled = True
            Exit For
        End If

        If ExportComponentToSTEP(activeParts(i)) Then
            successCount = successCount + 1
        Else
            failureCount = failureCount + 1
        End If
    Next i

    If wasCancelled Then
        MsgBox "STEP export cancelled by user." & vbCrLf & vbCrLf & _
               "Successfully exported before cancelling: " & successCount & " of " & activeParts.count & vbCrLf & _
               "Failed: " & failureCount & vbCrLf & vbCrLf & _
               "Files saved to:" & vbCrLf & stepFolder, vbExclamation, "Cancelled"
    Else
        MsgBox "STEP export complete!" & vbCrLf & vbCrLf & _
               "Successfully exported: " & successCount & vbCrLf & _
               "Failed: " & failureCount & vbCrLf & vbCrLf & _
               "Files saved to:" & vbCrLf & stepFolder, vbInformation, "Complete"
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
' EXPORT A SINGLE COMPONENT TO STEP
'
' Uses swSaveAsOptions_SaveCopy so the live assembly's
' reference to this component is left untouched.
'============================================================

Function ExportComponentToSTEP(swComp As SldWorks.Component2) As Boolean

    Dim modelDoc As SldWorks.ModelDoc2
    Dim filePath As String
    Dim fileName As String
    Dim exportPath As String
    Dim errors As Long
    Dim warnings As Long
    Dim saveResult As Boolean

    ExportComponentToSTEP = False

    On Error GoTo ErrorHandler

    Set modelDoc = swComp.GetModelDoc2
    If modelDoc Is Nothing Then
        Exit Function
    End If

    filePath = modelDoc.GetPathName
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)

    exportPath = stepFolder & fileName & ".STEP"

    errors = 0
    warnings = 0

    saveResult = modelDoc.Extension.SaveAs( _
                    exportPath, _
                    swSaveAsCurrentVersion, _
                    swSaveAsOptions_e.swSaveAsOptions_Silent, _
                    Nothing, _
                    errors, _
                    warnings)

    ExportComponentToSTEP = saveResult
    Exit Function

ErrorHandler:
    ExportComponentToSTEP = False

End Function


'============================================================
' CHECK IF STEP FOLDER HAS EXISTING STEP FILES
'============================================================

Function HasExistingSTEPs(folderPath As String) As Boolean

    Dim fileName As String

    HasExistingSTEPs = False

    fileName = Dir(folderPath & "*.STEP")

    If fileName <> "" Then
        HasExistingSTEPs = True
    End If

End Function


'============================================================
' PROMPT USER TO OVERWRITE STEP FILES
' Returns True if user chooses OVERWRITE, False if CANCEL
'============================================================

Function PromptOverwriteSTEPs() As Boolean

    Dim userChoice As VbMsgBoxResult

    PromptOverwriteSTEPs = False

    userChoice = MsgBox("STEP files already exist in the STEP folder." & vbCrLf & vbCrLf & _
                       "Would you like to overwrite them?" & vbCrLf & vbCrLf & _
                       "Click YES to overwrite and continue." & vbCrLf & _
                       "Click NO to cancel this operation.", _
                       vbQuestion + vbYesNo, _
                       "Existing STEP Files Found")

    If userChoice = vbYes Then
        PromptOverwriteSTEPs = True
    End If

End Function


'============================================================
' CLEAR ALL STEP FILES FROM FOLDER
'============================================================

Sub ClearSTEPFolder(folderPath As String)

    Dim fileName As String
    Dim filePath As String

    On Error Resume Next

    fileName = Dir(folderPath & "*.STEP")

    Do While fileName <> ""
        filePath = folderPath & fileName
        Kill filePath
        fileName = Dir
    Loop

    On Error GoTo 0

End Sub
