Option Explicit

'============================================================
' >>> DEPLOY TARGET <<<
' File:   Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\pdfcomponents.swp
' Module: pdfcomponents1
' Action: UPDATE NEEDED. Tools > Macro > Edit on pdfcomponents.swp,
'         select all code in module pdfcomponents1, delete, paste this
'         file's content in, save.
' Changed: Added swApp.UserControlBackground = True/False around the
'         drawing processing loop so drawings open/close silently in the
'         background (no window flashing/stealing focus) - faster and
'         less disruptive. Also removed a leftover debug MsgBox that
'         required a manual click every run.
' Changed 2: Output folder renamed from PDF\ to PDF_Components\ so
'         assembly and component PDFs no longer land in the same folder.
' Changed 3: Added ESC-key cancel support. New CancelRequested function
'         polls for the ESC key between each drawing; pressing it
'         prompts to confirm, then stops cleanly after the current
'         drawing (no corrupted/half-open state) and reports how many
'         were completed before cancelling.
'============================================================

'============================================================
' PACK AND GO - COMPONENT DRAWING PDF EXPORT
'
' Exports component drawings to PDF with automatic:
' - View positioning and scaling
' - Dimension auto-arrangement
'
' Run from the main assembly after Pack and Go completes.
' Reads component drawings from the PACK_TEST output folder.
'============================================================

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2

Public packTestFolder As String
Public destFolder As String
Public pdfFolder As String

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
        If MsgBox("Cancel the PDF export now?" & vbCrLf & vbCrLf & _
                  "Drawings already exported will be kept." & vbCrLf & _
                  "Click No to keep going.", _
                  vbQuestion + vbYesNo, "Cancel Export?") = vbYes Then
            CancelRequested = True
        End If
    End If

End Function

'============================================================
' MAIN ENTRY POINT
'============================================================

Sub RunPDFComponentDrawings()

    Dim assemblyPath As String
    Dim fileList() As String
    Dim fileCount As Long
    Dim i As Long
    Dim drawingPath As String
    Dim success As Boolean

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

    ' Extract folder path
    destFolder = Left(assemblyPath, InStrRev(assemblyPath, "\") - 1) & "\"
    pdfFolder = destFolder & "PDF_Components\"

    ' Verify this is a Pack and Go folder
    If InStr(1, destFolder, "PACK_TEST", vbTextCompare) = 0 Then
        MsgBox "This assembly does not appear to be from a Pack and Go output folder." & vbCrLf & vbCrLf & _
               "Please open the assembly from:" & vbCrLf & _
               "Z:\...\PACK_TEST\<part_number>\", vbExclamation
        Exit Sub
    End If

    ' Get list of .SLDDRW files
    fileCount = GetDrawingFiles(destFolder, fileList)

    If fileCount = 0 Then
        MsgBox "No component drawings found in:" & vbCrLf & destFolder, vbExclamation
        Exit Sub
    End If

    ' Create PDF folder if needed
    If Dir(pdfFolder, vbDirectory) = "" Then
        MkDir pdfFolder
    Else
        ' PDF folder exists - check if it has PDFs and prompt user
        If HasExistingPDFs(pdfFolder) Then
            If Not PromptOverwritePDFs() Then
                ' User chose cancel
                Exit Sub
            End If
            ' User chose overwrite - clear the folder
            ClearPDFFolder pdfFolder
        End If
    End If

    ' Build list of ACTIVE component names (non-suppressed only)
    Dim activeComponents() As String
    Dim activeCount As Long
    activeCount = GetActiveComponentNames(swModel, activeComponents)

    If activeCount = 0 Then
        MsgBox "No active components found in the assembly.", vbExclamation
        Exit Sub
    End If

    MsgBox "Found " & activeCount & " active components." & vbCrLf & _
           "Processing component drawings..." & vbCrLf & vbCrLf & _
           "Folder: " & destFolder & vbCrLf & vbCrLf & _
           "Press ESC at any time to cancel after the current drawing finishes." & vbCrLf & vbCrLf & _
           "This may take a few minutes.", vbInformation, "Starting PDF Export"

    ' Open/close drawings silently in the background - no window
    ' flashing or stealing focus while processing, which is both
    ' faster and lets you keep working in the foreground.
    swApp.UserControlBackground = True

    ' Process each drawing (only open if it's for an active component)
    Dim wasCancelled As Boolean
    Dim processedCount As Long
    wasCancelled = False
    processedCount = 0

    For i = 0 To fileCount - 1
        If CancelRequested() Then
            wasCancelled = True
            Exit For
        End If

        drawingPath = destFolder & fileList(i)
        
        ' SKIP if drawing is not for an active component
        If IsDrawingForActiveComponent(fileList(i), activeComponents, activeCount) Then
            success = ProcessComponentDrawing(drawingPath)
            processedCount = processedCount + 1
            
            If Not success Then
                MsgBox "Failed to process: " & fileList(i), vbExclamation
            End If
        End If
    Next i

    swApp.UserControlBackground = False

    If wasCancelled Then
        MsgBox "PDF export cancelled by user." & vbCrLf & vbCrLf & _
               "Drawings processed before cancelling: " & processedCount & " of " & fileCount & vbCrLf & vbCrLf & _
               "Files saved so far:" & vbCrLf & pdfFolder, vbExclamation, "Cancelled"
    Else
        MsgBox fileCount & " component PDFs created successfully!" & vbCrLf & vbCrLf & _
               "Files saved to:" & vbCrLf & pdfFolder, vbInformation, "Complete"
    End If

End Sub


'============================================================
' GET LIST OF DRAWING FILES
'============================================================

Function GetDrawingFiles(folderPath As String, ByRef fileArray() As String) As Long

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim count As Long
    Dim tempArray() As String
    Dim i As Long
    Dim isDuplicate As Boolean

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    count = 0
    ReDim tempArray(100)

    For Each file In folder.Files
        If LCase(Right(file.Name, 7)) = ".slddrw" Then
            ' Check if this filename already exists in array
            isDuplicate = False
            For i = 0 To count - 1
                If LCase(tempArray(i)) = LCase(file.Name) Then
                    isDuplicate = True
                    Exit For
                End If
            Next i

            ' Only add if not a duplicate
            If Not isDuplicate Then
                tempArray(count) = file.Name
                count = count + 1
            End If
        End If
    Next file

    If count > 0 Then
        ReDim Preserve tempArray(count - 1)
        fileArray = tempArray
    End If

    GetDrawingFiles = count
    Exit Function

ErrorHandler:
    GetDrawingFiles = 0
End Function


'============================================================
' GET LIST OF ACTIVE COMPONENT NAMES (non-suppressed)
'============================================================

Function GetActiveComponentNames(swAssembly As SldWorks.ModelDoc2, ByRef componentNames() As String) As Long

    Dim swConfig As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    Dim activeComponents As collection
    Dim i As Long
    Dim count As Long

    Set activeComponents = New collection

    On Error GoTo ErrorHandler

    ' Get root component
    Set swConfig = swAssembly.GetActiveConfiguration
    Set swRootComp = swConfig.GetRootComponent3(True)

    If swRootComp Is Nothing Then
        GetActiveComponentNames = 0
        Exit Function
    End If

    ' Recursively collect all active component filenames
    CollectActiveComponentFiles swRootComp, activeComponents

    count = activeComponents.count

    ' Deduplicate the list
    Dim uniqueComponents As New collection
    Dim j As Long
    Dim isDuplicate As Boolean

    For i = 1 To count
        isDuplicate = False
        For j = 1 To uniqueComponents.count
            If LCase(uniqueComponents(j)) = LCase(activeComponents(i)) Then
                isDuplicate = True
                Exit For
            End If
        Next j

        If Not isDuplicate Then
            uniqueComponents.Add activeComponents(i)
        End If
    Next i

    count = uniqueComponents.count

    If count > 0 Then
        ReDim componentNames(count - 1)
        For i = 1 To count
            componentNames(i - 1) = uniqueComponents(i)
        Next i
    End If

    GetActiveComponentNames = count
    Exit Function

ErrorHandler:
    GetActiveComponentNames = 0

End Function


'============================================================
' RECURSIVELY COLLECT ACTIVE COMPONENT FILENAMES (PARTS ONLY)
'============================================================

Sub CollectActiveComponentFiles(swComp As SldWorks.Component2, ByRef collection As collection)

    Dim children() As Object
    Dim i As Long
    Dim childComp As SldWorks.Component2
    Dim fileName As String
    Dim filePath As String
    Dim modelDoc As SldWorks.ModelDoc2

    ' Get component model document
    Set modelDoc = swComp.GetModelDoc2

    If Not modelDoc Is Nothing Then
        filePath = modelDoc.GetPathName
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

        ' Debug output
        Debug.Print "Component: " & fileName

        ' Only add PART files (.SLDPRT), skip assemblies (.SLDASM)
        If LCase(Right(fileName, 7)) = ".sldprt" Then
            collection.Add fileName
            Debug.Print "  -> ADDED"
        Else
            Debug.Print "  -> SKIPPED (not a part)"
        End If
    End If

    ' Get children (only active/non-suppressed)
    children = swComp.GetChildren

    If Not IsEmpty(children) Then
        Debug.Print "Found " & UBound(children) - LBound(children) + 1 & " children"
        For i = LBound(children) To UBound(children)
            If Not children(i) Is Nothing Then
                Set childComp = children(i)
                ' Recursively add active children (parts only)
                CollectActiveComponentFiles childComp, collection
            End If
        Next i
    Else
        Debug.Print "No children"
    End If

End Sub


'============================================================
' CHECK IF DRAWING FILENAME MATCHES AN ACTIVE COMPONENT
'============================================================

Function IsDrawingForActiveComponent(drawingFileName As String, activeComponents() As String, activeCount As Long) As Boolean

    Dim i As Long
    Dim drawingBaseName As String
    Dim componentBaseName As String

    IsDrawingForActiveComponent = False

    ' Extract filename without extension from drawing
    drawingBaseName = Left(drawingFileName, InStrRev(drawingFileName, ".") - 1)

    ' Check if any active component matches this drawing
    For i = 0 To activeCount - 1
        ' Extract filename without extension from component
        componentBaseName = Left(activeComponents(i), InStrRev(activeComponents(i), ".") - 1)

        ' Case-insensitive match
        If LCase(drawingBaseName) = LCase(componentBaseName) Then
            IsDrawingForActiveComponent = True
            Exit Function
        End If
    Next i

End Function


'============================================================
' PROCESS SINGLE COMPONENT DRAWING
'============================================================

Function ProcessComponentDrawing(drawingPath As String) As Boolean

    Dim swDraw As SldWorks.DrawingDoc
    Dim swDrawModel As SldWorks.ModelDoc2
    Dim swView1 As SldWorks.View
    Dim swView2 As SldWorks.View
    Dim swView3 As SldWorks.View
    Dim pdfPath As String
    Dim fileName As String
    Dim saveError As Long

    ProcessComponentDrawing = False

    On Error GoTo ErrorHandler

    ' Open the drawing (not read-only, so we can modify it)
    Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)

    If swDraw Is Nothing Then
        Exit Function
    End If

    Set swDrawModel = swDraw

    ' Get the first view to apply positioning
    Set swView1 = GetDrawingView(swDraw, 1)

    If swView1 Is Nothing Then
        swApp.CloseDoc drawingPath
        Exit Function
    End If

    ' Get the second view if it exists
    Set swView2 = GetDrawingView(swDraw, 2)

    ' Get the third view if it exists
    Set swView3 = GetDrawingView(swDraw, 3)

    ' Apply positioning and scaling
    PositionAndScaleViews swDraw, swDrawModel, swView1, swView2, swView3

    ' Save the drawing (to suppress "save modified" dialog)
    swDrawModel.Save

    ' Get PDF output path
    fileName = Left(drawingPath, InStrRev(drawingPath, ".") - 1)
    fileName = Right(fileName, Len(fileName) - InStrRev(fileName, "\"))
    pdfPath = pdfFolder & fileName & ".PDF"

    ' Export to PDF with all sheets selected
    saveError = swDrawModel.SaveAs3(pdfPath, 0, swSaveAsOptions_e.swSaveAsOptions_SaveReferenced)

    ' If SaveAs3 fails, try SaveAs2
    If saveError <> 0 Then
        swDrawModel.SaveAs2 pdfPath, swSaveAsCurrentVersion, False, True
    End If

    ' Close the drawing without saving again
    swApp.CloseDoc drawingPath

    ProcessComponentDrawing = True
    Exit Function

ErrorHandler:
    If Not swDraw Is Nothing Then
        swApp.CloseDoc drawingPath
    End If
    ProcessComponentDrawing = False

End Function


'============================================================
' CHECK IF COMPONENT IS SUPPRESSED IN MAIN ASSEMBLY
'============================================================

Function IsComponentSuppressed(swAssembly As SldWorks.ModelDoc2, componentPath As String) As Boolean

    Dim swConfig As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    Dim swComp As Object
    Dim i As Long
    Dim vComps As Variant
    Dim compPath As String

    IsComponentSuppressed = False

    On Error Resume Next

    Set swConfig = swAssembly.GetActiveConfiguration
    If swConfig Is Nothing Then
        Exit Function
    End If

    Set swRootComp = swConfig.GetRootComponent3(True)
    If swRootComp Is Nothing Then
        Exit Function
    End If

    ' Get all direct children
    vComps = swRootComp.GetChildren

    If Not IsEmpty(vComps) Then
        For i = 0 To UBound(vComps)
            Set swComp = vComps(i)
            compPath = swComp.GetPathName

            ' Check if this component's path matches
            If LCase(compPath) = LCase(componentPath) Then
                ' Found it - check if suppressed
                If swComp.IsSuppressed Then
                    IsComponentSuppressed = True
                End If
                Exit Function
            End If
        Next i
    End If

    On Error GoTo 0

End Function


'============================================================
' GET SPECIFIC DRAWING VIEW (1-based index)
'============================================================

Function GetDrawingView(swDraw As SldWorks.DrawingDoc, viewIndex As Long) As SldWorks.View

    Dim swView As SldWorks.View
    Dim i As Long

    Set swView = swDraw.GetFirstView

    If swView Is Nothing Then
        Set GetDrawingView = Nothing
        Exit Function
    End If

    ' Skip sheet view (index 0), get to view 1
    Set swView = swView.GetNextView

    ' Now get to the requested view
    For i = 1 To viewIndex - 1
        If swView Is Nothing Then
            Set GetDrawingView = Nothing
            Exit Function
        End If
        Set swView = swView.GetNextView
    Next i

    Set GetDrawingView = swView

End Function


'============================================================
' POSITION AND SCALE VIEWS
'
' Applies the positioning and scaling logic to the drawing.
' Adapted from the provided positioning code.
'============================================================

Sub PositionAndScaleViews(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, _
                          swView1 As SldWorks.View, swView2 As SldWorks.View, swView3 As SldWorks.View)

    Dim modelPath As String
    Dim fileName As String
    Dim baseName As String
    Dim fileParts() As String
    Dim partStyle As String
    Dim lengthPart As String
    Dim lengthPrefix As String
    Dim lengthPartIndex As Long
    Dim heightPartIndex As Long
    Dim lengthText As String
    Dim heightText As String
    Dim overallLength As Double
    Dim overallHeight As Double
    Dim scaleNumerator As Double
    Dim scaleDenominator As Double

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim swSheet As SldWorks.Sheet

    Dim gapMeters As Double
    Dim outline1 As Variant
    Dim outline2 As Variant
    Dim outline3 As Variant
    Dim width1 As Double
    Dim height1 As Double
    Dim width2 As Double
    Dim height2 As Double
    Dim width3 As Double
    Dim height3 As Double
    Dim totalWidth As Double
    Dim totalHeight As Double
    Dim startX As Double
    Dim startY As Double
    Dim horizontalOffsetMeters As Double
    Dim view1Left As Double
    Dim view1Bottom As Double
    Dim view2Left As Double
    Dim view2Bottom As Double
    Dim view3Left As Double
    Dim view3Bottom As Double

    Dim curPos1 As Variant
    Dim curPos2 As Variant
    Dim curPos3 As Variant
    Dim offset1X As Double
    Dim offset1Y As Double
    Dim offset2X As Double
    Dim offset2Y As Double
    Dim offset3X As Double
    Dim offset3Y As Double
    Dim target1X As Double
    Dim target1Y As Double
    Dim target2X As Double
    Dim target2Y As Double
    Dim target3X As Double
    Dim target3Y As Double
    Dim pos1(1) As Double
    Dim pos2(1) As Double
    Dim pos3(1) As Double

    Dim break1OK As Boolean
    Dim break2OK As Boolean
    Dim break3OK As Boolean
    Dim autoArrangeOK As Boolean
    Dim dimSelectedCount As Long
    Dim dimErrorText As String

    On Error Resume Next

    ' Get referenced model path
    modelPath = swView1.GetReferencedModelName
    If modelPath = "" Then
        Exit Sub
    End If

    ' Extract filename
    fileName = Mid(modelPath, InStrRev(modelPath, "\") + 1)
    If InStrRev(fileName, ".") > 0 Then
        baseName = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        baseName = fileName
    End If

    ' Extract dimensions flexibly from filename
    ' Handles various naming conventions: 3-C6000-2400-21, 2400-21-CG, etc.
    If Not ExtractDimensions(baseName, overallLength, overallHeight) Then
        Exit Sub
    End If

    ' Determine scale
    If overallHeight >= 36 Or overallLength >= 96 Then
        scaleNumerator = 1
        scaleDenominator = 12
    ElseIf overallHeight >= 24 Or overallLength >= 72 Then
        scaleNumerator = 1
        scaleDenominator = 10
    ElseIf overallHeight >= 18 Or overallLength >= 60 Then
        scaleNumerator = 1
        scaleDenominator = 8
    ElseIf overallHeight >= 12 Or overallLength >= 48 Then
        scaleNumerator = 1
        scaleDenominator = 6
    ElseIf overallHeight >= 6 Or overallLength >= 24 Then
        scaleNumerator = 1
        scaleDenominator = 4
    Else
        scaleNumerator = 1
        scaleDenominator = 2
    End If

    ' Break alignment
    break1OK = BreakViewAlignment(swDraw, swDrawModel, swView1)
    If Not swView2 Is Nothing Then
        break2OK = BreakViewAlignment(swDraw, swDrawModel, swView2)
    End If
    If Not swView3 Is Nothing Then
        break3OK = BreakViewAlignment(swDraw, swDrawModel, swView3)
    End If

    ' Apply scale
    swView1.UseSheetScale = False
    swView1.ScaleDecimal = scaleNumerator / scaleDenominator
    If Not swView2 Is Nothing Then
        swView2.UseSheetScale = False
        swView2.ScaleDecimal = scaleNumerator / scaleDenominator
    End If
    If Not swView3 Is Nothing Then
        swView3.UseSheetScale = False
        swView3.ScaleDecimal = scaleNumerator / scaleDenominator
    End If

    ' Rebuild
    swDraw.ForceRebuild3 False

    ' Get sheet size
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Sub
    swSheet.GetSize sheetWidth, sheetHeight

    ' Gap between views
    gapMeters = 0.0508

    ' Get outlines
    outline1 = swView1.GetOutline
    width1 = outline1(2) - outline1(0)
    height1 = outline1(3) - outline1(1)

    If Not swView2 Is Nothing Then
        outline2 = swView2.GetOutline
        width2 = outline2(2) - outline2(0)
        height2 = outline2(3) - outline2(1)
    Else
        width2 = 0#
        height2 = 0#
    End If

    If Not swView3 Is Nothing Then
        outline3 = swView3.GetOutline
        width3 = outline3(2) - outline3(0)
        height3 = outline3(3) - outline3(1)
    Else
        width3 = 0#
        height3 = 0#
    End If

    ' Calculate group size
    ' Layout: View1 and View2 side-by-side at top, View3 below View1
    Dim view1View2Width As Double
    Dim view1View2Height As Double
    Dim totalWidthWithView3 As Double
    Dim totalHeightWithView3 As Double

    ' Width and height of View1 + View2 pair
    If Not swView2 Is Nothing Then
        view1View2Width = width1 + gapMeters + width2
        view1View2Height = IIf(height1 >= height2, height1, height2)
    Else
        view1View2Width = width1
        view1View2Height = height1
    End If

    ' Calculate total with View3 (below View1)
    If Not swView3 Is Nothing Then
        ' Width: max of (View1+View2 pair) and View3, since View3 is below View1
        totalWidthWithView3 = IIf(view1View2Width >= width3, view1View2Width, width3)
        ' Height: View1+View2 pair height + gap + View3 height
        ' But View3 is only as wide as View1, so overall width is the pair width
        totalWidth = view1View2Width
        totalHeight = view1View2Height + gapMeters + height3
    Else
        totalWidth = view1View2Width
        totalHeight = view1View2Height
    End If

    ' Center on sheet
    startX = (sheetWidth - totalWidth) / 2#
    startY = (sheetHeight - totalHeight) / 2#

    ' Move 1 inch left
    horizontalOffsetMeters = 1# * 0.0254
    startX = startX - horizontalOffsetMeters

    ' View 1 position
    view1Left = startX
    If Not swView3 Is Nothing Then
        ' With 3 views: View1 at bottom of pair (View3 goes below)
        view1Bottom = startY + view1View2Height + gapMeters
    Else
        ' With 1-2 views: View1 centered vertically
        If Not swView2 Is Nothing Then
            view1Bottom = startY + (view1View2Height - height1) / 2#
        Else
            view1Bottom = startY
        End If
    End If

    ' View 2 position (top-right, when it exists)
    If Not swView2 Is Nothing Then
        view2Left = startX + width1 + gapMeters
        If Not swView3 Is Nothing Then
            ' With 3 views: View2 aligned at top of pair
            view2Bottom = startY + (view1View2Height - height2) / 2#
        Else
            ' With 2 views: View2 vertically centered with View1
            view2Bottom = startY + (view1View2Height - height2) / 2#
        End If
    End If

    ' View 3 position (below View1, only when it exists)
    If Not swView3 Is Nothing Then
        view3Left = startX
        view3Bottom = startY
    End If

    ' Get current positions
    curPos1 = swView1.Position
    If Not swView2 Is Nothing Then
        curPos2 = swView2.Position
    End If
    If Not swView3 Is Nothing Then
        curPos3 = swView3.Position
    End If

    ' Calculate offsets
    offset1X = curPos1(0) - outline1(0)
    offset1Y = curPos1(1) - outline1(1)
    If Not swView2 Is Nothing Then
        offset2X = curPos2(0) - outline2(0)
        offset2Y = curPos2(1) - outline2(1)
    End If
    If Not swView3 Is Nothing Then
        offset3X = curPos3(0) - outline3(0)
        offset3Y = curPos3(1) - outline3(1)
    End If

    ' Calculate targets
    target1X = view1Left + offset1X
    target1Y = view1Bottom + offset1Y
    If Not swView2 Is Nothing Then
        target2X = view2Left + offset2X
        target2Y = view2Bottom + offset2Y
    End If
    If Not swView3 Is Nothing Then
        target3X = view3Left + offset3X
        target3Y = view3Bottom + offset3Y
    End If

    ' Set positions
    pos1(0) = target1X
    pos1(1) = target1Y
    swView1.Position = pos1

    If Not swView2 Is Nothing Then
        pos2(0) = target2X
        pos2(1) = target2Y
        swView2.Position = pos2
    End If

    If Not swView3 Is Nothing Then
        pos3(0) = target3X
        pos3(1) = target3Y
        swView3.Position = pos3
    End If

    ' Final rebuild
    swDraw.ForceRebuild3 False

    ' Auto-arrange dimensions
    autoArrangeOK = AutoArrangeViewDimensions(swDrawModel, swView1, swView2, swView3, dimSelectedCount, dimErrorText)

    ' Final rebuild
    swDraw.ForceRebuild3 False

    On Error GoTo 0

End Sub


'============================================================
' AUTO ARRANGE DIMENSIONS (from provided code)
'============================================================

Function AutoArrangeViewDimensions( _
    swDrawModel As SldWorks.ModelDoc2, _
    swView1 As SldWorks.View, _
    swView2 As SldWorks.View, _
    swView3 As SldWorks.View, _
    ByRef selectedCountOut As Long, _
    ByRef errorTextOut As String) As Boolean

    Dim swAnn As SldWorks.Annotation
    Dim swDispDim As SldWorks.DisplayDimension
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swSelData As SldWorks.SelectData
    Dim selectedCount As Long
    Dim errorText As String
    Dim selOK As Boolean
    Dim arrangeOK As Boolean
    Dim jogOK As Boolean

    selectedCount = 0
    errorText = ""
    arrangeOK = False

    Set swSelMgr = swDrawModel.SelectionManager
    Set swSelData = swSelMgr.CreateSelectData

    On Error Resume Next
    swDrawModel.ClearSelection2 True
    On Error GoTo 0

    ' VIEW 1
    If Not swView1 Is Nothing Then
        Set swAnn = swView1.GetFirstAnnotation3
        Do While Not swAnn Is Nothing
            If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
                Set swDispDim = swAnn.GetSpecificAnnotation
                If Not swDispDim Is Nothing Then
                    ' Ordinate dimension
                    If IsOrdinateDimension(swDispDim) Then
                        On Error Resume Next
                        jogOK = swDispDim.AutoJogOrdinate()
                        On Error GoTo 0
                    End If
                    ' Select for auto arrange
                    selOK = False
                    On Error Resume Next
                    selOK = swAnn.Select3(True, swSelData)
                    On Error GoTo 0
                    If selOK Then
                        selectedCount = selectedCount + 1
                    End If
                End If
            End If
            Set swAnn = swAnn.GetNext3
        Loop
    End If

    ' VIEW 2
    If Not swView2 Is Nothing Then
        Set swAnn = swView2.GetFirstAnnotation3
        Do While Not swAnn Is Nothing
            If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
                Set swDispDim = swAnn.GetSpecificAnnotation
                If Not swDispDim Is Nothing Then
                    ' Ordinate dimension
                    If IsOrdinateDimension(swDispDim) Then
                        On Error Resume Next
                        jogOK = swDispDim.AutoJogOrdinate()
                        On Error GoTo 0
                    End If
                    ' Select for auto arrange
                    selOK = False
                    On Error Resume Next
                    selOK = swAnn.Select3(True, swSelData)
                    On Error GoTo 0
                    If selOK Then
                        selectedCount = selectedCount + 1
                    End If
                End If
            End If
            Set swAnn = swAnn.GetNext3
        Loop
    End If

    ' VIEW 3
    If Not swView3 Is Nothing Then
        Set swAnn = swView3.GetFirstAnnotation3
        Do While Not swAnn Is Nothing
            If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
                Set swDispDim = swAnn.GetSpecificAnnotation
                If Not swDispDim Is Nothing Then
                    ' Ordinate dimension
                    If IsOrdinateDimension(swDispDim) Then
                        On Error Resume Next
                        jogOK = swDispDim.AutoJogOrdinate()
                        On Error GoTo 0
                    End If
                    ' Select for auto arrange
                    selOK = False
                    On Error Resume Next
                    selOK = swAnn.Select3(True, swSelData)
                    On Error GoTo 0
                    If selOK Then
                        selectedCount = selectedCount + 1
                    End If
                End If
            End If
            Set swAnn = swAnn.GetNext3
        Loop
    End If

    ' NORMAL DIMENSION AUTO ARRANGE
    If selectedCount > 0 Then
        On Error Resume Next
        swDrawModel.Extension.AlignDimensions swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001
        arrangeOK = True
        On Error GoTo 0
    End If

    ' REBUILD
    swDrawModel.ForceRebuild3 False

    selectedCountOut = selectedCount
    errorTextOut = errorText

    On Error Resume Next
    swDrawModel.ClearSelection2 True
    On Error GoTo 0

    If arrangeOK Or selectedCount > 0 Then
        AutoArrangeViewDimensions = True
    Else
        AutoArrangeViewDimensions = False
    End If

End Function


'============================================================
' CHECK IF DIMENSION IS ORDINATE
'============================================================

Function IsOrdinateDimension(swDispDim As SldWorks.DisplayDimension) As Boolean

    Dim dimType As Long

    On Error GoTo NotOrdinate
    dimType = swDispDim.Type2
    If dimType = swOrdinateDimension Then
        IsOrdinateDimension = True
    Else
        IsOrdinateDimension = False
    End If
    Exit Function

NotOrdinate:
    IsOrdinateDimension = False

End Function


'============================================================
' BREAK VIEW ALIGNMENT
'============================================================

Function BreakViewAlignment( _
    swDraw As SldWorks.DrawingDoc, _
    swDrawModel As SldWorks.ModelDoc2, _
    swView As SldWorks.View) As Boolean

    Dim selected As Boolean

    On Error GoTo ErrHandler

    If swView Is Nothing Then
        BreakViewAlignment = False
        Exit Function
    End If

    swDrawModel.ClearSelection2 True

    selected = swDrawModel.Extension.SelectByID2( _
                    swView.Name, _
                    "DRAWINGVIEW", _
                    0, 0, 0, False, 0, Nothing, 0)

    If Not selected Then
        BreakViewAlignment = False
        Exit Function
    End If

    swDraw.BreakAlignment
    swDrawModel.ClearSelection2 True
    BreakViewAlignment = True
    Exit Function

ErrHandler:
    swDrawModel.ClearSelection2 True
    BreakViewAlignment = False

End Function


'============================================================
' CHECK IF STRING IS NUMERIC
'============================================================

Function IsNumeric(str As String) As Boolean

    On Error GoTo NotNumeric
    Dim x As Double
    x = CDbl(str)
    IsNumeric = True
    Exit Function

NotNumeric:
    IsNumeric = False

End Function


'============================================================
' EXTRACT DIMENSIONS FROM FLEXIBLE FILENAMES
'
' Handles all naming patterns:
'   - 3-C6000-2400-21 (with C/S/SL/SC/SR prefix)
'   - 2400-21-CG (gusset, height only)
'   - 40-3-C8800-S6700-2300-21 (complex)
'
' Returns: True if at least one dimension found
'          overallLength and overallHeight in inches (0 if not found)
'============================================================

Function ExtractDimensions(baseName As String, ByRef foundLength As Double, ByRef foundHeight As Double) As Boolean

    Dim fileParts() As String
    Dim i As Long
    Dim part As String
    Dim partUpper As String
    Dim prefix As String
    Dim numText As String
    Dim isGusset As Boolean
    Dim extractedDimensions As Variant
    Dim dimCount As Long
    Dim length As Double
    Dim height As Double
    Dim tempValue As Double
    Dim numericValues() As Double
    Dim numericCount As Long

    foundLength = 0#
    foundHeight = 0#
    ExtractDimensions = False
     
    ReDim numericValues(20)
    numericCount = 0

    fileParts = Split(baseName, "-")

    ' Check if it's a gusset (ends with -CG or -MG)
    isGusset = False
    If UBound(fileParts) >= 0 Then
        partUpper = UCase(fileParts(UBound(fileParts)))
        If partUpper = "CG" Or partUpper = "MG" Then
            isGusset = True
        End If
    End If

    ' Extract all numeric values and prefixed dimensions
    For i = 0 To UBound(fileParts)
        part = fileParts(i)
        partUpper = UCase(part)

        ' Check for prefixed dimensions (C, S, SL, SC, SR)
        If (Left(partUpper, 2) = "SL" Or Left(partUpper, 2) = "SC" Or Left(partUpper, 2) = "SR") Then
            prefix = Left(partUpper, 2)
            numText = Mid(part, 3)
            If IsNumeric(numText) Then
                numericValues(numericCount) = CDbl(numText) / 100#
                numericCount = numericCount + 1
            End If
        ElseIf (Left(partUpper, 1) = "C" Or Left(partUpper, 1) = "S") Then
            ' Single letter prefix (C or S)
            prefix = Left(partUpper, 1)
            numText = Mid(part, 2)
            If IsNumeric(numText) Then
                numericValues(numericCount) = CDbl(numText) / 100#
                numericCount = numericCount + 1
            End If
        ElseIf IsNumeric(part) Then
            ' Standalone numeric value (no prefix)
            numericValues(numericCount) = CDbl(part) / 100#
            numericCount = numericCount + 1
        End If
    Next i

    ' No dimensions found
    If numericCount = 0 Then
        Exit Function
    End If

    ' For gussets (e.g., 2400-21-CG): first numeric value is height
    If isGusset Then
        foundHeight = numericValues(0)
        ExtractDimensions = True
        Exit Function
    End If

    ' For other components with multiple dimensions
    If numericCount >= 2 Then
        ' Use first as length, second as height (or last two if more)
        foundLength = numericValues(0)
        foundHeight = numericValues(1)
    ElseIf numericCount = 1 Then
        ' Only one dimension: use for height
        foundHeight = numericValues(0)
    End If

    ExtractDimensions = (foundHeight > 0 Or foundLength > 0)

End Function


'============================================================
' CHECK IF PDF FOLDER HAS EXISTING PDFs
'============================================================

Function HasExistingPDFs(folderPath As String) As Boolean

    Dim fileName As String

    HasExistingPDFs = False

    ' Look for any .PDF files in the folder
    fileName = Dir(folderPath & "*.PDF")

    If fileName <> "" Then
        HasExistingPDFs = True
    End If

End Function


'============================================================
' PROMPT USER TO OVERWRITE PDFs
' Returns True if user chooses OVERWRITE, False if CANCEL
'============================================================

Function PromptOverwritePDFs() As Boolean

    Dim userChoice As VbMsgBoxResult

    PromptOverwritePDFs = False

    userChoice = MsgBox("PDFs already exist in the PDF folder." & vbCrLf & vbCrLf & _
                       "Would you like to overwrite them?" & vbCrLf & vbCrLf & _
                       "Click YES to overwrite and continue." & vbCrLf & _
                       "Click NO to cancel this operation.", _
                       vbQuestion + vbYesNo, _
                       "Existing PDFs Found")

    If userChoice = vbYes Then
        PromptOverwritePDFs = True
    End If

End Function


'============================================================
' CLEAR ALL PDFs FROM FOLDER
'============================================================

Sub ClearPDFFolder(folderPath As String)

    Dim fileName As String
    Dim filePath As String

    On Error Resume Next

    ' Delete all .PDF files in the folder
    fileName = Dir(folderPath & "*.PDF")

    Do While fileName <> ""
        filePath = folderPath & fileName
        Kill filePath
        fileName = Dir
    Loop

    On Error GoTo 0

End Sub
