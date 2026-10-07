Option Explicit

'============================================================
' >>> DEPLOY TARGET <<<
' File:   Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\pdfassemblies.swp
' Module: pdfassemblies1
' Action: Tools > Macro > Edit on pdfassemblies.swp, select all code in
'         module pdfassemblies1, delete, paste this file's content in, save.
' Changed: ProcessMainAssemblyViews - IMPORTANT FIX. The live version uses
'         old fixed SetPosition coordinates for Top/Front/Isometric views,
'         which is the exact bug that made views visually shift at
'         different scales. Replaced with outline-based centering plus an
'         anchor-to-outline offset correction (View.Position is an
'         internal anchor point, not the bounding-box corner, and that
'         gap scales with view scale). New measured center coordinates:
'         Top (5.31, 6.27), Front (5.31, 2.78), Isometric (11.29, 8.51) in.
' Changed 2: Added swApp.UserControlBackground = True/False around the
'         drawing processing loop so drawings open/close silently in the
'         background (no window flashing/stealing focus) - faster and
'         less disruptive.
' Changed 3: Output folder renamed from PDF\ to PDF_Assemblies\ so
'         assembly and component PDFs no longer land in the same folder.
' Changed 4: Removed the automatic OpenLogFile calls (which launched
'         Notepad) at completion and on error/cancel paths. The log
'         file is still written to LOGS\ for later troubleshooting, it
'         just no longer pops open and interrupts an automated packngo
'         run. OpenLogFile sub is still present if you want to call it
'         manually.
' Changed 5: ProcessMainAssemblyViews now auto-arranges dimensions
'         across all three views (Top/Front/Isometric) at the end,
'         using the same dimension auto-arrange technique already used
'         for component drawings (new AutoArrangeMainAssemblyDimensions
'         function, modeled on pdfcomponents.bas's
'         AutoArrangeViewDimensions). Previously the main assembly sheet
'         had no automatic dimension arrangement at all.
' Changed 6: Sub-assembly drawings now get dimension auto-arranging on
'         BOTH sheets/views that were previously missed:
'         - Sheet1 (exploded view) previously had NO dimension
'           auto-arranging at all - now calls
'           AutoArrangeMainAssemblyDimensions on the exploded view.
'         - Sheet2 previously only auto-arranged the Top view's
'           dimensions (via AutoAdjustTopViewDimensions) and skipped the
'           Isometric view entirely - now calls
'           AutoArrangeMainAssemblyDimensions with both views so
'           Isometric view dimensions get arranged too.
'         AutoAdjustTopViewDimensions is no longer called but left in
'         place in case it's still needed for reference.
' Changed 7: Added ESC-key cancel support. New CancelRequested function
'         polls for the ESC key between each sub-assembly drawing;
'         pressing it prompts to confirm, then stops cleanly after the
'         current drawing and reports how many were completed before
'         cancelling.
'============================================================

'============================================================
' PACK AND GO - SUB-ASSEMBLY PDF EXPORT
'
' Exports sub-assembly drawings to PDF with automatic:
' - Page 1: Exploded view positioning, scaling, and balloon arrangement
' - Page 2: Isometric view (top-right) and top view (centered)
'           with auto-scaling and dimension adjustment
'
' Run from the main assembly after Pack and Go completes.
' Processes all sub-assemblies found in the main assembly.
'============================================================

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2

Public destFolder As String
Public pdfFolder As String
Public logFilePath As String

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

Sub RunPDFAssemblyDrawings()

     Dim assemblyPath As String
     Dim subAssemblyDrawings() As String
     Dim drawingCount As Long
     Dim i As Long
     Dim successCount As Long
     Dim failureCount As Long
     Dim logDir As String

     Set swApp = Application.SldWorks
     Set swModel = swApp.ActiveDoc

     If swModel Is Nothing Then
         MsgBox "No document is open.", vbExclamation
         Exit Sub
     End If

     If swModel.GetType <> swDocASSEMBLY Then
         MsgBox "Please run this from the main assembly (from PACK_TEST folder).", vbExclamation
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
     pdfFolder = destFolder & "PDF_Assemblies\"
     logDir = destFolder & "LOGS\"

     ' Create LOGS folder if needed
     If Dir(logDir, vbDirectory) = "" Then
         MkDir logDir
     End If

     ' Setup log file
     logFilePath = logDir & "SubAssembly_PDF_" & Format(Now, "yyyy-mm-dd_HH-mm-ss") & ".txt"

     ' Log startup
     LogWrite "=================================================="
     LogWrite "SUB-ASSEMBLY PDF EXPORT LOG"
     LogWrite "=================================================="
     LogWrite "Start Time: " & Format(Now, "yyyy-mm-dd HH:mm:ss")
     LogWrite "Main Assembly: " & assemblyPath
     LogWrite "Destination Folder: " & destFolder
     LogWrite ""

     ' Verify this is a Pack and Go folder
     If InStr(1, destFolder, "PACK_TEST", vbTextCompare) = 0 Then
         LogWrite "ERROR: Not a PACK_TEST folder!"
         MsgBox "This assembly does not appear to be from a Pack and Go output folder." & vbCrLf & vbCrLf & _
                "Please open the assembly from:" & vbCrLf & _
                "Z:\...\PACK_TEST\<part_number>\", vbExclamation
         Exit Sub
     End If

     ' Get all sub-assembly drawing files
     LogWrite "Searching for sub-assembly drawings..."
     drawingCount = GetSubAssemblyDrawings(swModel, subAssemblyDrawings)

     If drawingCount = 0 Then
         LogWrite "ERROR: No sub-assembly drawings found!"
         MsgBox "No sub-assembly drawings found in:" & vbCrLf & destFolder, vbExclamation
         Exit Sub
     End If

     LogWrite "Found " & drawingCount & " sub-assembly drawings:"
     For i = 0 To drawingCount - 1
         LogWrite "  " & (i + 1) & ". " & subAssemblyDrawings(i)
     Next i
     LogWrite ""

     ' Create PDF folder if needed
     If Dir(pdfFolder, vbDirectory) = "" Then
         MkDir pdfFolder
         LogWrite "Created PDF folder"
     Else
         ' PDF folder exists - check if it has PDFs and prompt user
         LogWrite "PDF folder already exists"
         If HasExistingPDFs(pdfFolder) Then
             LogWrite "Existing PDFs found - prompting user"
             If Not PromptOverwritePDFs() Then
                 ' User chose cancel
                 LogWrite "User cancelled - exiting"
                 Exit Sub
             End If
             ' User chose overwrite - clear the folder
             LogWrite "User selected overwrite - clearing folder"
             ClearPDFFolder pdfFolder
         End If
     End If

     LogWrite ""
     MsgBox "Found " & drawingCount & " sub-assembly drawings." & vbCrLf & _
            "Processing sub-assemblies..." & vbCrLf & vbCrLf & _
            "Folder: " & destFolder & vbCrLf & vbCrLf & _
            "Press ESC at any time to cancel after the current drawing finishes." & vbCrLf & vbCrLf & _
            "This may take a few minutes.", vbInformation, "Starting Sub-Assembly PDF Export"

     LogWrite "Starting PDF export processing..."
     LogWrite ""

     ' Open/close drawings silently in the background - no window
     ' flashing or stealing focus while processing, which is both
     ' faster and lets you keep working in the foreground.
     swApp.UserControlBackground = True

     ' Process each sub-assembly drawing
     successCount = 0
     failureCount = 0
     Dim wasCancelled As Boolean
     wasCancelled = False

     For i = 0 To drawingCount - 1
         If CancelRequested() Then
             LogWrite "User cancelled the run - stopping."
             wasCancelled = True
             Exit For
         End If

         LogWrite "[" & (i + 1) & "/" & drawingCount & "] Processing: " & subAssemblyDrawings(i)
         If ProcessAssemblyDrawing(subAssemblyDrawings(i)) Then
             successCount = successCount + 1
             LogWrite "  SUCCESS"
         Else
             failureCount = failureCount + 1
             LogWrite "  FAILED"
         End If
     Next i

     swApp.UserControlBackground = False

     LogWrite ""
     LogWrite "=================================================="
     LogWrite IIf(wasCancelled, "PROCESSING CANCELLED BY USER", "PROCESSING COMPLETE")
     LogWrite "=================================================="
     LogWrite "Successfully processed: " & successCount
     LogWrite "Failed: " & failureCount
     LogWrite "End Time: " & Format(Now, "yyyy-mm-dd HH:mm:ss")
     LogWrite "Log file: " & logFilePath

     If wasCancelled Then
         MsgBox "Sub-assembly PDF export cancelled by user." & vbCrLf & vbCrLf & _
                "Successfully processed before cancelling: " & successCount & " of " & drawingCount & vbCrLf & vbCrLf & _
                "Log file: " & logFilePath, vbExclamation, "Cancelled"
     Else
         MsgBox "Sub-assembly PDF export complete!" & vbCrLf & vbCrLf & _
                "Successfully processed: " & successCount & vbCrLf & _
                "Failed: " & failureCount & vbCrLf & vbCrLf & _
                "Log file: " & logFilePath, vbInformation, "Complete"
     End If

End Sub


'============================================================
' GET ALL SUB-ASSEMBLY DRAWINGS
'
' Walks through active assembly components and finds all
' sub-assembly (.SLDDRW) files
'============================================================

Function GetSubAssemblyDrawings(swAssembly As SldWorks.ModelDoc2, ByRef drawingPaths() As String) As Long

     Dim swConfig As SldWorks.Configuration
     Dim swRootComp As SldWorks.Component2
     Dim drawingList As New collection
     Dim i As Long
     Dim fileName As String

     GetSubAssemblyDrawings = 0

     On Error Resume Next

     Set swConfig = swAssembly.GetActiveConfiguration
     If swConfig Is Nothing Then
         Exit Function
     End If

     Set swRootComp = swConfig.GetRootComponent3(True)
     If swRootComp Is Nothing Then
         Exit Function
     End If

     ' Recursively collect sub-assembly drawings
     CollectSubAssemblyDrawings swRootComp, drawingList

     ' Convert collection to array
     If drawingList.Count > 0 Then
         ReDim drawingPaths(drawingList.Count - 1)
         For i = 1 To drawingList.Count
             drawingPaths(i - 1) = drawingList.Item(i)
         Next i
         GetSubAssemblyDrawings = drawingList.Count
     End If

     On Error GoTo 0

End Function


'============================================================
' RECURSIVELY COLLECT SUB-ASSEMBLY DRAWING PATHS
'============================================================

Sub CollectSubAssemblyDrawings(swComp As SldWorks.Component2, ByRef collection As collection)

     Dim children() As Object
     Dim i As Long
     Dim j As Long
     Dim childComp As SldWorks.Component2
     Dim modelDoc As SldWorks.ModelDoc2
     Dim modelPath As String
     Dim drawingPath As String
     Dim baseName As String
     Dim isDuplicate As Boolean

     On Error Resume Next

     ' Get component model document
     Set modelDoc = swComp.GetModelDoc2

     If Not modelDoc Is Nothing Then
         ' Check if this component is an assembly
         If modelDoc.GetType = swDocASSEMBLY Then
             ' Try to find matching drawing
             modelPath = modelDoc.GetPathName
             If modelPath <> "" Then
                 ' Get base name without extension
                 baseName = Mid(modelPath, InStrRev(modelPath, "\") + 1)
                 If InStrRev(baseName, ".") > 0 Then
                     baseName = Left(baseName, InStrRev(baseName, ".") - 1)
                 End If

                 ' Look for matching .SLDDRW
                 drawingPath = destFolder & baseName & ".SLDDRW"
                 If Dir(drawingPath) <> "" Then
                     ' Check if this drawing path already exists in collection
                     isDuplicate = False
                     For j = 1 To collection.Count
                         If LCase(collection.Item(j)) = LCase(drawingPath) Then
                             isDuplicate = True
                             Exit For
                         End If
                     Next j

                     ' Only add if not a duplicate
                     If Not isDuplicate Then
                         collection.Add drawingPath
                     Else
                         LogWrite "  (Duplicate sub-assembly skipped: " & baseName & ")"
                     End If
                 End If
             End If
         End If
     End If

     ' Get children (only active/non-suppressed)
     children = swComp.GetChildren

     If Not IsEmpty(children) Then
         For i = LBound(children) To UBound(children)
             If Not children(i) Is Nothing Then
                 Set childComp = children(i)
                 ' Recursively process children
                 CollectSubAssemblyDrawings childComp, collection
             End If
         Next i
     End If

     On Error GoTo 0

End Sub


'============================================================
' PROCESS SINGLE ASSEMBLY DRAWING
'============================================================

Function ProcessAssemblyDrawing(drawingPath As String) As Boolean

     Dim swDraw As SldWorks.DrawingDoc
     Dim swDrawModel As SldWorks.ModelDoc2
     Dim baseName As String
     Dim isMainAssembly As Boolean

     ProcessAssemblyDrawing = False

     On Error GoTo ErrorHandler

     ' Open the drawing (not read-only, so we can modify it)
     LogWrite "  Opening drawing..."
     Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)

     If swDraw Is Nothing Then
         LogWrite "  ERROR: Failed to open drawing"
         Exit Function
     End If

     LogWrite "  Drawing opened successfully"
     Set swDrawModel = swDraw
      
     ' Get basename to determine if it's main or sub-assembly
     baseName = GetDrawingBaseName(drawingPath)
     isMainAssembly = IsMainAssemblyDrawing(baseName)
     
     If isMainAssembly Then
         LogWrite "  Detected: MAIN ASSEMBLY (using main assembly processing)"
         ProcessAssemblyDrawing = ProcessMainAssemblyDrawing(swDraw, swDrawModel, drawingPath)
     Else
         LogWrite "  Detected: SUB-ASSEMBLY (using sub-assembly processing)"
         ProcessAssemblyDrawing = ProcessSubAssemblyDrawing(swDraw, swDrawModel, drawingPath)
     End If
      
     Exit Function

ErrorHandler:
     LogWrite "  ERROR: " & Err.Description & " (Code: " & Err.Number & ")"
     If Not swDraw Is Nothing Then
         swApp.CloseDoc drawingPath
     End If
     ProcessAssemblyDrawing = False

End Function


'============================================================
' PROCESS SUB-ASSEMBLY DRAWING (Original Logic)
'============================================================

Function ProcessSubAssemblyDrawing(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, drawingPath As String) As Boolean

     Dim swExplodedView As SldWorks.View
     Dim swIsometricView As SldWorks.View
     Dim swTopView As SldWorks.View
     Dim pdfPath As String
     Dim fileName As String
     Dim saveError As Long

     ProcessSubAssemblyDrawing = False

     On Error GoTo ErrorHandler
      
     On Error Resume Next
       
     ' Process Sheet1 (Exploded View)
     LogWrite "  Activating Sheet1..."
     If swDrawModel.ActivateSheet("Sheet1") Then
         swDrawModel.ForceRebuild3 False
         LogWrite "    Sheet1 activated"
            
         ' Get exploded view
         LogWrite "    Looking for exploded view..."
         Set swExplodedView = GetFirstModelView(swDraw)
         If Not swExplodedView Is Nothing Then
             LogWrite "    Found exploded view, positioning..."
             ProcessSheet1Exploded swDraw, swDrawModel, swExplodedView
             LogWrite "    Exploded view positioned"
         Else
             LogWrite "    No exploded view found"
         End If
     Else
         LogWrite "    Sheet1 not found, skipping"
     End If
       
     ' Process Sheet2 (Isometric + Top View)
     LogWrite "  Activating Sheet2..."
     If swDrawModel.ActivateSheet("Sheet2") Then
         swDrawModel.ForceRebuild3 False
         LogWrite "    Sheet2 activated"
           
         ' Get isometric and top views
         LogWrite "    Looking for views..."
         Set swIsometricView = GetNamedView(swDraw, "Isometric")
         Set swTopView = GetNamedView(swDraw, "Top")
          
         If Not swIsometricView Is Nothing Or Not swTopView Is Nothing Then
             LogWrite "    Found assembly views, positioning..."
             ProcessSheet2Assembly swDraw, swDrawModel, swIsometricView, swTopView
             LogWrite "    Assembly views positioned"
         Else
             LogWrite "    No assembly views found"
         End If
     Else
         LogWrite "    Sheet2 not found, skipping"
     End If

     On Error GoTo ErrorHandler

     ' Save the drawing
     LogWrite "  Saving drawing..."
     swDrawModel.Save
     LogWrite "  Drawing saved"

     ' Get PDF output path
     fileName = Left(drawingPath, InStrRev(drawingPath, ".") - 1)
     fileName = Right(fileName, Len(fileName) - InStrRev(fileName, "\"))
     pdfPath = pdfFolder & fileName & ".PDF"

     ' Verify PDF folder exists
     LogWrite "  Verifying PDF folder: " & pdfFolder
     If Dir(pdfFolder, vbDirectory) = "" Then
         LogWrite "  ERROR: PDF folder does not exist!"
         ProcessSubAssemblyDrawing = False
         Exit Function
     End If
     LogWrite "  PDF folder verified"

     ' Export to PDF
     LogWrite "  Exporting to PDF: " & pdfPath
     On Error Resume Next
       
     ' SaveAs with .PDF extension works (returns -1 but creates file)
     LogWrite "  Using SaveAs to create PDF..."
     saveError = swDrawModel.SaveAs(pdfPath)
     LogWrite "  SaveAs returned code: " & saveError
       
     On Error GoTo ErrorHandler
      
     ' Verify file was actually created
     LogWrite "  Verifying PDF file exists..."
     If Dir(pdfPath) = "" Then
         LogWrite "  ERROR: PDF file not found at: " & pdfPath
         ProcessSubAssemblyDrawing = False
         Exit Function
     End If
      
     ' Get file info
     Dim fileSize As Long
     fileSize = FileLen(pdfPath)
     LogWrite "  PDF file verified: " & fileSize & " bytes"
       
     If fileSize = 0 Then
         LogWrite "  WARNING: PDF file is empty (0 bytes)!"
     End If

     ' Close the drawing without saving again
     swApp.CloseDoc drawingPath

     LogWrite "  PDF export complete: " & pdfPath
     ProcessSubAssemblyDrawing = True
     Exit Function

ErrorHandler:
     LogWrite "  ERROR: " & Err.Description & " (Code: " & Err.Number & ")"
     If Not swDraw Is Nothing Then
         swApp.CloseDoc drawingPath
     End If
     ProcessSubAssemblyDrawing = False

End Function


'============================================================
' PROCESS MAIN ASSEMBLY DRAWING (Single sheet with 3 views)
'============================================================

Function ProcessMainAssemblyDrawing(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, drawingPath As String) As Boolean

     Dim swTopView As SldWorks.View
     Dim swFrontView As SldWorks.View
     Dim swIsometricView As SldWorks.View
     Dim pdfPath As String
     Dim fileName As String
     Dim saveError As Long

     ProcessMainAssemblyDrawing = False

     On Error GoTo ErrorHandler
      
     On Error Resume Next
       
     ' Process Sheet1 (Main Assembly Views - Top, Front, Isometric)
     LogWrite "  Activating Sheet1..."
     If swDrawModel.ActivateSheet("Sheet1") Then
         swDrawModel.ForceRebuild3 False
         LogWrite "    Sheet1 activated"
            
         ' Get the three views (Top, Front, Isometric)
         LogWrite "    Looking for views..."
         Set swTopView = GetMainAssemblyView(swDraw, 1)
         Set swFrontView = GetMainAssemblyView(swDraw, 2)
         Set swIsometricView = GetMainAssemblyView(swDraw, 3)
          
         If Not swTopView Is Nothing Or Not swFrontView Is Nothing Or Not swIsometricView Is Nothing Then
             LogWrite "    Found main assembly views, positioning..."
             ProcessMainAssemblyViews swDraw, swDrawModel, swTopView, swFrontView, swIsometricView
             LogWrite "    Assembly views positioned"
         Else
             LogWrite "    No views found"
         End If
     Else
         LogWrite "    Sheet1 not found, skipping"
     End If

     On Error GoTo ErrorHandler

     ' Save the drawing
     LogWrite "  Saving drawing..."
     swDrawModel.Save
     LogWrite "  Drawing saved"

     ' Get PDF output path
     fileName = Left(drawingPath, InStrRev(drawingPath, ".") - 1)
     fileName = Right(fileName, Len(fileName) - InStrRev(fileName, "\"))
     pdfPath = pdfFolder & fileName & ".PDF"

     ' Verify PDF folder exists
     LogWrite "  Verifying PDF folder: " & pdfFolder
     If Dir(pdfFolder, vbDirectory) = "" Then
         LogWrite "  ERROR: PDF folder does not exist!"
         ProcessMainAssemblyDrawing = False
         Exit Function
     End If
     LogWrite "  PDF folder verified"

     ' Export to PDF
     LogWrite "  Exporting to PDF: " & pdfPath
     On Error Resume Next
       
     ' SaveAs with .PDF extension works (returns -1 but creates file)
     LogWrite "  Using SaveAs to create PDF..."
     saveError = swDrawModel.SaveAs(pdfPath)
     LogWrite "  SaveAs returned code: " & saveError
       
     On Error GoTo ErrorHandler
      
     ' Verify file was actually created
     LogWrite "  Verifying PDF file exists..."
     If Dir(pdfPath) = "" Then
         LogWrite "  ERROR: PDF file not found at: " & pdfPath
         ProcessMainAssemblyDrawing = False
         Exit Function
     End If
      
     ' Get file info
     Dim fileSize As Long
     fileSize = FileLen(pdfPath)
     LogWrite "  PDF file verified: " & fileSize & " bytes"
       
     If fileSize = 0 Then
         LogWrite "  WARNING: PDF file is empty (0 bytes)!"
     End If

     ' Close the drawing without saving again
     swApp.CloseDoc drawingPath

     LogWrite "  PDF export complete: " & pdfPath
     ProcessMainAssemblyDrawing = True
     Exit Function

ErrorHandler:
     LogWrite "  ERROR: " & Err.Description & " (Code: " & Err.Number & ")"
     If Not swDraw Is Nothing Then
         swApp.CloseDoc drawingPath
     End If
     ProcessMainAssemblyDrawing = False

End Function


'============================================================
' PROCESS SHEET1 - EXPLODED VIEW (SAME AS DEBUG MACRO)
'============================================================

Sub ProcessSheet1Exploded(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, swExplodedView As SldWorks.View)
    
    Dim baseName As String
    Dim overallLength As Double
    Dim overallHeight As Double
    Dim scaleNumerator As Double
    Dim scaleDenominator As Double
    Dim scaleValue As Double
    Dim pos(1) As Double
    Dim verifyPos As Variant
    
    On Error Resume Next
    
    ' Extract filename
    baseName = GetDrawingBaseName(swDrawModel.GetPathName)
    LogWrite "      Drawing name: " & baseName
    
    ' Extract dimensions
    LogWrite "      Extracting dimensions..."
    If Not ExtractDimensions(baseName, overallLength, overallHeight) Then
        LogWrite "      ERROR: Failed to extract dimensions"
        Exit Sub
    End If
    LogWrite "      Dimensions - Length: " & overallLength & ", Height: " & overallHeight
    LogWrite ""
    
    ' Determine scale
    LogWrite "      Determining scale..."
    DetermineScale overallLength, overallHeight, scaleNumerator, scaleDenominator
    LogWrite "      Scale: 1:" & scaleDenominator
    scaleValue = scaleNumerator / scaleDenominator
    LogWrite ""
    
    ' Break alignment
    LogWrite "      Breaking view alignment..."
    swExplodedView.BreakAlignment 0
    swDrawModel.ForceRebuild3 False
    LogWrite "      Alignment broken"
    LogWrite ""
    
    ' Apply calculated scale
    LogWrite "      Applying scale 1:" & scaleDenominator & "..."
    LogWrite "      BEFORE: UseSheetScale=" & swExplodedView.UseSheetScale & ", ScaleDecimal=" & swExplodedView.ScaleDecimal
    swExplodedView.UseSheetScale = False
    swExplodedView.ScaleDecimal = scaleValue
    LogWrite "      DESIRED: UseSheetScale=False, ScaleDecimal=" & scaleValue
    LogWrite "      AFTER: UseSheetScale=" & swExplodedView.UseSheetScale & ", ScaleDecimal=" & swExplodedView.ScaleDecimal
    
    If swExplodedView.ScaleDecimal = scaleValue Then
       LogWrite "      ✓ Scale set successfully"
    Else
       LogWrite "      ✗ WARNING: Scale did not stick!"
    End If
    LogWrite ""
    
    ' Rebuild
    LogWrite "      Rebuilding after scale change..."
    swDrawModel.ForceRebuild3 False
    LogWrite "      Scale after rebuild: " & swExplodedView.ScaleDecimal
    
    If swExplodedView.ScaleDecimal <> scaleValue Then
       LogWrite "      ✗ WARNING: Scale changed after rebuild!"
    Else
       LogWrite "      ✓ Scale maintained after rebuild"
    End If
    LogWrite ""
    
    ' Position view
    LogWrite "      Positioning exploded view..."
    LogWrite "      Setting to fixed position: X=8 in, Y=5 in"
    
    LogWrite "      BEFORE: " & swExplodedView.Position(0) & ", " & swExplodedView.Position(1)
    
    pos(0) = 8# * 0.0254  ' Convert inches to meters
    pos(1) = 5# * 0.0254
    
    LogWrite "      DESIRED: " & pos(0) & ", " & pos(1)
    swExplodedView.Position = pos
    
    verifyPos = swExplodedView.Position
    LogWrite "      AFTER: " & verifyPos(0) & ", " & verifyPos(1)
    
    If Abs(verifyPos(0) - pos(0)) < 0.001 And Abs(verifyPos(1) - pos(1)) < 0.001 Then
       LogWrite "      ✓ Position set successfully"
    Else
       LogWrite "      ✗ WARNING: Position did not update correctly!"
    End If
    LogWrite ""
    
    ' Final rebuild
    LogWrite "      Final rebuild after position change..."
    swDrawModel.ForceRebuild3 False
    
    verifyPos = swExplodedView.Position
    LogWrite "      After rebuild: " & verifyPos(0) & ", " & verifyPos(1)
    
    If Abs(verifyPos(0) - pos(0)) > 0.001 Or Abs(verifyPos(1) - pos(1)) > 0.001 Then
       LogWrite "      ✗ WARNING: Position changed after rebuild!"
    Else
       LogWrite "      ✓ Position maintained after rebuild"
    End If
        LogWrite ""
    
        ' Auto-arrange dimensions on the exploded view (same technique
        ' used for the Top/Isometric views and component drawings)
        LogWrite "      Auto-arranging dimensions on exploded view..."
        Dim dimSelectedCountSheet1 As Long
        Dim dimErrorTextSheet1 As String
        AutoArrangeMainAssemblyDimensions swDrawModel, swExplodedView, Nothing, Nothing, dimSelectedCountSheet1, dimErrorTextSheet1
        LogWrite "      Dimensions selected: " & dimSelectedCountSheet1
    
        ' Final rebuild after dimension arrangement
        swDrawModel.ForceRebuild3 False
    
        On Error GoTo 0
    
End Sub


'============================================================
' PROCESS SHEET2 - ASSEMBLY VIEWS (SAME AS DEBUG MACRO)
'============================================================

Sub ProcessSheet2Assembly(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, _
                         swIsometricView As SldWorks.View, swTopView As SldWorks.View)
    
    Dim baseName As String
    Dim overallLength As Double
    Dim overallHeight As Double
    Dim scaleNumerator As Double
    Dim scaleDenominator As Double
    Dim scaleValue As Double
    Dim posIso(1) As Double
    Dim posTop(1) As Double
    Dim verifyIsoPos As Variant
    Dim verifyTopPos As Variant
    
    On Error Resume Next
    
    ' Extract filename
    baseName = GetDrawingBaseName(swDrawModel.GetPathName)
    LogWrite "      Drawing name: " & baseName
    
    ' Extract dimensions
    LogWrite "      Extracting dimensions..."
    If Not ExtractDimensions(baseName, overallLength, overallHeight) Then
        LogWrite "      ERROR: Failed to extract dimensions"
        Exit Sub
    End If
    LogWrite "      Dimensions - Length: " & overallLength & ", Height: " & overallHeight
    LogWrite ""
    
    ' Determine scale
    LogWrite "      Determining scale..."
    DetermineScale overallLength, overallHeight, scaleNumerator, scaleDenominator
    LogWrite "      Scale: 1:" & scaleDenominator
    scaleValue = scaleNumerator / scaleDenominator
    LogWrite ""
    
    ' Break alignment
    LogWrite "      Breaking view alignment..."
    If Not swIsometricView Is Nothing Then
       swIsometricView.BreakAlignment 0
    End If
    If Not swTopView Is Nothing Then
       swTopView.BreakAlignment 0
    End If
    swDrawModel.ForceRebuild3 False
    LogWrite "      Alignment broken"
    LogWrite ""
    
    ' Apply calculated scale
    LogWrite "      Applying scale 1:" & scaleDenominator & " to all views..."
    
    If Not swIsometricView Is Nothing Then
       LogWrite "      Isometric BEFORE: UseSheetScale=" & swIsometricView.UseSheetScale & ", ScaleDecimal=" & swIsometricView.ScaleDecimal
       swIsometricView.UseSheetScale = False
       swIsometricView.ScaleDecimal = scaleValue
       LogWrite "      Isometric DESIRED: UseSheetScale=False, ScaleDecimal=" & scaleValue
       LogWrite "      Isometric AFTER: UseSheetScale=" & swIsometricView.UseSheetScale & ", ScaleDecimal=" & swIsometricView.ScaleDecimal
        
       If swIsometricView.ScaleDecimal = scaleValue Then
           LogWrite "      ✓ Isometric scale set successfully"
       Else
           LogWrite "      ✗ WARNING: Isometric scale did not stick!"
       End If
    End If
    
    If Not swTopView Is Nothing Then
       LogWrite "      Top BEFORE: UseSheetScale=" & swTopView.UseSheetScale & ", ScaleDecimal=" & swTopView.ScaleDecimal
       swTopView.UseSheetScale = False
       swTopView.ScaleDecimal = scaleValue
       LogWrite "      Top DESIRED: UseSheetScale=False, ScaleDecimal=" & scaleValue
       LogWrite "      Top AFTER: UseSheetScale=" & swTopView.UseSheetScale & ", ScaleDecimal=" & swTopView.ScaleDecimal
        
       If swTopView.ScaleDecimal = scaleValue Then
           LogWrite "      ✓ Top view scale set successfully"
       Else
           LogWrite "      ✗ WARNING: Top view scale did not stick!"
       End If
    End If
    LogWrite ""
    
    ' Rebuild
    LogWrite "      Rebuilding after scale change..."
    swDrawModel.ForceRebuild3 False
    
    If Not swIsometricView Is Nothing Then
       LogWrite "      Isometric scale after rebuild: " & swIsometricView.ScaleDecimal
    End If
    If Not swTopView Is Nothing Then
       LogWrite "      Top scale after rebuild: " & swTopView.ScaleDecimal
    End If
    LogWrite ""
    
    ' Position views
    LogWrite "      Positioning views..."
    
    If Not swIsometricView Is Nothing Then
       LogWrite "      Isometric view:"
       LogWrite "        Setting to fixed position: X=11 in, Y=7.5 in"
       LogWrite "        BEFORE: " & swIsometricView.Position(0) & ", " & swIsometricView.Position(1)
        
       posIso(0) = 11# * 0.0254
       posIso(1) = 7.5 * 0.0254
        
       LogWrite "        DESIRED: " & posIso(0) & ", " & posIso(1)
       swIsometricView.Position = posIso
        
       verifyIsoPos = swIsometricView.Position
       LogWrite "        AFTER: " & verifyIsoPos(0) & ", " & verifyIsoPos(1)
        
       If Abs(verifyIsoPos(0) - posIso(0)) < 0.001 And Abs(verifyIsoPos(1) - posIso(1)) < 0.001 Then
           LogWrite "        ✓ Isometric position set successfully"
       Else
           LogWrite "        ✗ WARNING: Isometric position did not update correctly!"
       End If
    End If
    
    If Not swTopView Is Nothing Then
       LogWrite "      Top view:"
       LogWrite "        Setting to fixed position: X=7.5 in, Y=3.5 in"
       LogWrite "        BEFORE: " & swTopView.Position(0) & ", " & swTopView.Position(1)
        
       posTop(0) = 7.5 * 0.0254
       posTop(1) = 3.5 * 0.0254
        
       LogWrite "        DESIRED: " & posTop(0) & ", " & posTop(1)
       swTopView.Position = posTop
        
       verifyTopPos = swTopView.Position
       LogWrite "        AFTER: " & verifyTopPos(0) & ", " & verifyTopPos(1)
        
       If Abs(verifyTopPos(0) - posTop(0)) < 0.001 And Abs(verifyTopPos(1) - posTop(1)) < 0.001 Then
           LogWrite "        ✓ Top view position set successfully"
       Else
           LogWrite "        ✗ WARNING: Top view position did not update correctly!"
       End If
    End If
    LogWrite ""
    
    ' Final rebuild
    LogWrite "      Final rebuild after position changes..."
    swDrawModel.ForceRebuild3 False
    
    If Not swIsometricView Is Nothing Then
       verifyIsoPos = swIsometricView.Position
       LogWrite "      Isometric after rebuild: " & verifyIsoPos(0) & ", " & verifyIsoPos(1)
        
       If Abs(verifyIsoPos(0) - posIso(0)) > 0.001 Or Abs(verifyIsoPos(1) - posIso(1)) > 0.001 Then
           LogWrite "      ✗ WARNING: Isometric position changed after rebuild!"
       Else
           LogWrite "      ✓ Isometric position maintained after rebuild"
       End If
    End If
    
    If Not swTopView Is Nothing Then
       verifyTopPos = swTopView.Position
       LogWrite "      Top after rebuild: " & verifyTopPos(0) & ", " & verifyTopPos(1)
        
       If Abs(verifyTopPos(0) - posTop(0)) > 0.001 Or Abs(verifyTopPos(1) - posTop(1)) > 0.001 Then
           LogWrite "      ✗ WARNING: Top view position changed after rebuild!"
       Else
           LogWrite "      ✓ Top view position maintained after rebuild"
       End If
    End If
    
    On Error GoTo 0
    
End Sub


'============================================================
' GET FIRST MODEL VIEW (EXPLODED VIEW)
'============================================================

Function GetFirstModelView(swDraw As SldWorks.DrawingDoc) As SldWorks.View

     Dim swView As SldWorks.View

     Set swView = swDraw.GetFirstView

     If swView Is Nothing Then
         Set GetFirstModelView = Nothing
         Exit Function
     End If

     ' Skip sheet view, get to first model view
     Set swView = swView.GetNextView

     Set GetFirstModelView = swView

End Function


'============================================================
' GET NAMED VIEW (e.g., "Isometric", "Top")
' Uses position-based heuristic: 1st model view = Top, 2nd = Isometric
'============================================================

Function GetNamedView(swDraw As SldWorks.DrawingDoc, viewName As String) As SldWorks.View

     Dim swView As SldWorks.View
     Dim viewCount As Long
      
     Set GetNamedView = Nothing
     viewCount = 0

     On Error Resume Next

     Set swView = swDraw.GetFirstView

     Do While Not swView Is Nothing
         If Not swView.ReferencedDocument Is Nothing Then
             viewCount = viewCount + 1
              
             ' Use heuristic: 1st model view = Top, 2nd = Isometric
             If viewName = "Top" And viewCount = 1 Then
                 Set GetNamedView = swView
                 Exit Function
             ElseIf viewName = "Isometric" And viewCount = 2 Then
                 Set GetNamedView = swView
                 Exit Function
             End If
         End If
         Set swView = swView.GetNextView
     Loop

     On Error GoTo 0

End Function


'============================================================
' DETERMINE SCALE - Calculate proper scale for drawing size
'============================================================

Sub DetermineScale(overallLength As Double, overallHeight As Double, _
                   ByRef scaleNumerator As Double, ByRef scaleDenominator As Double)
    
    ' Scale thresholds based on drawing dimensions (in inches)
    ' Hierarchy: Check Height first, then Length
    
    If overallHeight >= 36 Or overallLength >= 96 Then
        ' Very large drawings: 1:20 scale
        scaleNumerator = 1
        scaleDenominator = 20
    ElseIf overallHeight >= 24 Or overallLength >= 72 Then
        ' Large drawings: 1:18 scale
        scaleNumerator = 1
        scaleDenominator = 18
    ElseIf overallHeight >= 18 Or overallLength >= 60 Then
        ' Medium-large drawings: 1:16 scale
        scaleNumerator = 1
        scaleDenominator = 16
    ElseIf overallHeight >= 12 Or overallLength >= 48 Then
        ' Medium drawings: 1:14 scale
        scaleNumerator = 1
        scaleDenominator = 14
    ElseIf overallHeight >= 6 Or overallLength >= 24 Then
        ' Small drawings: 1:12 scale
        scaleNumerator = 1
        scaleDenominator = 12
    Else
        ' Very small drawings: 1:10 scale
        scaleNumerator = 1
        scaleDenominator = 10
    End If
    
End Sub


'============================================================
' POSITION AND SCALE EXPLODED VIEW (PAGE 1)
'============================================================

Sub PositionAndScaleExplodedView(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, _
                                 swExplodedView As SldWorks.View)

     Dim modelPath As String
     Dim baseName As String
     Dim overallLength As Double
     Dim overallHeight As Double
     Dim scaleNumerator As Double
     Dim scaleDenominator As Double
     Dim sheetWidth As Double
     Dim sheetHeight As Double
     Dim swSheet As SldWorks.Sheet
     Dim outline As Variant
     Dim viewWidth As Double
     Dim viewHeight As Double
     Dim centerX As Double
     Dim centerY As Double
     Dim curPos As Variant
     Dim offsetX As Double
     Dim offsetY As Double
     Dim targetX As Double
     Dim targetY As Double
     Dim pos(1) As Double

     On Error Resume Next

     ' Get referenced model path
     modelPath = swExplodedView.GetReferencedModelName
     If modelPath = "" Then
         LogWrite "      WARNING: Could not get referenced model name"
         Exit Sub
     End If
     LogWrite "      Referenced model: " & modelPath

     ' Extract filename
     Dim fileName As String
     fileName = Mid(modelPath, InStrRev(modelPath, "\") + 1)
     If InStrRev(fileName, ".") > 0 Then
         baseName = Left(fileName, InStrRev(fileName, ".") - 1)
     Else
         baseName = fileName
     End If
     LogWrite "      Base name: " & baseName

     ' Extract dimensions flexibly from filename
     LogWrite "      Extracting dimensions from filename..."
     If Not ExtractDimensions(baseName, overallLength, overallHeight) Then
         LogWrite "      ERROR: Failed to extract dimensions"
         Exit Sub
     End If
     LogWrite "      Dimensions - Length: " & overallLength & ", Height: " & overallHeight

     ' Determine scale
     LogWrite "      Determining scale..."
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
     LogWrite "      Scale: 1:" & scaleDenominator

     ' Break alignment
     LogWrite "      Breaking view alignment..."
     BreakViewAlignment swDraw, swDrawModel, swExplodedView
     LogWrite "      Alignment broken - COMPLETED"

     ' Apply scale
     LogWrite "      Applying scale 1:" & scaleDenominator & "..."
     Dim scaleValue As Double
     scaleValue = scaleNumerator / scaleDenominator
      
     ' Log BEFORE state
     LogWrite "        BEFORE: UseSheetScale=" & swExplodedView.UseSheetScale & ", ScaleDecimal=" & swExplodedView.ScaleDecimal
     LogWrite "        DESIRED: UseSheetScale=False, ScaleDecimal=" & scaleValue
      
     swExplodedView.UseSheetScale = False
     swExplodedView.ScaleDecimal = scaleValue
      
     ' Log AFTER state
     LogWrite "        AFTER: UseSheetScale=" & swExplodedView.UseSheetScale & ", ScaleDecimal=" & swExplodedView.ScaleDecimal
     If swExplodedView.ScaleDecimal = scaleValue Then
         LogWrite "        ✓ Scale set successfully"
     Else
         LogWrite "        ✗ WARNING: Scale did not stick! Expected " & scaleValue & " but got " & swExplodedView.ScaleDecimal
     End If

     ' Rebuild
     LogWrite "      Rebuilding after scale change..."
     swDraw.ForceRebuild3 False
     LogWrite "      Scale after rebuild: " & swExplodedView.ScaleDecimal
     If swExplodedView.ScaleDecimal <> scaleValue Then
         LogWrite "        ✗ WARNING: Scale changed after rebuild!"
     End If

     ' Get sheet size
     LogWrite "      Getting sheet size..."
     Set swSheet = swDraw.GetCurrentSheet
     If swSheet Is Nothing Then 
         LogWrite "      ERROR: Could not get current sheet"
         Exit Sub
     End If
     swSheet.GetSize sheetWidth, sheetHeight
     LogWrite "      Sheet size: " & sheetWidth & " x " & sheetHeight & " (check units - may be in meters)"

     ' Get outline
     LogWrite "      Getting view outline..."
     outline = swExplodedView.GetOutline
     viewWidth = outline(2) - outline(0)
     viewHeight = outline(3) - outline(1)
     LogWrite "      View size: " & viewWidth & " x " & viewHeight & " (check units)"
     LogWrite "      Outline array: [" & outline(0) & ", " & outline(1) & ", " & outline(2) & ", " & outline(3) & "]"

     ' Center view on sheet
     LogWrite "      Calculating position..."
     LogWrite "      WARNING: Check if sheet size (" & sheetWidth & ") and view size (" & viewWidth & ") are in same units!"
     centerX = (sheetWidth - viewWidth) / 2#
     centerY = (sheetHeight - viewHeight) / 2#

     ' Move 1 inch left
     centerX = centerX - (1# * 0.0254)
     LogWrite "      Center position: " & centerX & ", " & centerY & " (may be incorrect if units mismatch)"

     ' Get current position
     curPos = swExplodedView.Position
     offsetX = curPos(0) - outline(0)
     offsetY = curPos(1) - outline(1)
     LogWrite "      Current position: " & curPos(0) & ", " & curPos(1)
     LogWrite "      Offset: " & offsetX & ", " & offsetY

     ' Calculate target position
     targetX = centerX + offsetX
     targetY = centerY + offsetY
     LogWrite "      Target position: " & targetX & ", " & targetY & " (may be incorrect if units mismatch)"

     ' Set position using fixed coordinates (like the working debug macro)
     ' Exploded view: X=8 inches, Y=5 inches
     LogWrite "      Setting position..."
     LogWrite "        BEFORE: " & swExplodedView.Position(0) & ", " & swExplodedView.Position(1)
      
     pos(0) = 8# * 0.0254  ' Convert inches to meters
     pos(1) = 5# * 0.0254
      
     LogWrite "        DESIRED: " & pos(0) & ", " & pos(1)
     swExplodedView.Position = pos
       
     ' Verify position was set
     Dim verifyPos As Variant
     verifyPos = swExplodedView.Position
     LogWrite "        AFTER: " & verifyPos(0) & ", " & verifyPos(1)
       
     If Abs(verifyPos(0) - pos(0)) < 0.001 And Abs(verifyPos(1) - pos(1)) < 0.001 Then
         LogWrite "        ✓ Position set successfully"
     Else
         LogWrite "        ✗ WARNING: Position did not update correctly!"
     End If

     ' Rebuild
     LogWrite "      Final rebuild after position change..."
     swDraw.ForceRebuild3 False
      
     ' Verify position after rebuild
     verifyPos = swExplodedView.Position
     LogWrite "        After rebuild: " & verifyPos(0) & ", " & verifyPos(1)
     If Abs(verifyPos(0) - pos(0)) > 0.001 Or Abs(verifyPos(1) - pos(1)) > 0.001 Then
         LogWrite "        ✗ WARNING: Position changed after rebuild!"
     Else
         LogWrite "        ✓ Position maintained after rebuild"
     End If
     LogWrite "      Final rebuild complete"

     On Error GoTo 0

End Sub

'============================================================
' POSITION AND SCALE ASSEMBLY VIEWS (PAGE 2)
'
' Isometric view positioned at top-right
' Top view positioned centered below
'============================================================

Sub PositionAndScaleAssemblyViews(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, _
                                 swIsometricView As SldWorks.View, swTopView As SldWorks.View)

     Dim modelPath As String
     Dim baseName As String
     Dim overallLength As Double
     Dim overallHeight As Double
     Dim scaleNumerator As Double
     Dim scaleDenominator As Double
     Dim swSheet As SldWorks.Sheet
     Dim posIso(1) As Double
     Dim posTop(1) As Double

     On Error Resume Next

     ' Get model path from isometric or top view
     LogWrite "      Getting model path..."
     If Not swIsometricView Is Nothing Then
         modelPath = swIsometricView.GetReferencedModelName
         LogWrite "      Model path from isometric view: " & modelPath
     ElseIf Not swTopView Is Nothing Then
         modelPath = swTopView.GetReferencedModelName
         LogWrite "      Model path from top view: " & modelPath
     End If

     If modelPath = "" Then
         LogWrite "      ERROR: Could not get model path"
         Exit Sub
     End If

     ' Extract filename
     Dim fileName As String
     fileName = Mid(modelPath, InStrRev(modelPath, "\") + 1)
     If InStrRev(fileName, ".") > 0 Then
         baseName = Left(fileName, InStrRev(fileName, ".") - 1)
     Else
         baseName = fileName
     End If
     LogWrite "      Base name: " & baseName

     ' Extract dimensions
     LogWrite "      Extracting dimensions..."
     If Not ExtractDimensions(baseName, overallLength, overallHeight) Then
         LogWrite "      ERROR: Failed to extract dimensions"
         Exit Sub
     End If
     LogWrite "      Dimensions - Length: " & overallLength & ", Height: " & overallHeight

     ' Determine scale
     LogWrite "      Determining scale..."
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
     LogWrite "      Scale: 1:" & scaleDenominator

     ' Break alignment for both views
     LogWrite "      Breaking view alignment..."
     If Not swIsometricView Is Nothing Then
         LogWrite "        Breaking isometric view alignment..."
         BreakViewAlignment swDraw, swDrawModel, swIsometricView
     End If
     If Not swTopView Is Nothing Then
         LogWrite "        Breaking top view alignment..."
         BreakViewAlignment swDraw, swDrawModel, swTopView
     End If
     LogWrite "      Alignment broken"

     ' Apply scale to both views
     LogWrite "      Applying scale to views..."
     If Not swIsometricView Is Nothing Then
         LogWrite "        Setting isometric view scale..."
         swIsometricView.UseSheetScale = False
         swIsometricView.ScaleDecimal = scaleNumerator / scaleDenominator
         LogWrite "        Isometric scale: " & swIsometricView.ScaleDecimal
     End If
     If Not swTopView Is Nothing Then
         LogWrite "        Setting top view scale..."
         swTopView.UseSheetScale = False
         swTopView.ScaleDecimal = scaleNumerator / scaleDenominator
         LogWrite "        Top view scale: " & swTopView.ScaleDecimal
     End If
     LogWrite "      Scale applied"

     ' Rebuild
     LogWrite "      Rebuilding..."
     swDraw.ForceRebuild3 False
     LogWrite "      Rebuild complete"

     ' Setting view positions (using fixed coordinates)
     LogWrite "      Setting view positions..."
     If Not swIsometricView Is Nothing Then
         LogWrite "        Isometric view:"
         LogWrite "          BEFORE: " & swIsometricView.Position(0) & ", " & swIsometricView.Position(1)
          
         ' Isometric at X=11 inches, Y=7.5 inches
         posIso(0) = 11# * 0.0254  ' Convert inches to meters
         posIso(1) = 7.5 * 0.0254
          
         LogWrite "          DESIRED: " & posIso(0) & ", " & posIso(1)
         swIsometricView.Position = posIso
           
         ' Verify isometric position was set
         Dim verifyIsoPos As Variant
         verifyIsoPos = swIsometricView.Position
         LogWrite "          AFTER: " & verifyIsoPos(0) & ", " & verifyIsoPos(1)
           
         If Abs(verifyIsoPos(0) - posIso(0)) < 0.001 And Abs(verifyIsoPos(1) - posIso(1)) < 0.001 Then
             LogWrite "          ✓ Isometric position set successfully"
         Else
             LogWrite "          ✗ WARNING: Isometric position did not update correctly!"
         End If
     End If

     If Not swTopView Is Nothing Then
         LogWrite "        Top view:"
         LogWrite "          BEFORE: " & swTopView.Position(0) & ", " & swTopView.Position(1)
          
         ' Top view at X=7.5 inches, Y=3.5 inches
         posTop(0) = 7.5 * 0.0254  ' Convert inches to meters
         posTop(1) = 3.5 * 0.0254
          
         LogWrite "          DESIRED: " & posTop(0) & ", " & posTop(1)
         swTopView.Position = posTop
           
         ' Verify top position was set
         Dim verifyTopPos As Variant
         verifyTopPos = swTopView.Position
         LogWrite "          AFTER: " & verifyTopPos(0) & ", " & verifyTopPos(1)
           
         If Abs(verifyTopPos(0) - posTop(0)) < 0.001 And Abs(verifyTopPos(1) - posTop(1)) < 0.001 Then
             LogWrite "          ✓ Top view position set successfully"
         Else
             LogWrite "          ✗ WARNING: Top view position did not update correctly!"
         End If
     End If

     ' Rebuild
     LogWrite "      Rebuilding after position changes..."
     swDraw.ForceRebuild3 False
       
     ' Verify positions after rebuild
     If Not swIsometricView Is Nothing Then
         verifyIsoPos = swIsometricView.Position
         LogWrite "        Isometric after rebuild: " & verifyIsoPos(0) & ", " & verifyIsoPos(1)
         If Abs(verifyIsoPos(0) - targetIsoX) > 0.001 Or Abs(verifyIsoPos(1) - targetIsoY) > 0.001 Then
             LogWrite "        ✗ WARNING: Isometric position changed after rebuild!"
         Else
             LogWrite "        ✓ Isometric position maintained"
         End If
     End If
     If Not swTopView Is Nothing Then
         verifyTopPos = swTopView.Position
         LogWrite "        Top view after rebuild: " & verifyTopPos(0) & ", " & verifyTopPos(1)
         If Abs(verifyTopPos(0) - targetTopX) > 0.001 Or Abs(verifyTopPos(1) - targetTopY) > 0.001 Then
             LogWrite "        ✗ WARNING: Top view position changed after rebuild!"
         Else
             LogWrite "        ✓ Top view position maintained"
         End If
     End If
     LogWrite "      Rebuild complete"

     ' Auto-arrange dimensions on BOTH the top view and isometric view
     ' (previously only the top view's dimensions were arranged here -
     ' isometric view dimensions were left untouched)
     LogWrite "      Adjusting top + isometric view dimensions..."
     Dim dimSelectedCountSheet2 As Long
     Dim dimErrorTextSheet2 As String
     AutoArrangeMainAssemblyDimensions swDrawModel, swTopView, swIsometricView, Nothing, dimSelectedCountSheet2, dimErrorTextSheet2
     LogWrite "      Dimensions selected: " & dimSelectedCountSheet2
     LogWrite "      Dimension adjustment complete"

     ' Final rebuild
     LogWrite "      Final rebuild..."
     swDraw.ForceRebuild3 False
     LogWrite "      Final rebuild complete"

     On Error GoTo 0

End Sub


'============================================================
' AUTO ADJUST TOP VIEW DIMENSIONS
'
' Handles the single dimension on the top view
'============================================================

Sub AutoAdjustTopViewDimensions(swDrawModel As SldWorks.ModelDoc2, swTopView As SldWorks.View)

     Dim swAnn As SldWorks.Annotation
     Dim swDispDim As SldWorks.DisplayDimension
     Dim swSelMgr As SldWorks.SelectionMgr
     Dim swSelData As SldWorks.SelectData
     Dim selectedCount As Long
     Dim selOK As Boolean

     selectedCount = 0

     If swTopView Is Nothing Then
         Exit Sub
     End If

     Set swSelMgr = swDrawModel.SelectionManager
     Set swSelData = swSelMgr.CreateSelectData

     On Error Resume Next
     swDrawModel.ClearSelection2 True
     On Error GoTo 0

     ' Get all dimensions in top view
     Set swAnn = swTopView.GetFirstAnnotation3
     Do While Not swAnn Is Nothing
         If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
             Set swDispDim = swAnn.GetSpecificAnnotation
             If Not swDispDim Is Nothing Then
                 ' Auto-jog ordinate dimension
                 If IsOrdinateDimension(swDispDim) Then
                     On Error Resume Next
                     swDispDim.AutoJogOrdinate()
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

     ' Auto arrange
     If selectedCount > 0 Then
         On Error Resume Next
         swDrawModel.Extension.AlignDimensions swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001
         On Error GoTo 0
     End If

     ' Clear selection
     On Error Resume Next
     swDrawModel.ClearSelection2 True
     On Error GoTo 0

End Sub


'============================================================
' BREAK VIEW ALIGNMENT
'============================================================

Function BreakViewAlignment(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, _
                            swView As SldWorks.View) As Boolean

     BreakViewAlignment = False

     On Error Resume Next
     swView.BreakAlignment 0
     swDraw.ForceRebuild3 False
     BreakViewAlignment = True
     On Error GoTo 0

End Function


'============================================================
' CHECK IF DIMENSION IS ORDINATE
'============================================================

Function IsOrdinateDimension(swDispDim As SldWorks.DisplayDimension) As Boolean

     On Error Resume Next
     IsOrdinateDimension = (swDispDim.GetDimensionType = swOrdinateDimension)
     On Error GoTo 0

End Function


'============================================================
' IS MAIN ASSEMBLY DRAWING - Detect if filename is main assembly
' Main assembly filenames are longer with format: ##-#-C#####-S#####-####-##
' Sub-assembly filenames are shorter: #-C####-####-## or similar
'============================================================

Function IsMainAssemblyDrawing(baseName As String) As Boolean

     Dim fileParts() As String
     Dim cCount As Long
     Dim sCount As Long
     Dim i As Long
     Dim part As String
     Dim partUpper As String
     
     IsMainAssemblyDrawing = False
     cCount = 0
     sCount = 0
     
     fileParts = Split(baseName, "-")
     
     ' Count C and S prefixes
     For i = 0 To UBound(fileParts)
          part = fileParts(i)
          partUpper = UCase(part)
          
          ' Look for C prefix (C followed by digits)
          If Left(partUpper, 1) = "C" And Len(part) > 1 Then
               If IsNumeric(Mid(part, 2)) Then
                    cCount = cCount + 1
               End If
          End If
          
          ' Look for S prefix (S followed by digits, but not SL/SC/SR)
          If Left(partUpper, 1) = "S" And Len(part) > 1 Then
               If Not (Left(partUpper, 2) = "SL" Or Left(partUpper, 2) = "SC" Or Left(partUpper, 2) = "SR") Then
                    If IsNumeric(Mid(part, 2)) Then
                         sCount = sCount + 1
                    End If
               End If
          End If
     Next i
     
     ' Main assembly has BOTH C and S values
     IsMainAssemblyDrawing = (cCount >= 1 And sCount >= 1)
     
End Function


'============================================================
' PROCESS MAIN ASSEMBLY VIEWS - Scale and position all views
'============================================================

Sub ProcessMainAssemblyViews(swDraw As SldWorks.DrawingDoc, swDrawModel As SldWorks.ModelDoc2, _
                             swTopView As SldWorks.View, swFrontView As SldWorks.View, _
                             swIsometricView As SldWorks.View)
    
    Dim baseName As String
    Dim lengthC As Double
    Dim lengthS As Double
    Dim height As Double
    Dim maxDimension As Double
    Dim scaleNumerator As Double
    Dim scaleDenominator As Double
    Dim scaleValue As Double
    
    On Error Resume Next
    
    ' Extract filename
    baseName = GetDrawingBaseName(swDrawModel.GetPathName)
    
    ' Extract dimensions from main assembly name
    LogWrite "      Extracting dimensions from main assembly..."
    If Not ExtractMainAssemblyDimensions(baseName, lengthC, lengthS, height) Then
        LogWrite "      ERROR: Failed to extract dimensions"
        Exit Sub
    End If
    LogWrite "      Dimensions - C: " & lengthC & ", S: " & lengthS & ", Height: " & height
    LogWrite ""
    
    ' Use larger of C or S for scale calculation
    maxDimension = IIf(lengthC > lengthS, lengthC, lengthS)
    LogWrite "      Max dimension (larger of C/S): " & maxDimension
    
    ' Determine scale for main assembly
    LogWrite "      Determining scale..."
    DetermineMainAssemblyScale maxDimension, scaleNumerator, scaleDenominator
    LogWrite "      Scale: 1:" & scaleDenominator
    scaleValue = scaleNumerator / scaleDenominator
    LogWrite ""
    
    ' Break alignment on all views - try more aggressive unlocking
    LogWrite "      Breaking view alignment..."
    If Not swTopView Is Nothing Then
        swTopView.BreakAlignment 0
        swTopView.Locked = False
    End If
    If Not swFrontView Is Nothing Then
        swFrontView.BreakAlignment 0
        swFrontView.Locked = False
    End If
    If Not swIsometricView Is Nothing Then
        swIsometricView.BreakAlignment 0
        swIsometricView.Locked = False
    End If
    swDrawModel.ForceRebuild3 False
    LogWrite "      Alignment broken and views unlocked"
    LogWrite ""
    
    ' Apply calculated scale to all views
    LogWrite "      Applying scale 1:" & scaleDenominator & " to all views..."
    
    If Not swTopView Is Nothing Then
        LogWrite "      Top BEFORE: UseSheetScale=" & swTopView.UseSheetScale & ", ScaleDecimal=" & swTopView.ScaleDecimal
        swTopView.UseSheetScale = False
        swTopView.ScaleDecimal = scaleValue
        LogWrite "      Top AFTER: UseSheetScale=" & swTopView.UseSheetScale & ", ScaleDecimal=" & swTopView.ScaleDecimal
        
        If swTopView.ScaleDecimal = scaleValue Then
            LogWrite "      ✓ Top view scale set successfully"
        Else
            LogWrite "      ✗ WARNING: Top view scale did not stick!"
        End If
    End If
    
    If Not swFrontView Is Nothing Then
        LogWrite "      Front BEFORE: UseSheetScale=" & swFrontView.UseSheetScale & ", ScaleDecimal=" & swFrontView.ScaleDecimal
        swFrontView.UseSheetScale = False
        swFrontView.ScaleDecimal = scaleValue
        LogWrite "      Front AFTER: UseSheetScale=" & swFrontView.UseSheetScale & ", ScaleDecimal=" & swFrontView.ScaleDecimal
        
        If swFrontView.ScaleDecimal = scaleValue Then
            LogWrite "      ✓ Front view scale set successfully"
        Else
            LogWrite "      ✗ WARNING: Front view scale did not stick!"
        End If
    End If
    
    If Not swIsometricView Is Nothing Then
        LogWrite "      Isometric BEFORE: UseSheetScale=" & swIsometricView.UseSheetScale & ", ScaleDecimal=" & swIsometricView.ScaleDecimal
        swIsometricView.UseSheetScale = False
        swIsometricView.ScaleDecimal = scaleValue
        LogWrite "      Isometric AFTER: UseSheetScale=" & swIsometricView.UseSheetScale & ", ScaleDecimal=" & swIsometricView.ScaleDecimal
        
        If swIsometricView.ScaleDecimal = scaleValue Then
            LogWrite "      ✓ Isometric view scale set successfully"
        Else
            LogWrite "      ✗ WARNING: Isometric view scale did not stick!"
        End If
    End If
    LogWrite ""
    
    ' Rebuild after scale changes
    LogWrite "      Rebuilding after scale changes..."
    swDrawModel.ForceRebuild3 False
    LogWrite "      Rebuild complete"
    LogWrite ""
    
    ' Position views based on their CENTERS instead of absolute coords
    LogWrite "      Positioning views by center..."
    Dim pos(1) As Double
    Dim verifyPos As Variant
    Dim outline As Variant
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim viewCenterX As Double
    Dim viewCenterY As Double
    Dim desiredCenterX As Double
    Dim desiredCenterY As Double
    Dim posX As Double
    Dim posY As Double
    
    ' Define desired center positions for each view (in inches, will convert to meters)
    ' Measured via GET_VIEW_POSITIONS.bas (outline-based CENTER values)
    ' Top view center: X=5.31, Y=6.27
    ' Front view center: X=5.31, Y=2.78
    ' Isometric view center: X=11.29, Y=8.51
    
    If Not swTopView Is Nothing Then
        LogWrite "      Top view (positioning by center):"
        outline = swTopView.GetOutline
        If Not IsEmpty(outline) Then
            viewWidth = outline(2) - outline(0)
            viewHeight = outline(3) - outline(1)
            viewCenterX = outline(0) + (viewWidth / 2)
            viewCenterY = outline(1) + (viewHeight / 2)
            
            desiredCenterX = 5.31 * 0.0254
            desiredCenterY = 6.27 * 0.0254
            
            ' NOTE: View.Position is the view's internal anchor point, NOT the
            ' bounding-box (outline) corner. The anchor-to-outline offset scales
            ' with the view scale, so it must be measured fresh each time rather
            ' than assumed to be zero - otherwise the view appears to "move" when
            ' the scale changes even though the math below is internally consistent.
            Dim curPosTop As Variant
            Dim offsetXTop As Double, offsetYTop As Double
            curPosTop = swTopView.Position
            offsetXTop = outline(0) - curPosTop(0)
            offsetYTop = outline(1) - curPosTop(1)
            
            ' Calculate position so view center ends up at desired location,
            ' correcting for the anchor->outline offset measured above
            posX = (desiredCenterX - (viewWidth / 2)) - offsetXTop
            posY = (desiredCenterY - (viewHeight / 2)) - offsetYTop
            
            LogWrite "        View center: " & viewCenterX & ", " & viewCenterY
            LogWrite "        Desired center: " & desiredCenterX & ", " & desiredCenterY
            LogWrite "        Anchor offset: " & offsetXTop & ", " & offsetYTop
            LogWrite "        Setting position to: " & posX & ", " & posY
            LogWrite "        BEFORE: " & swTopView.Position(0) & ", " & swTopView.Position(1)
            
            pos(0) = posX
            pos(1) = posY
            swTopView.Position = pos
            
            verifyPos = swTopView.Position
            LogWrite "        AFTER: " & verifyPos(0) & ", " & verifyPos(1)
            
            ' Verify center is now correct
            Dim newOutline As Variant
            newOutline = swTopView.GetOutline
            Dim newCenterX As Double, newCenterY As Double
            newCenterX = newOutline(0) + ((newOutline(2) - newOutline(0)) / 2)
            newCenterY = newOutline(1) + ((newOutline(3) - newOutline(1)) / 2)
            LogWrite "        New center: " & newCenterX & ", " & newCenterY
        End If
    End If
    
    If Not swFrontView Is Nothing Then
        LogWrite "      Front view (positioning by center):"
        outline = swFrontView.GetOutline
        If Not IsEmpty(outline) Then
            viewWidth = outline(2) - outline(0)
            viewHeight = outline(3) - outline(1)
            
            desiredCenterX = 5.31 * 0.0254
            desiredCenterY = 2.78 * 0.0254
            
            Dim curPosFront As Variant
            Dim offsetXFront As Double, offsetYFront As Double
            curPosFront = swFrontView.Position
            offsetXFront = outline(0) - curPosFront(0)
            offsetYFront = outline(1) - curPosFront(1)
            
            posX = (desiredCenterX - (viewWidth / 2)) - offsetXFront
            posY = (desiredCenterY - (viewHeight / 2)) - offsetYFront
            
            LogWrite "        View size: " & viewWidth & " x " & viewHeight
            LogWrite "        Desired center: " & desiredCenterX & ", " & desiredCenterY
            LogWrite "        Anchor offset: " & offsetXFront & ", " & offsetYFront
            LogWrite "        Setting position to: " & posX & ", " & posY
            LogWrite "        BEFORE: " & swFrontView.Position(0) & ", " & swFrontView.Position(1)
            
            pos(0) = posX
            pos(1) = posY
            swFrontView.Position = pos
            
            verifyPos = swFrontView.Position
            LogWrite "        AFTER: " & verifyPos(0) & ", " & verifyPos(1)
            
            Dim newOutline2 As Variant
            newOutline2 = swFrontView.GetOutline
            Dim newCenterX2 As Double, newCenterY2 As Double
            newCenterX2 = newOutline2(0) + ((newOutline2(2) - newOutline2(0)) / 2)
            newCenterY2 = newOutline2(1) + ((newOutline2(3) - newOutline2(1)) / 2)
            LogWrite "        New center: " & newCenterX2 & ", " & newCenterY2
        End If
    End If
    
    If Not swIsometricView Is Nothing Then
        LogWrite "      Isometric view (positioning by center):"
        outline = swIsometricView.GetOutline
        If Not IsEmpty(outline) Then
            viewWidth = outline(2) - outline(0)
            viewHeight = outline(3) - outline(1)
            
            desiredCenterX = 11.29 * 0.0254
            desiredCenterY = 8.51 * 0.0254
            
            Dim curPosIso As Variant
            Dim offsetXIso As Double, offsetYIso As Double
            curPosIso = swIsometricView.Position
            offsetXIso = outline(0) - curPosIso(0)
            offsetYIso = outline(1) - curPosIso(1)
            
            posX = (desiredCenterX - (viewWidth / 2)) - offsetXIso
            posY = (desiredCenterY - (viewHeight / 2)) - offsetYIso
            
            LogWrite "        View size: " & viewWidth & " x " & viewHeight
            LogWrite "        Desired center: " & desiredCenterX & ", " & desiredCenterY
            LogWrite "        Anchor offset: " & offsetXIso & ", " & offsetYIso
            LogWrite "        Setting position to: " & posX & ", " & posY
            LogWrite "        BEFORE: " & swIsometricView.Position(0) & ", " & swIsometricView.Position(1)
            
            pos(0) = posX
            pos(1) = posY
            swIsometricView.Position = pos
            
            verifyPos = swIsometricView.Position
            LogWrite "        AFTER: " & verifyPos(0) & ", " & verifyPos(1)
            
            Dim newOutline3 As Variant
            newOutline3 = swIsometricView.GetOutline
            Dim newCenterX3 As Double, newCenterY3 As Double
            newCenterX3 = newOutline3(0) + ((newOutline3(2) - newOutline3(0)) / 2)
            newCenterY3 = newOutline3(1) + ((newOutline3(3) - newOutline3(1)) / 2)
            LogWrite "        New center: " & newCenterX3 & ", " & newCenterY3
        End If
    End If
    LogWrite ""
    
    ' Final rebuild
    LogWrite "      Final rebuild after positioning..."
    swDrawModel.ForceRebuild3 False
    
    If Not swTopView Is Nothing Then
        verifyPos = swTopView.Position
        LogWrite "      Top after rebuild: " & verifyPos(0) & ", " & verifyPos(1)
    End If
    If Not swFrontView Is Nothing Then
        verifyPos = swFrontView.Position
        LogWrite "      Front after rebuild: " & verifyPos(0) & ", " & verifyPos(1)
    End If
    If Not swIsometricView Is Nothing Then
        verifyPos = swIsometricView.Position
        LogWrite "      Isometric after rebuild: " & verifyPos(0) & ", " & verifyPos(1)
    End If
    LogWrite ""
    
    LogWrite "      Views positioned and scaled"
    
    ' Auto-arrange dimensions across all three views (same logic used
    ' for component drawings) so dimensions don't overlap/crowd views
    ' after the positions/scale above were set.
    LogWrite "      Auto-arranging dimensions..."
    Dim dimSelectedCount As Long
    Dim dimErrorText As String
    Dim autoArrangeOK As Boolean
    autoArrangeOK = AutoArrangeMainAssemblyDimensions(swDrawModel, swTopView, swFrontView, swIsometricView, dimSelectedCount, dimErrorText)
    LogWrite "      Dimensions selected: " & dimSelectedCount
    
    ' Final rebuild after dimension arrangement
    LogWrite "      Final rebuild after dimension arrangement..."
    swDrawModel.ForceRebuild3 False
    
    On Error GoTo 0
    
End Sub


'============================================================
' AUTO ARRANGE MAIN ASSEMBLY DIMENSIONS (Top/Front/Isometric)
'
' Same technique as AutoArrangeViewDimensions in pdfcomponents.bas,
' extended to the main assembly's three views.
'============================================================

Function AutoArrangeMainAssemblyDimensions( _
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

    ' VIEW 1 (Top)
    If Not swView1 Is Nothing Then
        Set swAnn = swView1.GetFirstAnnotation3
        Do While Not swAnn Is Nothing
            If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
                Set swDispDim = swAnn.GetSpecificAnnotation
                If Not swDispDim Is Nothing Then
                    If IsOrdinateDimension(swDispDim) Then
                        On Error Resume Next
                        jogOK = swDispDim.AutoJogOrdinate()
                        On Error GoTo 0
                    End If
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

    ' VIEW 2 (Front)
    If Not swView2 Is Nothing Then
        Set swAnn = swView2.GetFirstAnnotation3
        Do While Not swAnn Is Nothing
            If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
                Set swDispDim = swAnn.GetSpecificAnnotation
                If Not swDispDim Is Nothing Then
                    If IsOrdinateDimension(swDispDim) Then
                        On Error Resume Next
                        jogOK = swDispDim.AutoJogOrdinate()
                        On Error GoTo 0
                    End If
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

    ' VIEW 3 (Isometric)
    If Not swView3 Is Nothing Then
        Set swAnn = swView3.GetFirstAnnotation3
        Do While Not swAnn Is Nothing
            If swAnn.GetType = swAnnotationType_e.swDisplayDimension Then
                Set swDispDim = swAnn.GetSpecificAnnotation
                If Not swDispDim Is Nothing Then
                    If IsOrdinateDimension(swDispDim) Then
                        On Error Resume Next
                        jogOK = swDispDim.AutoJogOrdinate()
                        On Error GoTo 0
                    End If
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
        AutoArrangeMainAssemblyDimensions = True
    Else
        AutoArrangeMainAssemblyDimensions = False
    End If

End Function


'============================================================
' GET MAIN ASSEMBLY VIEW - By position (1st=Top, 2nd=Front, 3rd=Iso)
'============================================================

Function GetMainAssemblyView(swDraw As SldWorks.DrawingDoc, viewIndex As Long) As SldWorks.View
    
    Dim swView As SldWorks.View
    Dim viewCount As Long
    
    Set GetMainAssemblyView = Nothing
    viewCount = 0
    
    On Error Resume Next
    
    Set swView = swDraw.GetFirstView
    
    Do While Not swView Is Nothing
        If Not swView.ReferencedDocument Is Nothing Then
            viewCount = viewCount + 1
            If viewCount = viewIndex Then
                Set GetMainAssemblyView = swView
                Exit Function
            End If
        End If
        Set swView = swView.GetNextView
    Loop
    
    On Error GoTo 0
    
End Function


'============================================================
' DETERMINE MAIN ASSEMBLY SCALE - Based on C/S dimension
'============================================================

Sub DetermineMainAssemblyScale(maxDimension As Double, _
                               ByRef scaleNumerator As Double, ByRef scaleDenominator As Double)
    
    ' Scale thresholds based on maximum dimension (C or S) in inches
    ' Pattern: Every 40-unit increase adds 10 to denominator
    ' Formula: scaleDenominator = 15 + INT((maxDimension - 0.01) / 40) * 10
    
    scaleNumerator = 1
    scaleDenominator = 15 + Int((maxDimension - 0.01) / 40) * 10
    
End Sub


'============================================================
' EXTRACT DIMENSIONS FROM MAIN ASSEMBLY FILENAME
' Example: 40-3-C50000-S50000-2400-21
' Extracts: C value, S value, and height value
'============================================================

Function ExtractMainAssemblyDimensions(baseName As String, ByRef foundC As Double, _
                                       ByRef foundS As Double, ByRef foundHeight As Double) As Boolean
    
    Dim fileParts() As String
    Dim i As Long
    Dim part As String
    Dim partUpper As String
    Dim numText As String
    Dim cFound As Boolean
    Dim sFound As Boolean
    Dim heightFound As Boolean
    
    foundC = 0#
    foundS = 0#
    foundHeight = 0#
    ExtractMainAssemblyDimensions = False
    cFound = False
    sFound = False
    heightFound = False
    
    fileParts = Split(baseName, "-")
    
    ' Extract C, S, and first standalone 4-digit number (height)
    For i = 0 To UBound(fileParts)
        part = fileParts(i)
        partUpper = UCase(part)
        
        ' Look for C dimension
        If Not cFound And Left(partUpper, 1) = "C" Then
            numText = Mid(part, 2)
            If IsNumeric(numText) Then
                foundC = CDbl(numText) / 100#
                cFound = True
            End If
        End If
        
        ' Look for S dimension (but not SL/SC/SR)
        If Not sFound And Left(partUpper, 1) = "S" And Len(part) > 1 Then
            If Not (Left(partUpper, 2) = "SL" Or Left(partUpper, 2) = "SC" Or Left(partUpper, 2) = "SR") Then
                numText = Mid(part, 2)
                If IsNumeric(numText) Then
                    foundS = CDbl(numText) / 100#
                    sFound = True
                End If
            End If
        End If
        
        ' Look for height (4-digit number)
        If Not heightFound And Len(part) = 4 And IsNumeric(part) Then
            foundHeight = CDbl(part) / 100#
            heightFound = True
        End If
        
        ' Exit early if we found all three
        If cFound And sFound And heightFound Then
            Exit For
        End If
    Next i
    
    ' Success if we found both C and S values
    ExtractMainAssemblyDimensions = (cFound And sFound And foundHeight > 0)
    
End Function


'============================================================
' EXTRACT DIMENSIONS FROM FLEXIBLE FILENAMES
'============================================================

Function ExtractDimensions(baseName As String, ByRef foundLength As Double, ByRef foundHeight As Double) As Boolean

     Dim fileParts() As String
     Dim i As Long
     Dim part As String
     Dim partUpper As String
     Dim prefix As String
     Dim numText As String
     Dim isGusset As Boolean
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
             prefix = Left(partUpper, 1)
             numText = Mid(part, 2)
             If IsNumeric(numText) Then
                 numericValues(numericCount) = CDbl(numText) / 100#
                 numericCount = numericCount + 1
             End If
         ElseIf IsNumeric(part) Then
             numericValues(numericCount) = CDbl(part) / 100#
             numericCount = numericCount + 1
         End If
     Next i

     ' No dimensions found
     If numericCount = 0 Then
         Exit Function
     End If

     ' For gussets: first numeric value is height
     If isGusset Then
         foundHeight = numericValues(0)
         ExtractDimensions = True
         Exit Function
     End If

     ' For other components with multiple dimensions
     If numericCount >= 2 Then
         foundLength = numericValues(0)
         foundHeight = numericValues(1)
     ElseIf numericCount = 1 Then
         foundHeight = numericValues(0)
     End If

     ExtractDimensions = (foundHeight > 0 Or foundLength > 0)

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
' HELPER FUNCTIONS
'============================================================

Function GetDrawingBaseName(fullPath As String) As String
    
    Dim fileName As String
    Dim baseName As String
    
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    If InStrRev(fileName, ".") > 0 Then
        baseName = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        baseName = fileName
    End If
    
    GetDrawingBaseName = baseName
    
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


'============================================================
' LOG WRITE - Append message to log file
'============================================================

Sub LogWrite(message As String)

     Dim fileNum As Integer
     Dim timestamp As String

     fileNum = FreeFile

     On Error Resume Next

     ' Append to log file
     Open logFilePath For Append As #fileNum
     Print #fileNum, message
     Close #fileNum

     On Error GoTo 0

End Sub


'============================================================
' OPEN LOG FILE - Opens log file in default text editor
'============================================================

Sub OpenLogFile()

     Dim shell As Object

     On Error Resume Next

     Set shell = CreateObject("WScript.Shell")
     shell.Run "notepad.exe " & logFilePath, 1, False

     On Error GoTo 0

End Sub


