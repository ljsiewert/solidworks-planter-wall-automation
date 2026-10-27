VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "UIDesignTableSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TABLE_ERROR As Long = vbObjectError + 2200
Private Const INPUT_COUNT As Long = 10
Private Const OUTPUT_COUNT As Long = 13
Private Const LABEL_COLUMN As Long = 16
Private Const VALUE_COLUMN As Long = 17
Private model As SldWorks.ModelDoc2
Private table As SldWorks.DesignTable
Private sheet As Object
Private protectedFormulas As Variant
Private configurationName As String
Private saveGuard As UIReferenceSaveGuard
Private workbookWindows As Collection
Private windowVisibility As Collection

Public Sub SetSaveGuard(ByVal guard As UIReferenceSaveGuard)
    Set saveGuard = guard
End Sub

Public Sub Apply(ByVal document As SldWorks.ModelDoc2, ByVal inputs As Variant, _
                 ByRef outputs As Variant)
    Dim i As Long
    Dim units As Variant

    Set model = document
    configurationName = model.ConfigurationManager.ActiveConfiguration.Name
    units = model.GetUnits
    If units(0) <> swLengthUnit_e.swINCHES Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", "Set the template's document length units to inches before running UI mode."
    End If
    Set table = model.GetDesignTable
    If table Is Nothing Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", "The template has no design table."
    End If
    If table.LinkToFile Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", "This adapter requires an embedded, unlinked design table."
    End If
    If table.IsActive Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", "Close the design table editor before running UI mode."
    End If
    Rebuild "checking the original template (no inputs have been changed)"
    OpenSheet
    ValidateLayout
    UITraceBegin "Excel initial Calculate"
    sheet.Calculate
    UITraceEnd "Excel initial Calculate"
    protectedFormulas = Array(sheet.Range("C3:O3").Formula, sheet.Range("Q10:Q11").Formula)
    ' Prevent Excel from interpreting the fractional thickness label as a date.
    sheet.Range("Q8").NumberFormat = "@"
    For i = 0 To INPUT_COUNT - 1
        If InputIsWritten(inputs, i) Then sheet.Cells(InputRow(i), VALUE_COLUMN).Value2 = inputs(i)
    Next i
    UITraceBegin "Excel input Calculate"
    sheet.Calculate
    UITraceEnd "Excel input Calculate"
    VerifyFormulas
    VerifyInputs inputs
    outputs = ReadOutputs()
    CloseSheet swDesignTableUpdateOptions_e.swUpdateDesignTableAll
    Rebuild "applying the design table inputs (pass 1 of 2)"
    Rebuild "applying the design table inputs (pass 2 of 2)"
End Sub

Public Sub CloseEditor()
    RestoreWorkbookWindows
    If table Is Nothing Then Exit Sub
    If table.IsActive Then CloseSheet swDesignTableUpdateOptions_e.swUpdateDesignTableNone
End Sub

Private Function InputRow(ByVal index As Long) As Long
    If index <= 6 Then InputRow = index + 3 Else InputRow = index + 6
End Function

' The new baseplate width is only sent when the override is YES; otherwise the sheet value is left alone.
Private Function InputIsWritten(ByVal inputs As Variant, ByVal index As Long) As Boolean
    InputIsWritten = (index <> 9) Or (CStr(inputs(8)) = "YES")
End Function

Private Sub OpenSheet()
    UITraceBegin "Open embedded design table"
    If table.IsActive Then
        Set sheet = table.Worksheet
    Else
        Set sheet = table.EditTable2(True)
    End If
    If sheet Is Nothing Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", "SolidWorks could not open the embedded Excel design table."
    End If
    HideWorkbookWindows
    UITraceEnd "Open embedded design table"
End Sub

Private Sub HideWorkbookWindows()
    Dim window As Object
    Set workbookWindows = New Collection
    Set windowVisibility = New Collection
    For Each window In sheet.Parent.Windows
        workbookWindows.Add window
        windowVisibility.Add CBool(window.Visible)
        window.Visible = False
    Next window
End Sub

Private Sub RestoreWorkbookWindows()
    Dim window As Object
    Dim index As Long
    If workbookWindows Is Nothing Then Exit Sub
    Do While workbookWindows.Count > 0
        index = workbookWindows.Count
        Set window = workbookWindows(index)
        window.Visible = CBool(windowVisibility(index))
        workbookWindows.Remove index
        windowVisibility.Remove index
    Loop
    Set workbookWindows = Nothing
    Set windowVisibility = Nothing
End Sub

Private Sub CloseSheet(ByVal updateOption As Long)
    ' Restore and release workbook windows before SolidWorks closes its Excel editor.
    RestoreWorkbookWindows
    If Not saveGuard Is Nothing Then saveGuard.ProtectLoadedReferences
    ' Release the Excel worksheet before SolidWorks closes the table's Excel window.
    Set sheet = Nothing
    UITraceBegin "Commit and close embedded design table"
    If table.UpdateTable(updateOption, True) = False Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", "SolidWorks could not commit/close the design table. Review the open table before retrying."
    End If
    UITraceEnd "Commit and close embedded design table"
End Sub

Private Sub ValidateLayout()
    Dim headers As Variant
    Dim labels As Variant
    Dim i As Long
    Dim address As Variant

    headers = Array("$DESCRIPTION", "$VALUE@SIDE@Equations", "$VALUE@WALL_COUNT@Equations", _
        "$VALUE@OVERALL_LENGTH@Equations", "$VALUE@OVERALL_WIDTH@Equations", _
        "$VALUE@OVERALL_HEIGHT@Equations", "$VALUE@RETURN_TYPE@Equations", _
        "$VALUE@CENTER_PLANTER_LENGTH@Equations", "$VALUE@SIDE_PLANTER_LENGTH@Equations", _
        "$VALUE@OVERALL_THICKNESS@Equations", "$VALUE@OVERALL_MATERIAL@Equations", _
        "$VALUE@DEPTH_OVERRIDE@Equations", "$VALUE@CONTINUATION@Equations", _
        "$VALUE@DEPTH_OVERRIDE_DIMENSION@Equations")
    labels = Array("WALL COUNT:", "RETURN TYPE:", "OVERALL LENGTH:", "OVERALL WIDTH:", _
                   "OVERALL HEIGHT:", "OVERALL THICKNESS:", "OVERALL MATERIAL:", _
                   "DOES THIS PLANTER MOUNT TO ANOTHER PLANTER?", "DO YOU WANT TO OVERRIDE BASEPLATE WIDTH?", _
                   "BASEPLATE WIDTH NEW VALUE:")
    For i = 0 To UBound(headers)
        If StrComp(CStr(sheet.Cells(2, i + 2).Value2), CStr(headers(i)), vbTextCompare) <> 0 Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Unexpected design table header in row 2, column " & CStr(i + 2)
        End If
    Next i
    For i = 0 To INPUT_COUNT - 1
        If Squash(CStr(sheet.Cells(InputRow(i), LABEL_COLUMN).Value2)) <> labels(i) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Unexpected input label at P" & CStr(InputRow(i))
        End If
        If sheet.Cells(InputRow(i), VALUE_COLUMN).HasFormula Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Input Q" & CStr(InputRow(i)) & " is a formula; refusing to overwrite it."
        End If
    Next i
    If StrComp(CStr(sheet.Range("A3").Value2), configurationName, vbTextCompare) <> 0 Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", _
            "Select the design table's configuration before running: " & CStr(sheet.Range("A3").Value2)
    End If
    For Each address In Array("C3", "D3", "E3", "F3", "G3", "H3", "I3", "J3", "K3", "L3", "M3", "N3", "O3", "Q10", "Q11")
        If Not sheet.Range(CStr(address)).HasFormula Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Expected a calculated formula at " & CStr(address)
        End If
    Next address
End Sub

Private Function Squash(ByVal text As String) As String
    text = Trim$(text)
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    Squash = text
End Function

Private Sub VerifyFormulas()
    Dim actual As Variant
    Dim saved As Variant
    Dim i As Long

    actual = sheet.Range("C3:O3").Formula
    saved = protectedFormulas(0)
    For i = 1 To 13
        If actual(1, i) <> saved(1, i) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "A design table output formula changed."
        End If
    Next i
    actual = sheet.Range("Q10:Q11").Formula
    saved = protectedFormulas(1)
    For i = 1 To 2
        If actual(i, 1) <> saved(i, 1) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "A planter length formula changed."
        End If
    Next i
End Sub

Private Sub VerifyInputs(ByVal expected As Variant)
    Dim i As Long
    For i = 0 To INPUT_COUNT - 1
        If InputIsWritten(expected, i) Then
            If CStr(sheet.Cells(InputRow(i), VALUE_COLUMN).Value2) <> CStr(expected(i)) Then
                Err.Raise TABLE_ERROR, "UIDesignTableSession", "Excel did not retain input Q" & CStr(InputRow(i))
            End If
        End If
    Next i
End Sub

Private Function ReadOutputs() As Variant
    Dim addresses As Variant
    Dim values(0 To OUTPUT_COUNT - 1) As Double
    Dim i As Long

    addresses = Array("D3", "C3", "H3", "E3", "F3", "G3", "I3", "J3", "L3", "K3", "M3", "N3", "O3")
    For i = 0 To OUTPUT_COUNT - 1
        values(i) = UIReadTableNumber(sheet.Range(CStr(addresses(i))).Value2, CStr(addresses(i)))
    Next i
    ReadOutputs = values
End Function

Private Sub Rebuild(ByVal stage As String)
    Dim activeName As String
    If Not saveGuard Is Nothing Then saveGuard.ProtectLoadedReferences
    activeName = model.ConfigurationManager.ActiveConfiguration.Name
    If StrComp(activeName, configurationName, vbTextCompare) <> 0 Then
        If model.ShowConfiguration2(configurationName) = False Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", _
                "Cannot activate configuration: " & configurationName & vbCrLf & _
                "Current configuration: " & model.ConfigurationManager.ActiveConfiguration.Name
        End If
    End If
    activeName = model.ConfigurationManager.ActiveConfiguration.Name
    If StrComp(activeName, configurationName, vbTextCompare) <> 0 Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", _
            "Configuration activation did not select " & configurationName & _
            ". Current configuration: " & activeName
    End If
    ' SolidWorks can return True as 1; VBA Not would invert its bits, not test for zero.
    UITraceBegin "Template ForceRebuild3: " & stage
    If model.ForceRebuild3(False) = False Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", _
            "The model failed to rebuild after " & stage & ". Check feature/component errors in SolidWorks."
    End If
    UITraceEnd "Template ForceRebuild3: " & stage
End Sub
