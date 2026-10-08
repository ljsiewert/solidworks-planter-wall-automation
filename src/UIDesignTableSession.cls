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
Private model As SldWorks.ModelDoc2
Private table As SldWorks.DesignTable
Private sheet As Object
Private protectedFormulas As Variant
Private configurationName As String
Private saveGuard As UIReferenceSaveGuard

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
    protectedFormulas = Array(sheet.Range("C3:L3").Formula, sheet.Range("N10:N11").Formula)
    ' Prevent Excel from interpreting the fractional thickness label as a date.
    sheet.Range("N8").NumberFormat = "@"
    For i = 0 To 6
        sheet.Cells(i + 3, 14).Value2 = inputs(i)
    Next i
    UITraceBegin "Excel input Calculate"
    sheet.Calculate
    UITraceEnd "Excel input Calculate"
    VerifyFormulas
    VerifyInputs inputs
    outputs = ReadOutputs()
    CloseSheet swDesignTableUpdateOptions_e.swUpdateDesignTableAll
    Rebuild "applying the design table inputs"
End Sub

Public Sub CloseEditor()
    If table Is Nothing Then Exit Sub
    If table.IsActive Then CloseSheet swDesignTableUpdateOptions_e.swUpdateDesignTableNone
End Sub

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
    UITraceEnd "Open embedded design table"
End Sub

Private Sub CloseSheet(ByVal updateOption As Long)
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
        "$VALUE@OVERALL_THICKNESS@Equations", "$VALUE@OVERALL_MATERIAL@Equations")
    labels = Array("WALL COUNT:", "RETURN TYPE:", "OVERALL LENGTH:", "OVERALL WIDTH:", _
                   "OVERALL HEIGHT:", "OVERALL THICKNESS:", "OVERALL MATERIAL:")
    For i = 0 To UBound(headers)
        If StrComp(CStr(sheet.Cells(2, i + 2).Value2), CStr(headers(i)), vbTextCompare) <> 0 Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Unexpected design table header in row 2, column " & CStr(i + 2)
        End If
    Next i
    For i = 0 To 6
        If CStr(sheet.Cells(i + 3, 13).Value2) <> labels(i) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Unexpected input label at M" & CStr(i + 3)
        End If
        If sheet.Cells(i + 3, 14).HasFormula Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Input N" & CStr(i + 3) & " is a formula; refusing to overwrite it."
        End If
    Next i
    If StrComp(CStr(sheet.Range("A3").Value2), configurationName, vbTextCompare) <> 0 Then
        Err.Raise TABLE_ERROR, "UIDesignTableSession", _
            "Select the design table's configuration before running: " & CStr(sheet.Range("A3").Value2)
    End If
    For Each address In Array("C3", "D3", "E3", "F3", "G3", "H3", "I3", "J3", "K3", "L3", "N10", "N11")
        If Not sheet.Range(CStr(address)).HasFormula Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Expected a calculated formula at " & CStr(address)
        End If
    Next address
End Sub

Private Sub VerifyFormulas()
    Dim actual As Variant
    Dim saved As Variant
    Dim i As Long

    actual = sheet.Range("C3:L3").Formula
    saved = protectedFormulas(0)
    For i = 1 To 10
        If actual(1, i) <> saved(1, i) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "A design table output formula changed."
        End If
    Next i
    actual = sheet.Range("N10:N11").Formula
    saved = protectedFormulas(1)
    For i = 1 To 2
        If actual(i, 1) <> saved(i, 1) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "A planter length formula changed."
        End If
    Next i
End Sub

Private Sub VerifyInputs(ByVal expected As Variant)
    Dim i As Long
    For i = 0 To 6
        If CStr(sheet.Cells(i + 3, 14).Value2) <> CStr(expected(i)) Then
            Err.Raise TABLE_ERROR, "UIDesignTableSession", "Excel did not retain input N" & CStr(i + 3)
        End If
    Next i
End Sub

Private Function ReadOutputs() As Variant
    Dim addresses As Variant
    Dim values(0 To 9) As Double
    Dim i As Long

    addresses = Array("D3", "C3", "H3", "E3", "F3", "G3", "I3", "J3", "L3", "K3")
    For i = 0 To 9
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
