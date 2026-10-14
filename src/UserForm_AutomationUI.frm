Option Explicit

' Code-behind source: paste into a new UserForm named UserForm_AutomationUI.
' Controls are created at runtime; no designer layout or .frx is required.
Private WithEvents runButton As MSForms.CommandButton
Private WithEvents cancelButton As MSForms.CommandButton
Private WithEvents browseButton As MSForms.CommandButton
Private WithEvents wallsBox As MSForms.ComboBox
Private WithEvents quoteBox As MSForms.CheckBox
Private WithEvents previewBox As MSForms.CheckBox
Private WithEvents propertiesButton As MSForms.CommandButton
Private WithEvents propertiesDoneButton As MSForms.CommandButton
Private WithEvents propertiesResetButton As MSForms.CommandButton
Private WithEvents paintedToggle As MSForms.ToggleButton

Private sideBox As MSForms.ComboBox
Private returnBox As MSForms.ComboBox
Private materialBox As MSForms.ComboBox
Private thicknessBox As MSForms.ComboBox
Private lengthBox As MSForms.TextBox
Private widthBox As MSForms.TextBox
Private heightBox As MSForms.TextBox
Private priceBox As MSForms.TextBox
Private gapBox As MSForms.TextBox
Private marginBox As MSForms.TextBox
Private cutRateBox As MSForms.TextBox
Private bendRateBox As MSForms.TextBox
Private weldRateBox As MSForms.TextBox
Private paintRateBox As MSForms.TextBox
Private quoteMarginBox As MSForms.TextBox
Private propertiesFrame As MSForms.Frame
Private folderBox As MSForms.TextBox
Private pdfAssembliesBox As MSForms.CheckBox
Private pdfComponentsBox As MSForms.CheckBox
Private dxfBox As MSForms.CheckBox
Private stepBox As MSForms.CheckBox
Private acceptedValue As Boolean

Public Property Get Accepted() As Boolean
    Accepted = acceptedValue
End Property

Private Sub UserForm_Initialize()
    Dim note As MSForms.Label
    Dim index As Long

    Me.Caption = "Planter Wall Automation"
    Me.Width = 470
    Me.Height = 515
    acceptedValue = False

    Set wallsBox = AddCombo("cboWallCount", "Wall count", 16, Array("1", "2", "3", "4"))
    Set sideBox = AddCombo("cboSide", "Side orientation", 46, Array("N/A"))
    Set returnBox = AddCombo("cboReturnType", "Return type", 76, _
        Array("No return", "Single return 1 inch", "Single return 2 inches", "Double return"))
    Set lengthBox = AddText("txtOverallLength", "Overall length (in)", 106, "")
    Set widthBox = AddText("txtOverallWidth", "Overall width (in)", 136, "")
    Set heightBox = AddText("txtOverallHeight", "Overall height (in)", 166, "")
    Set materialBox = AddCombo("cboMaterial", "Material", 196, Array("Mild Steel", "Borcon Weathering Steel"))
    materialBox.Width = 178
    Set paintedToggle = Me.Controls.Add("Forms.ToggleButton.1", "tglPainted", True)
    With paintedToggle
        .Left = 358
        .Top = 196
        .Width = 80
        .Height = 22
        .Value = False
        .ControlTipText = "Include paint (both sides of every part) in the quote."
    End With
    paintedToggle_Click
    Set thicknessBox = AddCombo("cboThickness", "Material thickness", 226, Array("3/16 inch", "1/4 inch"))
    Set folderBox = AddText("txtOutputFolder", "Output parent folder", 256, packngo1.packTestFolder)
    folderBox.Width = 220
    Set browseButton = AddButton("cmdBrowse", "...", 403, 256, 35)

    Set previewBox = AddCheck("chkPreviewOnly", "Preview only (no files created)", 16, 288)
    Set quoteBox = AddCheck("chkQuoteOnly", "Quote only (input-only material estimate)", 16, 318)
    Set pdfAssembliesBox = AddCheck("chkPDFAssemblies", "PDF assemblies", 16, 350)
    Set pdfComponentsBox = AddCheck("chkPDFComponents", "PDF components", 235, 350)
    Set dxfBox = AddCheck("chkDXF", "DXF files", 16, 376)
    Set stepBox = AddCheck("chkSTEP", "STEP files", 235, 376)

    Set note = Me.Controls.Add("Forms.Label.1", "lblNotes", True)
    note.Caption = "Quote rates, gap and margins are under Properties; defaults are provisional." & vbCrLf & _
                   "Sheet layouts are estimates, not laser-ready cutting programs."
    note.Left = 16
    note.Top = 408
    note.Width = 425
    note.Height = 32

    Set propertiesButton = AddButton("cmdProperties", "Properties...", 16, 448, 100)
    propertiesButton.ControlTipText = "Quote material price, labor and paint rates, margin, part gap and edge margin"
    Set runButton = AddButton("cmdRun", "Run", 250, 448, 90)
    runButton.Default = True
    Set cancelButton = AddButton("cmdCancel", "Cancel", 348, 448, 90)
    cancelButton.Cancel = True
    BuildProperties

    wallsBox.ListIndex = 0
    returnBox.ListIndex = 0
    materialBox.ListIndex = 0
    thicknessBox.ListIndex = 0
    For index = 0 To wallsBox.ListCount - 1
        If wallsBox.List(index) = packngo1.wallCount Then wallsBox.ListIndex = index
    Next index
    wallsBox_Change
    If packngo1.wallCount = "2" Then
        Select Case packngo1.side
            Case "1": sideBox.Value = "Right"
            Case "2": sideBox.Value = "Left"
        End Select
    End If
    Select Case packngo1.returnType
        Case "0", "1", "2", "3": returnBox.ListIndex = CLng(packngo1.returnType)
    End Select
    Select Case packngo1.overallMaterial
        Case "1", "2": materialBox.ListIndex = CLng(packngo1.overallMaterial) - 1
    End Select
    Select Case packngo1.thicknessLabel
        Case "1", "2": thicknessBox.ListIndex = CLng(packngo1.thicknessLabel) - 1
    End Select
End Sub

' Properties panel: an overlay frame so quote settings stay adjustable without cluttering the form.
Private Sub BuildProperties()
    Dim heading As MSForms.Label
    Set propertiesFrame = Me.Controls.Add("Forms.Frame.1", "fraProperties", True)
    With propertiesFrame
        .Caption = "Quote Properties"
        .Left = 6
        .Top = 6
        .Width = Me.InsideWidth - 12
        .Height = Me.InsideHeight - 12
        .Visible = False
    End With
    Set heading = propertiesFrame.Controls.Add("Forms.Label.1", "lblPropertiesNote", True)
    heading.Caption = "Used by Quote only. Labor and paint rates are placeholders until time studies are done."
    heading.Left = 10
    heading.Top = 8
    heading.Width = 420
    heading.Height = 28
    Set priceBox = AddPropertyText("txtPricePerPound", "Quote material ($/lb)", 40, "0.90")
    Set cutRateBox = AddPropertyText("txtCutRate", "Cutting labor ($/sheet)", 70, FormatRate(UIQuoteDefaultRate("Cut")))
    Set bendRateBox = AddPropertyText("txtBendRate", "Bending labor ($/bent part)", 100, FormatRate(UIQuoteDefaultRate("Bend")))
    Set weldRateBox = AddPropertyText("txtWeldRate", "Welding labor ($/part)", 130, FormatRate(UIQuoteDefaultRate("Weld")))
    Set paintRateBox = AddPropertyText("txtPaintRate", "Paint ($/sq ft, both sides)", 160, FormatRate(UIQuoteDefaultRate("Paint")))
    Set quoteMarginBox = AddPropertyText("txtQuoteMargin", "Default margin (%)", 190, FormatRate(UIQuoteDefaultRate("Margin")))
    Set gapBox = AddPropertyText("txtPartGap", "Quote part gap (in)", 220, "0.25")
    gapBox.ControlTipText = "Operations baseline: 0.25 in; typical range 0-0.25 in. Zero models touching edges, not a shared laser toolpath."
    Set marginBox = AddPropertyText("txtEdgeMargin", "Quote edge margin (in)", 250, "0.25")
    Set heading = propertiesFrame.Controls.Add("Forms.Label.1", "lblPropertiesHelp", True)
    heading.Caption = "Bending applies only when a return is selected (wall and side wall parts). " & _
        "Welding is charged for every part. Margin can also be changed on the results page."
    heading.Left = 10
    heading.Top = 282
    heading.Width = 420
    heading.Height = 40
    heading.ForeColor = RGB(89, 89, 89)
    Set propertiesResetButton = propertiesFrame.Controls.Add("Forms.CommandButton.1", "cmdResetProperties", True)
    With propertiesResetButton
        .Caption = "Restore Defaults"
        .Left = 10
        .Top = propertiesFrame.InsideHeight - 34
        .Width = 120
        .Height = 24
    End With
    Set propertiesDoneButton = propertiesFrame.Controls.Add("Forms.CommandButton.1", "cmdPropertiesDone", True)
    With propertiesDoneButton
        .Caption = "Done"
        .Left = propertiesFrame.InsideWidth - 100
        .Top = propertiesFrame.InsideHeight - 34
        .Width = 90
        .Height = 24
    End With
End Sub

Private Function AddPropertyText(ByVal name As String, ByVal caption As String, _
                                 ByVal top As Single, ByVal text As String) As MSForms.TextBox
    Dim label As MSForms.Label
    Dim box As MSForms.TextBox
    Set label = propertiesFrame.Controls.Add("Forms.Label.1", name & "Label", True)
    label.Caption = caption
    label.Left = 10
    label.Top = top + 3
    label.Width = 170
    label.Height = 19
    Set box = propertiesFrame.Controls.Add("Forms.TextBox.1", name, True)
    box.Left = 190
    box.Top = top
    box.Width = 120
    box.Height = 22
    box.Text = text
    Set AddPropertyText = box
End Function

Private Function FormatRate(ByVal value As Double) As String
    FormatRate = Format$(value, "0.00")
End Function

Private Sub ShowProperties(ByVal showPanel As Boolean)
    Dim item As Object
    For Each item In Me.Controls
        If item.Parent Is Me Then
            If Not item Is propertiesFrame Then item.Visible = Not showPanel
        End If
    Next item
    propertiesFrame.Visible = showPanel
    runButton.Default = Not showPanel
    cancelButton.Cancel = Not showPanel
End Sub

Private Sub propertiesButton_Click()
    ShowProperties True
End Sub

' Values are validated when a quote runs; Done just closes the panel.
Private Sub propertiesDoneButton_Click()
    ShowProperties False
End Sub

Private Sub propertiesResetButton_Click()
    priceBox.Text = "0.90"
    cutRateBox.Text = FormatRate(UIQuoteDefaultRate("Cut"))
    bendRateBox.Text = FormatRate(UIQuoteDefaultRate("Bend"))
    weldRateBox.Text = FormatRate(UIQuoteDefaultRate("Weld"))
    paintRateBox.Text = FormatRate(UIQuoteDefaultRate("Paint"))
    quoteMarginBox.Text = FormatRate(UIQuoteDefaultRate("Margin"))
    gapBox.Text = "0.25"
    marginBox.Text = "0.25"
End Sub

Private Sub paintedToggle_Click()
    If paintedToggle.Value Then
        paintedToggle.Caption = "Painted"
        paintedToggle.ForeColor = RGB(31, 73, 125)
        paintedToggle.Font.Bold = True
    Else
        paintedToggle.Caption = "Not painted"
        paintedToggle.ForeColor = RGB(89, 89, 89)
        paintedToggle.Font.Bold = False
    End If
End Sub

Private Function AddCombo(ByVal name As String, ByVal caption As String, _
                          ByVal top As Single, ByVal choices As Variant) As MSForms.ComboBox
    Dim combo As MSForms.ComboBox
    Dim choice As Variant
    AddLabel name & "Label", caption, top
    Set combo = Me.Controls.Add("Forms.ComboBox.1", name, True)
    combo.Left = 175
    combo.Top = top
    combo.Width = 263
    combo.Height = 22
    combo.Style = fmStyleDropDownList
    For Each choice In choices
        combo.AddItem CStr(choice)
    Next choice
    Set AddCombo = combo
End Function

Private Function AddText(ByVal name As String, ByVal caption As String, _
                         ByVal top As Single, ByVal text As String) As MSForms.TextBox
    Dim box As MSForms.TextBox
    AddLabel name & "Label", caption, top
    Set box = Me.Controls.Add("Forms.TextBox.1", name, True)
    box.Left = 175
    box.Top = top
    box.Width = 263
    box.Height = 22
    box.Text = text
    Set AddText = box
End Function

Private Sub AddLabel(ByVal name As String, ByVal caption As String, ByVal top As Single)
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", name, True)
    label.Caption = caption
    label.Left = 16
    label.Top = top + 3
    label.Width = 155
    label.Height = 19
End Sub

Private Function AddCheck(ByVal name As String, ByVal caption As String, _
                          ByVal left As Single, ByVal top As Single) As MSForms.CheckBox
    Dim box As MSForms.CheckBox
    Set box = Me.Controls.Add("Forms.CheckBox.1", name, True)
    box.Caption = caption
    box.Left = left
    box.Top = top
    box.Width = 215
    box.Height = 22
    box.TripleState = False
    Set AddCheck = box
End Function

Private Function AddButton(ByVal name As String, ByVal caption As String, _
                           ByVal left As Single, ByVal top As Single, _
                           ByVal width As Single) As MSForms.CommandButton
    Dim button As MSForms.CommandButton
    Set button = Me.Controls.Add("Forms.CommandButton.1", name, True)
    button.Caption = caption
    button.Left = left
    button.Top = top
    button.Width = width
    button.Height = 24
    Set AddButton = button
End Function

Private Sub wallsBox_Change()
    If sideBox Is Nothing Then Exit Sub
    sideBox.Clear
    If wallsBox.Value = "2" Then
        sideBox.AddItem "Left"
        sideBox.AddItem "Right"
        sideBox.Enabled = True
        sideBox.ListIndex = -1
    Else
        sideBox.AddItem "N/A"
        sideBox.ListIndex = 0
        sideBox.Enabled = False
    End If
End Sub

Private Sub quoteBox_Click()
    If pdfAssembliesBox Is Nothing Then Exit Sub
    If quoteBox.Value Then
        pdfAssembliesBox.Value = False
        pdfComponentsBox.Value = False
        dxfBox.Value = False
        stepBox.Value = False
        previewBox.Value = False
    End If
    pdfAssembliesBox.Enabled = Not quoteBox.Value And Not previewBox.Value
    pdfComponentsBox.Enabled = Not quoteBox.Value And Not previewBox.Value
    dxfBox.Enabled = Not quoteBox.Value And Not previewBox.Value
    stepBox.Enabled = Not quoteBox.Value And Not previewBox.Value
    previewBox.Enabled = Not quoteBox.Value
    folderBox.Enabled = Not quoteBox.Value And Not previewBox.Value
    browseButton.Enabled = Not quoteBox.Value And Not previewBox.Value
End Sub

Private Sub previewBox_Click()
    If quoteBox Is Nothing Then Exit Sub
    If previewBox.Value Then
        quoteBox.Value = False
        pdfAssembliesBox.Value = False
        pdfComponentsBox.Value = False
        dxfBox.Value = False
        stepBox.Value = False
    End If
    quoteBox.Enabled = Not previewBox.Value
    folderBox.Enabled = Not previewBox.Value
    browseButton.Enabled = Not previewBox.Value
    quoteBox_Click
End Sub

Private Sub browseButton_Click()
    Dim shell As Object
    Dim selected As Object
    On Error GoTo Failed
    Set shell = CreateObject("Shell.Application")
    Set selected = shell.BrowseForFolder(0, "Select the output parent folder", &H1 Or &H40)
    If Not selected Is Nothing Then folderBox.Text = selected.Self.Path
    Exit Sub
Failed:
    MsgBox "Could not open the folder browser. You can type an existing folder instead." & _
           vbCrLf & Err.Description, vbExclamation, "Output Folder"
End Sub

Private Sub runButton_Click()
    Dim length As Double
    Dim width As Double
    Dim height As Double
    Dim selectedFolder As String
    Dim gauge As String
    Dim material As String
    Dim pricePerPound As Double
    Dim partGap As Double
    Dim edgeMargin As Double
    Dim quoteMargin As Double

    On Error GoTo InvalidInput
    If wallsBox.ListIndex < 0 Then
        Err.Raise vbObjectError + 2100, , "Choose a wall count."
    End If
    UIValidateSelections CLng(wallsBox.Value), CStr(sideBox.Value), _
                         CStr(returnBox.Value), CStr(materialBox.Value), CStr(thicknessBox.Value)
    length = ReadDimension(lengthBox, "Overall length")
    width = ReadDimension(widthBox, "Overall width")
    height = ReadDimension(heightBox, "Overall height")
    UIValidateDimensions length, width, height
    selectedFolder = Trim$(folderBox.Text)
    If Not previewBox.Value And Not quoteBox.Value Then Call UIOutputRoot(selectedFolder)

    quoteOnly = CBool(quoteBox.Value)
    previewOnly = CBool(previewBox.Value)
    Set quoteValues = Nothing
    If quoteOnly Then
        On Error GoTo InvalidProperty
        pricePerPound = UIQuoteParsePrice(priceBox.Text)
        partGap = ReadClearance(gapBox, "Part gap")
        edgeMargin = ReadClearance(marginBox, "Sheet edge margin")
        quoteMargin = UIQuoteParseRate(quoteMarginBox.Text, "Default margin")
        If quoteMargin > 95 Then Err.Raise vbObjectError + 2100, , "Default margin must be between 0% and 95%."
        Set quoteValues = UICalculateQuote(CLng(wallsBox.Value), CStr(sideBox.Value), _
            CStr(returnBox.Value), CStr(materialBox.Value), CStr(thicknessBox.Value), length, width, height, pricePerPound, partGap, edgeMargin)
        UIQuotePricing quoteValues, CBool(paintedToggle.Value), UIQuoteParseRate(cutRateBox.Text, "Cutting labor"), _
            UIQuoteParseRate(bendRateBox.Text, "Bending labor"), UIQuoteParseRate(weldRateBox.Text, "Welding labor"), _
            UIQuoteParseRate(paintRateBox.Text, "Paint rate"), quoteMargin
        On Error GoTo InvalidInput
        exportPDFAssemblies = False
        exportPDFComponents = False
        exportDXF = False
        exportSTEP = False
        outputFolder = ""
        acceptedValue = True
        Me.Hide
        Exit Sub
    End If

    UIMapSelections CLng(wallsBox.Value), CStr(sideBox.Value), _
                    CStr(returnBox.Value), CStr(materialBox.Value), CStr(thicknessBox.Value)
    packngo1.overallLength = CStr(length)
    packngo1.overallWidth = CStr(width)
    packngo1.overallHeight = CStr(height)
    gauge = IIf(thicknessBox.Value = "3/16 inch", "3/16", "1/4")
    material = IIf(materialBox.Value = "Mild Steel", "MILD STEEL", "BORCON WEATHERING STEEL")
    designTableInputs = Array(UIWallTableLabel(CLng(wallsBox.Value), CStr(sideBox.Value)), _
        UIReturnTableLabel(CStr(returnBox.Value)), length, width, height, gauge, material)
    outputFolder = selectedFolder
    quoteOnly = CBool(quoteBox.Value)
    exportPDFAssemblies = Not quoteOnly And CBool(pdfAssembliesBox.Value)
    exportPDFComponents = Not quoteOnly And CBool(pdfComponentsBox.Value)
    exportDXF = Not quoteOnly And CBool(dxfBox.Value)
    exportSTEP = Not quoteOnly And CBool(stepBox.Value)
    previewOnly = CBool(previewBox.Value)
    If previewOnly Then
        outputFolder = ""
        UIDisableExports
    End If
    If Not previewOnly Then UICheckExportFiles
    acceptedValue = True
    Me.Hide
    Exit Sub
InvalidInput:
    MsgBox Err.Description, vbExclamation, "Review Form Inputs"
    Exit Sub
InvalidProperty:
    Set quoteValues = Nothing
    ShowProperties True
    MsgBox Err.Description & vbCrLf & "Open Properties to correct the quote settings.", vbExclamation, "Review Quote Properties"
End Sub

Private Function ReadClearance(ByVal box As MSForms.TextBox, ByVal label As String) As Double
    If Trim$(box.Text) = "0" Or Trim$(box.Text) = "0.0" Or Trim$(box.Text) = "0.00" Then
        ReadClearance = 0
    Else
        ReadClearance = ReadDimension(box, label)
    End If
End Function

Private Function ReadDimension(ByVal box As MSForms.TextBox, ByVal label As String) As Double
    On Error GoTo InvalidInput
    ReadDimension = UIParseInches(box.Text)
    Exit Function
InvalidInput:
    If Not box.Parent Is propertiesFrame Then box.SetFocus
    Err.Raise vbObjectError + 2100, "ReadDimension", label & ": " & Err.Description
End Function

Private Sub cancelButton_Click()
    acceptedValue = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        acceptedValue = False
        Me.Hide
    End If
End Sub
