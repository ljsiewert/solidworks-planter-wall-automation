Option Explicit

' Code-behind source: paste into a new UserForm named UserForm_AutomationUI.
' Controls are created at runtime; no designer layout or .frx is required.
Private WithEvents runButton As MSForms.CommandButton
Private WithEvents cancelButton As MSForms.CommandButton
Private WithEvents browseButton As MSForms.CommandButton
Private WithEvents wallsBox As MSForms.ComboBox
Private WithEvents quoteBox As MSForms.CheckBox

Private sideBox As MSForms.ComboBox
Private returnBox As MSForms.ComboBox
Private materialBox As MSForms.ComboBox
Private thicknessBox As MSForms.ComboBox
Private lengthBox As MSForms.TextBox
Private widthBox As MSForms.TextBox
Private heightBox As MSForms.TextBox
Private centerBox As MSForms.TextBox
Private sideLengthBox As MSForms.TextBox
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
    Me.Height = 545
    acceptedValue = False

    Set wallsBox = AddCombo("cboWallCount", "Wall count", 16, Array("1", "2", "3", "4"))
    Set sideBox = AddCombo("cboSide", "Side orientation", 46, Array("N/A"))
    Set returnBox = AddCombo("cboReturnType", "Return type", 76, _
        Array("No return", "Single return 1 inch", "Single return 2 inches", "Double return"))
    Set lengthBox = AddText("txtOverallLength", "Overall length (in)", 106, "")
    Set widthBox = AddText("txtOverallWidth", "Overall width (in)", 136, "")
    Set heightBox = AddText("txtOverallHeight", "Overall height (in)", 166, "")
    Set centerBox = AddText("txtCenterLength", "Center planter length (in)", 196, "Calculated by design table")
    centerBox.Locked = True
    centerBox.TabStop = False
    Set sideLengthBox = AddText("txtSideLength", "Side planter length (in)", 226, "Calculated by design table")
    sideLengthBox.Locked = True
    sideLengthBox.TabStop = False
    Set materialBox = AddCombo("cboMaterial", "Material", 256, Array("Mild Steel", "Borcon Weathering Steel"))
    Set thicknessBox = AddCombo("cboThickness", "Material thickness", 286, Array("3/16 inch", "1/4 inch"))
    Set folderBox = AddText("txtOutputFolder", "Output parent folder", 316, packngo1.packTestFolder)
    folderBox.Width = 220
    Set browseButton = AddButton("cmdBrowse", "...", 403, 316, 35)

    Set quoteBox = AddCheck("chkQuoteOnly", "Quote only (assembly PDFs only)", 16, 348)
    Set pdfAssembliesBox = AddCheck("chkPDFAssemblies", "PDF assemblies", 16, 380)
    Set pdfComponentsBox = AddCheck("chkPDFComponents", "PDF components", 235, 380)
    Set dxfBox = AddCheck("chkDXF", "DXF files", 16, 406)
    Set stepBox = AddCheck("chkSTEP", "STEP files", 235, 406)
    pdfAssembliesBox.Value = True

    Set note = Me.Controls.Add("Forms.Label.1", "lblNotes", True)
    note.Caption = "New output only; no overwrite. A PACK_TEST subfolder is added if needed." & vbCrLf & _
                   "Selected export macros may still show status/error dialogs."
    note.Left = 16
    note.Top = 438
    note.Width = 425
    note.Height = 32

    Set runButton = AddButton("cmdRun", "Run", 250, 478, 90)
    runButton.Default = True
    Set cancelButton = AddButton("cmdCancel", "Cancel", 348, 478, 90)
    cancelButton.Cancel = True

    wallsBox.ListIndex = -1
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
        pdfAssembliesBox.Value = True
        pdfComponentsBox.Value = False
        dxfBox.Value = False
        stepBox.Value = False
    End If
    pdfAssembliesBox.Enabled = Not quoteBox.Value
    pdfComponentsBox.Enabled = Not quoteBox.Value
    dxfBox.Enabled = Not quoteBox.Value
    stepBox.Enabled = Not quoteBox.Value
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
    Call UIOutputRoot(selectedFolder)

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
    exportPDFAssemblies = quoteOnly Or CBool(pdfAssembliesBox.Value)
    exportPDFComponents = Not quoteOnly And CBool(pdfComponentsBox.Value)
    exportDXF = Not quoteOnly And CBool(dxfBox.Value)
    exportSTEP = Not quoteOnly And CBool(stepBox.Value)
    UICheckExportFiles
    acceptedValue = True
    Me.Hide
    Exit Sub
InvalidInput:
    MsgBox Err.Description, vbExclamation, "Review Form Inputs"
End Sub

Private Function ReadDimension(ByVal box As MSForms.TextBox, ByVal label As String) As Double
    On Error GoTo InvalidInput
    ReadDimension = UIParseInches(box.Text)
    Exit Function
InvalidInput:
    box.SetFocus
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
