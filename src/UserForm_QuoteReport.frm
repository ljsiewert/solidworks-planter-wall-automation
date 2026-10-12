Option Explicit

' Code-behind source: paste into a new UserForm named UserForm_QuoteReport.
Private reportBox As MSForms.TextBox
Private summaryBox As MSForms.TextBox
Private componentsBox As MSForms.ListBox
Private gussetsBox As MSForms.ListBox
Private totalsBox As MSForms.ListBox
Private columnLabels As Collection
Private showingDetails As Boolean
Private WithEvents detailsButton As MSForms.CommandButton
Private WithEvents closeButton As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim note As MSForms.Label
    Dim heading As MSForms.Label
    Dim captions As Variant
    Dim positions As Variant
    Dim widths As Variant
    Dim i As Long
    Me.Caption = "Planter Material Estimate"
    Me.Width = 780
    Me.Height = 570
    Set summaryBox = Me.Controls.Add("Forms.TextBox.1", "txtQuoteSummary", True)
    With summaryBox
        .Left = 12
        .Top = 12
        .Width = 744
        .Height = 165
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
        .Font.Size = 11
        .TabStop = False
    End With
    Set columnLabels = New Collection
    captions = Array("Component", "Qty", "Length (in)", "Width/height (in)", "Total wt (lb)", "Material cost")
    positions = Array(12, 212, 262, 362, 492, 602)
    widths = Array(200, 50, 100, 130, 110, 150)
    For i = 0 To 5
        Set heading = Me.Controls.Add("Forms.Label.1", "lblColumn" & CStr(i), True)
        heading.Caption = captions(i)
        heading.Left = positions(i)
        heading.Top = 182
        heading.Width = widths(i)
        heading.Height = 20
        heading.Font.Bold = True
        columnLabels.Add heading
    Next i
    Set componentsBox = Me.Controls.Add("Forms.ListBox.1", "lstComponents", True)
    With componentsBox
        .Left = 12
        .Top = 204
        .Width = 744
        .Height = 104
        .ColumnCount = 6
        .ColumnWidths = "200;50;100;130;110;130"
        .Font.Size = 10
    End With
    captions = Array("Gusset", "Qty", "Height (in)", "Top width (in)", "Base width (in)", "Total wt (lb)", "Material cost")
    positions = Array(12, 192, 242, 332, 422, 512, 622)
    widths = Array(180, 50, 90, 90, 90, 110, 130)
    For i = 0 To 6
        Set heading = Me.Controls.Add("Forms.Label.1", "lblGussetColumn" & CStr(i), True)
        heading.Caption = captions(i)
        heading.Left = positions(i)
        heading.Top = 316
        heading.Width = widths(i)
        heading.Height = 20
        heading.Font.Bold = True
        columnLabels.Add heading
    Next i
    Set gussetsBox = Me.Controls.Add("Forms.ListBox.1", "lstGussets", True)
    With gussetsBox
        .Left = 12
        .Top = 338
        .Width = 744
        .Height = 40
        .ColumnCount = 7
        .ColumnWidths = "180;50;90;90;90;110;110"
        .Font.Size = 10
    End With
    Set totalsBox = Me.Controls.Add("Forms.ListBox.1", "lstMaterialTotal", True)
    With totalsBox
        .Left = 12
        .Top = 382
        .Width = 744
        .Height = 22
        .ColumnCount = 3
        .ColumnWidths = "490;110;120"
        .Font.Bold = True
        .Font.Size = 10
    End With
    Set reportBox = Me.Controls.Add("Forms.TextBox.1", "txtQuoteReport", True)
    With reportBox
        .Left = 12
        .Top = 12
        .Width = 744
        .Height = 382
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
        .Visible = False
    End With
    Set note = Me.Controls.Add("Forms.Label.1", "lblEstimateNotes", True)
    note.Left = 12
    note.Top = 412
    note.Width = 744
    note.Height = 70
    note.Caption = "Material only - not a final manufacturing quote. Sheet offcuts, fabrication, labor and margin excluded." & vbCrLf & _
        "Holes/reliefs not deducted. Flat heights use calibrated measurements; other heights are predictions." & vbCrLf & _
        "Both steels use an estimating density of 7850 kg/m^3; BORCON-specific density is unverified." & vbCrLf & _
        "All table weights/costs are quantity totals. Combined total includes walls, base plates, ends and gussets."
    Set detailsButton = Me.Controls.Add("Forms.CommandButton.1", "cmdDetails", True)
    With detailsButton
        .Caption = "Technical Details"
        .Left = 12
        .Top = 484
        .Width = 130
        .Height = 26
    End With
    Set closeButton = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With closeButton
        .Caption = "Close"
        .Left = 666
        .Top = 484
        .Width = 90
        .Height = 26
        .Default = True
        .Cancel = True
    End With
End Sub

Public Sub SetResult(ByVal result As UIQuoteResult)
    summaryBox.Text = UIQuoteSummary(result)
    componentsBox.List = UIQuoteComponentRows(result)
    gussetsBox.List = UIQuoteGussetRows(result)
    totalsBox.List = UIQuoteTotalRows(result)
    reportBox.Text = UIQuoteReport(result)
End Sub

Private Sub detailsButton_Click()
    Dim heading As MSForms.Label
    showingDetails = Not showingDetails
    summaryBox.Visible = Not showingDetails
    componentsBox.Visible = Not showingDetails
    gussetsBox.Visible = Not showingDetails
    totalsBox.Visible = Not showingDetails
    For Each heading In columnLabels
        heading.Visible = Not showingDetails
    Next heading
    reportBox.Visible = showingDetails
    If showingDetails Then
        detailsButton.Caption = "Back to Summary"
    Else
        detailsButton.Caption = "Technical Details"
    End If
End Sub

Private Sub closeButton_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
