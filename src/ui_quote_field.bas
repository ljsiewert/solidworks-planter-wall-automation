Attribute VB_Name = "ui_quote_field"
Option Explicit

Public Sub UIShowFieldQuote(ByVal designs As Collection)
    Dim report As UserForm_FieldQuote
    Dim selected As UIQuoteResult
    Set report = New UserForm_FieldQuote
    report.SetDesigns designs
    Do
        report.Show vbModal
        Set selected = report.PreviewDesign
        If selected Is Nothing Then Exit Do
        UIPreviewQuoteDesign selected
    Loop
    Unload report
End Sub

Public Sub UIFieldReplaceDesign(ByVal designs As Collection, ByVal index As Long, ByVal result As UIQuoteResult)
    If index < 1 Or index > designs.Count Then
        Err.Raise vbObjectError + 2400, "UIFieldReplaceDesign", "Select a saved planter design to update."
    End If
    designs.Add result, Before:=index
    designs.Remove index + 1
End Sub

Public Sub UIFieldSetMargin(ByVal designs As Collection, ByVal margin As Double)
    Dim result As UIQuoteResult
    If margin < 0 Or margin > 95 Then
        Err.Raise vbObjectError + 2400, "UIFieldSetMargin", "Margin must be between 0% and 95%."
    End If
    For Each result In designs
        UIQuoteSetMargin result, margin
    Next result
End Sub

' Costs are only usable as a complete field quote when every design is priced.
' Totals: copies, sheets, material, labor, paint, cost, profit, sell, all-priced.
Public Function UIFieldQuoteTotals(ByVal designs As Collection) As Variant
    Dim totals(0 To 8) As Variant
    Dim result As UIQuoteResult
    Dim i As Long
    If designs.Count = 0 Then
        Err.Raise vbObjectError + 2400, "UIFieldQuoteTotals", "Add at least one planter design."
    End If
    For i = 0 To 7
        totals(i) = 0
    Next i
    totals(8) = True
    For Each result In designs
        totals(0) = totals(0) + result.Quantity
        totals(1) = totals(1) + result.SheetCount
        totals(2) = totals(2) + result.SheetMaterialCost
        totals(3) = totals(3) + result.LaborCost
        totals(4) = totals(4) + result.PaintCost
        totals(5) = totals(5) + result.TotalCost
        totals(6) = totals(6) + result.Profit
        totals(7) = totals(7) + result.SellPrice
        If Not result.HasPricing Then totals(8) = False
    Next result
    UIFieldQuoteTotals = totals
End Function
