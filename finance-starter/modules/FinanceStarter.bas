Attribute VB_Name = "FinanceStarter"
Option Explicit

Public Sub InvoiceTotals()
    Dim ws As Worksheet
    Dim summary As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Invoices")
    Set summary = ThisWorkbook.Worksheets("Summary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' TODO: loop from row 2 to lastRow.
    ' Column E = invoice amount; column F = paid amount.
    ' Use Currency variables and write counts/totals to Summary.
End Sub
