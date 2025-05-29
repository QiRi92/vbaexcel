'module name : PrintTO_PDF
'Print from invoice worksheet into PDF
'https://www.youtube.com/watch?v=eqRJKkwxOzQ&list=PLIBeRriXvKzCB-4ydRujpkkOGUpRJoz9k

Option Explicit

Sub PrintInvoiceToPDF()
Dim PDFName As String, FilePath As String
With Invoice
  PDFName = .Range("E3").Value & "_Invoice#_" & .Range("G3").Value & "_" & Format(now, "MM-DD-YYYY_HH-MM-SS") 'Customer Name & Invoice Number 
  With. PageSetup
    .PrintArea = "D2:H39"
    .LeftMargin = Application.InchesToPoints(0.5)
    .RightMargin = Application.InchesToPoints(0.5)
    .TopMargin = Application.InchesToPoints(0.75)
    .BottomMargin = Application.InchesToPoints(0.75)
    .CenterHorizontally = True
    .CenterVertically = True
  End With
  FilePath = ThisWorkbook.Path & "\" & PDFName & ".pdf"
  .ExportAsFixedFormat xlTypePDF, FileName:=ThisWorkbook.Path & "\" & PDFName & ".pdf", OpenAfterPublish:=True, IgnorePrintAreas:=False
End With
End Sub

'Print from Customers worksheet into PDF

Sub PrintCustomerList()
Dim LastRow As Long
With Customers
    LastRow = .Range("A99999").End(xlUp).Row  'Last Customer Row
    With .PageSetup
        .Orientation = xlLandscape
        .PrintArea = "A1:H" & LastRow
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    
    .ExportAsFixedFormat xlTypePDF, ThisWorkbook.Path & "\CustomerList.pdf", , , False, , , True


End With

End Sub
