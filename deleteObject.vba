Option Explicit
Sub ModifyTicket177329()

    Dim filePath1 As String
    Dim fileName As String
    Dim fileArray1 As Variant
    Dim filePath2 As String
    Dim fileArray2 As Variant
    filePath1 = "C:\bpc_svn\BIN\ABP_Blank_Format\FY2023_Budget\"
    filePath2 = "C:\bpc_svn\BIN\ABP_Blank_Format\FY2023_Forecast\"
    fileArray1 = Array( _
    "01_Order External(Budget).xlsx", _
    "01_Order External(Budget)_JA.xlsx", _
    "01_Order Internal(Budget).xlsx", _
    "01_Order Internal(Budget)_JA.xlsx", _
    "02_Sales External(Budget).xlsx", _
    "02_Sales External(Budget)_JA.xlsx", _
    "02_Sales Internal(Budget).xlsx", _
    "02_Sales Internal(Budget)_JA.xlsx", _
    "04_Inventory(Budget).xlsx", _
    "04_Inventory(Budget)_JA.xlsx", _
    "05_Expense PL(Budget).xlsx", _
    "05_Expense PL(Budget)_JA.xlsx", _
    "11_(1st Half)Factorial analysis of MFC P+L(Budget).xlsx", _
    "11_(1st Half)Factorial analysis of MFC P+L(Budget)_By_Factory.xlsx", _
    "11_(1st Half)Factorial analysis of MFC P+L(Budget)_By_Factory_JA.xlsx", _
    "11_(1st Half)Factorial analysis of MFC P+L(Budget)_JA.xlsx", _
    "11_Factorial analysis of MFC P+L(Budget).xlsx", _
    "11_Factorial analysis of MFC P+L(Budget)_By_Factory.xlsx", _
    "11_Factorial analysis of MFC P+L(Budget)_By_Factory_JA.xlsx", _
    "11_Factorial analysis of MFC P+L(Budget)_JA.xlsx", _
    "01_Order External(Budget)_JA.xlsx")
    fileArray2 = Array( _
    "01_Order External (Forecast) _JA.xlsx", _
    "01_Order External (Forecast).xlsx", _
    "02_Sales External (Forecast) _JA.xlsx", _
    "02_Sales External (Forecast).xlsx", _
    "02_Sales Internal (Forecast) _JA.xlsx", _
    "02_Sales Internal (Forecast).xlsx", _
    "04_Inventory(Forecast) _JA.xlsx", _
    "04_Inventory(Forecast).xlsx", _
    "05_Expense (Forecast) _JA.xlsx", _
    "05_Expense (Forecast).xlsx")
    Dim i As Integer
    For i = 0 To UBound(fileArray1)
      fileName = fileArray1(i)
      DeleteSheetOptionsByFile filePath1 & fileName
    Next

    For i = 0 To UBound(fileArray2)
      fileName = fileArray2(i)
      DeleteSheetOptionsByFile filePath2 & fileName
    Next
    MsgBox "処理が完了しました。"


End Sub

Sub DeleteSheetOptionsByFile(formFile As String)

  Dim wb As Workbook
  Set wb = Workbooks.Open(fileName:=formFile)
  wb.Activate
  DoEvents
  Dim WS As Worksheet

  For Each WS In wb.Worksheets
    DeleteSheetOptionsBySheet (WS)
  Next WS

  wb.Close SaveChanges:=True
  Set WS = Nothing
  Set wb = Nothing

End Sub

Sub DeleteSheetOptionsBySheet(sheet As Worksheet)

For Each oleObj In sheet.OLEObjects
 If oleObj.Name = "FPMExcelClientSheetOptionstb1" Then
    oleObj.Delete
 End If
Next

End Sub


