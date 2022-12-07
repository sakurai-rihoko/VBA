Option Explicit

Sub ModifyTicket179402()

    Dim filePath As String
    Dim fileName As String
    Dim budgetArray As Variant
    filePath = "C:\bpc_svn\BIN\Z7_PROFITPLAN\"

    budgetArray = Array( _
    "05_Expense(Fct-PL_FY2022)_SEP.xlsm", _
    "05_Expense(Fct-PL_FY2022).xlsm")
    Dim i As Integer
    For i = 0 To UBound(budgetArray)
      
      fileName = budgetArray(i)
      ModifyProcess179402 filePath & fileName
    Next
    MsgBox "処理が完了しました。"


End Sub

Sub ModifyProcess179402(formFile As String)
  'Application.ScreenUpdating = True
  Application.Calculation = xlCalculationManual
  Dim wb As Workbook
  Set wb = Workbooks.Open(fileName:=formFile)
  wb.Activate
  DoEvents
  Dim WS As Worksheet
  Dim sheetArray As Variant

  sheetArray = Array("Expense Plan(Sheet Metal)", "Expense Plan(Micro Welding)", "Expense Plan(Cutting)", "Expense Plan(Stamping Press)", "Expense Plan(Grinding)", "Expense Plan(Other)")

  For Each WS In wb.Worksheets
        If IsInArray(WS.Name, sheetArray) Then
        WS.Activate
        DoEvents
        WS.Unprotect Password:="amadapass1"
        RowModify WS
        WS.EnableOutlining = True
        WS.Protect Password:="amadapass1", UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingColumns:=True
        End If
   Next WS
    Worksheets("Expense Plan(Total)").Activate
    Application.CutCopyMode = False
    wb.Close SaveChanges:=True
    Set WS = Nothing
    Set wb = Nothing
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub RowModify(WS As Worksheet)
  Dim rowNo As Long
  Dim dataArray As Variant
  Dim dataArray2 As Variant
  dataArray = Array("研究開発費", "展示会費", "コンピュータ関係費用", "減価償却費", "旅費交通費", "広告宣伝費", "賃借料", "貸倒引当金繰入額", "貸倒損失", "その他経費")
  dataArray2 = Array("人件費 計")
  For rowNo = 1 To 112
    If IsInArray(WS.Range("F" & rowNo).Value, dataArray) Then
      WS.Range("CF" & rowNo) = "X"
      WS.Range("CG" & rowNo) = ""
      WS.Range("CH" & rowNo) = ""
      WS.Range("CI" & rowNo) = ""
      WS.Range("CJ" & rowNo) = "※コメント必須科目を設定する"
      WS.Range("CG" & 59).Copy
      WS.Range("CF" & rowNo).PasteSpecial Paste:=xlPasteFormats
      
    End If
    If IsInArray(WS.Range("D" & rowNo).Value, dataArray2) Then
      WS.Range("CF" & rowNo) = "X"
      WS.Range("CG" & rowNo) = ""
      WS.Range("CH" & rowNo) = ""
      WS.Range("CI" & rowNo) = ""
      WS.Range("CJ" & rowNo) = "※コメント必須科目を設定する"
      WS.Range("CG" & 59).Copy
      WS.Range("CF" & rowNo).PasteSpecial Paste:=xlPasteFormats
    End If
  Next rowNo
    WS.Range("N" & 30) = "YTD"
    ActiveWindow.ScrollColumn = 1
    WS.Range("C" & 32).Activate
End Sub
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function



