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
    MsgBox "蠑ｵ譽溷￠蟋ｰ讀�蛛溷ｑ蛛溷→荳�"


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
    Worksheets(1).Activate
    Application.CutCopyMode = False
    wb.Close SaveChanges:=True
    Set WS = Nothing
    Set wb = Nothing
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub RowModify(WS As Worksheet)
  Dim rowNo As Long
  Dim dataArray As Variant
  dataArray = Array("諱門ｯｩ譌� 蟇�", "蟆句ｪｶ螂先分譌�", "謠･蟶ｵ螟帶欄", "蜒仙�槫�ｺ蜆丈ｹ募�槫ｨｭ蟄ｸ譌捺｢｡", "蟆ｭ螢吝ｽ丞ｪ晄欄", "讀�譌灘ｲ取頃譌�", "蟲蟠俶┣謠ｱ譌�", "謐灘ｺ∵､�", "謌�謳｢蝣ｷ鞫牙ｬ･蟄樊当螯�", "謌�謳｢諛晏ｹ�", "蛛ｦ蛛ｺ諛�螳ｱ譌�")
 
  For rowNo = 1 To 112
    If IsInArray(WS.Range("F" & rowNo).Value, dataArray) Then
      WS.Range("CF" & rowNo) = "X"
      WS.Range("CG" & rowNo) = ""
      WS.Range("CH" & rowNo) = ""
      WS.Range("CI" & rowNo) = ""
      WS.Range("CJ" & rowNo) = "莉ｸ蜒仙ш蜆槫Λ譏∵�ｵ螢｢譬壼が諢晄歯蛛｡蛯�"
    End If
  Next rowNo
    WS.Range("N" & 30) = "YTD"
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



