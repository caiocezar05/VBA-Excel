Attribute VB_Name = "Automatic_A4"
Sub Automatic_A4()
Attribute Automatic_A4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Variaveis
    Dim xRCount As Long
    Dim xSht As Worksheet
    Dim xNSht As Worksheet
    Dim I As Long
    Dim xTRrow As Integer
    Dim xCol As New Collection
    Dim xTitle As String
    Dim xSUpdate As Boolean

'

'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("B:B").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Rows("1:1").Select
    Range("B1").Activate
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "REF."
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Conta"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Nome da conta"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Saldo inicial"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Débito"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Crédito"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Movimentação"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Saldo final"
    Range("A1").Select

' Fit colunas
'

'
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[1],11)"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A706")
    Range("A2:A706").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'
' Colocar as letrinhas
'

'
    Selection.Replace What:="1.01.01.001", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.01.002", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.01.003", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.01.004", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.01.005", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.01.006", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Selection.Replace What:="1.01.01.007", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    

    Selection.Replace What:="1.01.02.001", Replacement:="E1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   
    Selection.Replace What:="1.01.02.001", Replacement:="E1", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.02.002", Replacement:="E2", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Selection.Replace What:="1.01.03.001", Replacement:="E2", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.03.001", Replacement:="E1", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Selection.Replace What:="1.01.03.003", Replacement:="E4", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
            Selection.Replace What:="1.01.03.004", Replacement:="E4", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   
    Selection.Replace What:="1.01.03.005", Replacement:="O1", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   
    Selection.Replace What:="1.01.06.001", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.06.003", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.08.001", Replacement:="F", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1.01.08.001", Replacement:="G", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   Selection.Replace What:="1.07.03.003", Replacement:="E5", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="1.07.05.002", Replacement:="H", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        'continua
            Selection.Replace What:="1.07.06.002", Replacement:="K1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                    Selection.Replace What:="1.07.06.003", Replacement:="K1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                    Selection.Replace What:="1.07.06.004", Replacement:="K1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                           Selection.Replace What:="1.07.07.001", Replacement:="L1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        'passivo
        Selection.Replace What:="2.01.01.001", Replacement:="N1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="2.01.01.002", Replacement:="N1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                Selection.Replace What:="2.01.02.001", Replacement:="M1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                        Selection.Replace What:="2.01.02.002", Replacement:="M1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
         Selection.Replace What:="2.01.03.001", Replacement:="P", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
         Selection.Replace What:="2.01.06.004", Replacement:="P", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="2.01.03.002", Replacement:="P", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="2.01.04.001", Replacement:="O2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                Selection.Replace What:="2.01.08.002", Replacement:="O2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                        Selection.Replace What:="2.01.08.001", Replacement:="O2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="2.01.06.001", Replacement:="N2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="2.01.06.002", Replacement:="N2.2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                Selection.Replace What:="2.01.06.003", Replacement:="N2.2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="2.02.02.001", Replacement:="M2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                Selection.Replace What:="2.02.02.002", Replacement:="M2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                Selection.Replace What:="2.02.06.003", Replacement:="N2.1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
                        Selection.Replace What:="2.02.09.001", Replacement:="S1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                                Selection.Replace What:="1.07.03.001", Replacement:="J", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
                                       Selection.Replace What:="1.07.03.002", Replacement:="J", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        Range("A:A").Select
        Selection.Font.Bold = True
        Selection.Font.Color = -16776961

            Rows("1:1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A:A").Select
    Selection.Find(What:="C", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
        
        Range("A2", Selection.Offset(-1, 0)).EntireRow.Delete
        Range("A:A").Select
        
        Selection.Find(What:="Total", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
        
         Range(Selection, Selection.End(xlDown)).EntireRow.Delete
         Range("1:1").Select
         Range("E:G").EntireColumn.Delete
         Range("A1").Select
         ActiveSheet.Name = "Balancete"
         
             Set xSht = ActiveSheet
    On Error Resume Next
    xRCount = xSht.Cells(xSht.Rows.Count, 1).End(xlUp).Row
    xTitle = "A1:C1"
    xTRrow = xSht.Range(xTitle).Cells(1).Row
    For I = 2 To xRCount
        Call xCol.Add(xSht.Cells(I, 1).Text, xSht.Cells(I, 1).Text)
    Next
    xSUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    For I = 1 To xCol.Count
        Call xSht.Range(xTitle).AutoFilter(1, CStr(xCol.Item(I)))
        Set xNSht = Nothing
        Set xNSht = Worksheets(CStr(xCol.Item(I)))
        If xNSht Is Nothing Then
            Set xNSht = Worksheets.Add(, Sheets(Sheets.Count))
            xNSht.Name = CStr(xCol.Item(I))
        Else
            xNSht.Move , Sheets(Sheets.Count)
        End If
        xSht.Range("A" & xTRrow & ":A" & xRCount).EntireRow.Copy xNSht.Range("A1")
        xNSht.Columns.AutoFit
    Next
    xSht.AutoFilterMode = False
    xSht.Activate
    Application.ScreenUpdating = xSUpdate

         
End Sub
