Attribute VB_Name = "Inserir_DF_eng_txt"
Sub Inserir_dados_txt()
Attribute Inserir_dados_txt.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Inserir Macro
'

'
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\caio.santos\Desktop\text.txt", Destination:=Range("$A$1"))
        
        .Name = "text"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

' tirar_numeros Macro
'

'
    Columns("A:Z").Select
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

        
    Selection.SpecialCells(xlCellTypeConstants, 21).Select
    Selection.ClearContents
    Range("A1").Select

' concatenar Macro
'

'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[1],"" "",RC[2],"" "",RC[3],"" "",RC[4],"" "",RC[5],"" "",RC[6],"" "",RC[7],"" "",RC[8],"" "",RC[9],"" "",RC[10],"" "",RC[11],"" "",RC[12])"
    Selection.AutoFill Destination:=Range("A1:A275"), Type:=xlFillDefault
    Range("A1:A275").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select


' limpar_excesso Macro
'

'
    Range("A1:J317").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A280").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A261").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A242").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A223").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A204").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A185").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A166").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A147").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A128").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A109").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A90").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A71").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A52").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A33").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A14").Select
    ActiveWindow.LargeScroll Down:=-1
    Columns("B:R").Select
    Selection.QueryTable.Delete
    Selection.ClearContents
    Range("A1").Select
    Columns("A:A").EntireColumn.AutoFit

' inserir2 Macro
'

'
    Range("B1").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\caio.santos\Desktop\text.txt", Destination:=Range("$B$1"))
        
        .Name = "text_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

' tirar_texto Macro
'

'
    Columns("B:Q").Select
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="-", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.SpecialCells(xlCellTypeConstants, 22).Select
    Selection.ClearContents
    
' desloc_left Macro
    Columns("B:O").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll ToRight:=-1
    Range("A1").Select
End Sub

Sub Tratar_num_DF_txt()
'
' Convert_num Macro
'

    Range("A1").Select
    Selection.Copy
    Columns("B:G").Select
    Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Style = "Comma"

' arredondar Macro
    
    Range("A1").Select
    ActiveCell = 1
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "=ROUND(RC[-6],R1C1)"
    Selection.AutoFill Destination:=Range("H1:M1"), Type:=xlFillDefault
    Range("H1:M1").Select
    Selection.AutoFill Destination:=Range("H1:M150"), Type:=xlFillDefault
    Range("H1:M150").Select
    Selection.Copy

    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("H:M").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=-1
    Range("A1").Select
    ActiveCell = ""
    Columns("B:B").Select
    Range("B22").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    
End Sub


Sub FormataçãoDF()
'
' Ajustes Macro
'

'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.ColumnWidth = 1


'
' ajuste2 Macro
'


    Range("C2: E2 ").Select
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    
    Selection.Cut
    Selection.End(xlDown).Offset(-2, 0).Select
    ActiveSheet.Paste
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .ReadingOrder = xlContext
    End With
    
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
        Range("C1:E1").Select
        Selection.SpecialCells(xlCellTypeConstants, 23).Select
    
    Selection.Cut
    Selection.End(xlDown).Offset(-1, 0).Select
    ActiveSheet.Paste
   
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .ReadingOrder = xlContext
    End With

        With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Selection.Offset(-1, 0).Select
    ActiveCell.FormulaR1C1 = "For the year ended as December, 31"
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)

        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
     With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .ReadingOrder = xlContext
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub
   

Sub Calc_AV_HV()
'
' AV%
    Cells.Find(What:="100%", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
        
     Do
     Selection.Offset(0, -1).Select
     Loop Until Not IsNumeric(Selection) And Selection <> ""
    
    Do
    Selection.Offset(0, 2).Select
    Selection.EntireColumn.Insert
    Rw = Selection.Offset(0, -1)



    Selection.Offset(0, -1).EntireColumn.Select
    Selection.SpecialCells(xlCellTypeConstants, 1).Offset(0, 1).Select
    
    For Each cell In Selection
    If ActiveCell.Offset(0, -1) <> "" Then
    ActiveCell = ActiveCell.Offset(0, -1) / Rw
    Selection.NumberFormat = "0.00%"
    ActiveCell.Offset(1, 0).Select
    End If
   Next
   Selection.Offset(-1, 0).Select
   
  Loop While Selection.Offset(0, 1) <> ""


' AH%
'

    Selection.EntireColumn.Select
    Selection.SpecialCells(xlCellTypeConstants, 1).Offset(0, 2).Select
    
    
    Selection.FormulaR1C1 = "=(RC[-5]/RC[-3])-1"
    Selection.NumberFormat = "0.00%"

 ' A$%
    Selection.EntireColumn.Select
    Selection.SpecialCells(xlCellTypeFormulas, 1).Offset(0, 2).Select
    Selection.FormulaR1C1 = "=RC[-5]-RC[-7]"
    Selection.NumberFormat = "#,##0.0"
    
    End Sub
    Sub formatares()
' Formatação de cabeçalho AV%
    Cells.Find(What:="For the year ended", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate



    Range(Selection, Selection.Offset(2, 0)).Select
    Selection.Copy
    Selection.Offset(0, 4).Select
    ActiveSheet.Paste
    ActiveCell = "AV%"
    ActiveCell.Offset(2, 0).Select
    Range(Selection, Selection.Offset(0, 2)).Select
    Selection = "Cálculo"
    


' Formatação de cabeçalho AH%
    Range(Selection, Selection.Offset(-2, 0)).Select
    Selection.Copy
    Range(ActiveCell.Offset(0, 2), ActiveCell.Offset(2, 3)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveCell = "AH%"

   ActiveCell.Offset(1, 0).Select
   Range(Selection, Selection.Offset(0, 1)).Select
   
    Selection.FormulaR1C1 = "=RC[-3] &"" to ""& RC[-4]"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False

        Selection.EntireColumn.AutoFit
        Selection.Offset(1, 0).Select
        Selection = "Cálculo"
       
       
 'format ultima
       
    Range(Selection, Selection.Offset(-2, 0)).Select
    Selection.Copy
    Selection.Offset(0, 3).Select
    ActiveSheet.Paste
    ActiveCell = "Variação R$"
    
    Application.CutCopyMode = False
    Selection.EntireColumn.AutoFit
    Range("A1").Select
    
    Selection.SpecialCells(xlCellTypeFormulas, 16).Select
    Selection = "N.M."
    Selection.HorizontalAlignment = xlCenter
    Range("F:F, J:J, M:M").Select
    Selection.ColumnWidth = 1
    
End Sub




