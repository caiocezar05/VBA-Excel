Attribute VB_Name = "Calculo_AV_HV_monetV"
Sub Calc_AV_HV()
Attribute Calc_AV_HV.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AV%
'
'
    Cells.Find(What:="100%", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
        If Selection.Offset(0, -1) = "" Then
        Selection.End(xlToLeft).Offset(0, 2).Select
        End If
        
        If Selection.Offset(0, -1) <> "" Then
        Selection = ""
        Selection.Offset(0, 1).Select
        End If
        

       R1 = ActiveCell.Row
       C1 = ActiveCell.Column
       
       
    ActiveCell.FormulaR1C1 = "=RC[-4]/RC[-4]"
        ActiveCell.Replace What:="/D", Replacement:="/D$", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Selection.Copy
    Range(ActiveCell, ActiveCell.Offset(150, 2)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00%"


' AH%
'

'
    ActiveCell.Offset(0, 4).Select
    ActiveCell.FormulaR1C1 = "=(RC[-8]/RC[-7])-1"
    Selection.Copy
    Range(ActiveCell, ActiveCell.Offset(150, 1)).Select
    ActiveSheet.Paste
    Selection.NumberFormat = "0.00%"
 ' A$%
    ActiveCell.Offset(0, 3).Select
    ActiveCell.FormulaR1C1 = "=RC[-10]-RC[-11]"
    Application.CutCopyMode = False
    Selection.Copy
    
    Range(ActiveCell, ActiveCell.Offset(150, 1)).Select
    ActiveSheet.Paste
    Selection.NumberFormat = "0.0"
    
    Range("C5:C160").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete

    
'
' formt Macro
'

'


    
    
' Formatação de cabeçalho AV%
    Cells(R1, C1).Offset(-1, 0).Select
    Range(Selection, Selection.Offset(0, 2)).Select
    Selection.FormulaR1C1 = "=RC[-4]"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    
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
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

Selection.Offset(-1, 0).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "AV%"
    Selection.Columns.EntireColumn.AutoFit


' Formatação de cabeçalho AH%
    Selection.Offset(1, 2).Select
    Range(Selection, Selection.Offset(0, 1)).Select
    Selection.FormulaR1C1 = "=RC[-3] &"" to ""& RC[-4]"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
 With Selection
        .WrapText = True
        .ColumnWidth = 13
        .EntireRow.AutoFit
End With
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
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

Selection.Offset(-1, 0).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "AH%"
    Selection.Columns.EntireColumn.AutoFit
   
   
   
  'Formatar ultima
   Range(Selection, Selection.Offset(1, 0)).Select
   Selection.Copy
   Selection.Offset(0, 3).Select
   Selection.ColumnWidth = 13
   ActiveSheet.Paste
   ActiveCell = "Variação R$"
   Application.CutCopyMode = False
   Range("B4").Select
   
   
End Sub
