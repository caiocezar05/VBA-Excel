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
        
       
    ActiveCell.FormulaR1C1 = "=RC[-4]/RC[-4]"
        ActiveCell.Replace What:="/C", Replacement:="/C$", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Selection.Copy
    Selection.SpecialCells(xlCellTypeConstants, 1).Offset(0, 4).Select
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
    Selection.NumberFormat = "#,##0.0"
    
    Range("B5:    B160 ").Select ""
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete

    End Sub
'
' formt Macro
'

'

Sub formact()
    
    
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
   
End Sub
