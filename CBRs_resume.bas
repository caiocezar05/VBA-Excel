Attribute VB_Name = "CBRs_resume"


Sub Resumo_CBRs()
Attribute Resumo_CBRs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Separar as informações necessárias

 
Do
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:3").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Selection.Copy
    Range("D2").Select
    ActiveSheet.Paste
    Range("A2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C2").Select
    ActiveSheet.Paste
    Cells.Find(What:="assunto", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Range("E2").Select
    ActiveSheet.Paste

' Macro2 Macro
'

'
    Cells.Find(What:="Reportamo-nos", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Range(Selection, Selection.Offset(1, 0)).Select
    Selection.Copy
    Range("G2").Select
    ActiveSheet.Paste
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]&R[1]C[-1]"
    Range("H2").Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H2,G3, G2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("A3:A150").Select
    Selection.EntireRow.Delete
    
    Columns("A:A").ColumnWidth = 16
    Columns("B:B").ColumnWidth = 14
    Columns("C:C").ColumnWidth = 36
    Columns("D:D").ColumnWidth = 60
    Rows("2:2").Select
    Selection.WrapText = True
    Range("A1").Select
    
If ActiveSheet.Index <> Sheets.Count Then
ActiveSheet.Next.Select
Else
Exit Do
End If
Loop
    

'Combinar tudo em uma sheet
    Dim I As Long
    Dim xRg As Range
    On Error Resume Next
    Worksheets.Add Sheets(1)
    ActiveSheet.Name = "Resumo de CBR"
    Columns("A:A").ColumnWidth = 16
    Columns("B:B").ColumnWidth = 14
    Columns("C:C").ColumnWidth = 36
    Columns("D:D").ColumnWidth = 60

    
   For I = 2 To Sheets.Count
        Set xRg = Sheets(1).UsedRange
        If I > 2 Then
            Set xRg = Sheets(1).Cells(xRg.Rows.Count + 1, 1)
        End If
        Sheets(I).Activate
        ActiveSheet.UsedRange.Copy xRg
    Next
    
    Worksheets("Resumo de CBR").Select
    
    ' tratamentos finais
     Columns("A:A").Select
    Range("A2").Activate
    Selection.Replace What:="Brasília, ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de janeiro de ", Replacement:="/01/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de fevereiro de ", Replacement:="/02/", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de março de ", Replacement:="/03/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de abril de ", Replacement:="/04/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de maio de ", Replacement:="/05/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de junho de ", Replacement:="/06/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de julho de ", Replacement:="/07/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de agosto de ", Replacement:="/08/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de setembro de ", Replacement:="/09/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de outubro de ", Replacement:="/10/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de novembro de ", Replacement:="/11/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" de dezembro de ", Replacement:="/12/", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("C:C").Select
    Selection.Replace What:="Assunto: ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A1").Select
    For Each Worksheet In Worksheets
If ActiveSheet.Name = "Resumo de CBR" Then
ActiveSheet.Next.Select
End If
Application.DisplayAlerts = False
ActiveSheet.Delete
Next
   
    
End Sub
