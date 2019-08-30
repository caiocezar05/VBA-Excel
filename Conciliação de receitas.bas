Attribute VB_Name = "Módulo1"
Sub Conciliação_fornecedores()
Importar_Razão

Conciliação_2parcelas
Conciliação_parcelaunica
Acabamento
End Sub
Sub Verificar()
 Range("L2").Select

      
    Do Until IsEmpty(ActiveCell)
    If Selection.Offset(1, 0) = 0 And Selection.Offset(2, 0) = 0 And Selection.Offset(3, 0) = 0 And Selection.Offset(4, 0) = 0 And Selection.Offset(5, 0) = 0 Then

   MsgBox ("Há Valores que essa humild macro não consegue conciliar. Dê uma olhadinha :D")
   Exit Sub
   End If
   Selection.Offset(1, 0).Select
   Loop
   
End Sub

Sub Importar_Razão()
'IMPORTAR RAZÂO

Dim arq As String
    Dim w As Worksheet


arq = Application.GetOpenFilename("Arquivo de Retorno (*xls*), *.*", Title:="Escolha o arquivo a ser importado", MultiSelect:=False)


        If arq = "" Then
        MsgBox "Você deveria ter escolhido um arquivo...", vbOKOnly, "Processo abortado"
        Exit Sub
    End If

'Mudar o nome para a sheet de conciliação desejada
Set w = Sheets("211011102")

Application.Workbooks.Open (arq)
Columns("B:D").Delete Shift:=xlToLeft
    Columns("C:C").Insert Shift:=xlToRight
    
    'formula
    Range("L:L").Delete
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]>0,RC[-1],0)"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]<0,RC[-2],0)"
    
    Range("L2:M2").Copy
    
    Range("K2", Range("k2").End(xlDown)).Offset(0, 1).Select
    
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("K:K").Delete
    Range("A2:L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    w.Cells(1, 1).End(xlDown).Offset(1, 0).EntireRow.Select
    Selection.Insert Shift:=xlDown
    
    
    'Fechar esse treco
    Application.DisplayAlerts = False
    
        ActiveWorkbook.Close SaveChanges:=False
    
    Application.DisplayAlerts = True

Columns("D:J").Select
    Selection.EntireColumn.Hidden = False
    ActiveSheet.Next.Select
    Columns("D:J").Select
    Selection.EntireColumn.Hidden = False
    ActiveSheet.Previous.Select
Range("M2").Select

 'Formula de midle
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],FIND(""OP"",RC[-3])+3,7)"
    

    Selection.AutoFill Destination:=Range(ActiveCell, Range("J10000").End(xlUp).Offset(0, 3))
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:= _
        Range("M1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Range("A1").Select

    End Sub
    
    Sub Conciliação_2parcelas()
    Range("L2").Select

      
    Do Until IsEmpty(ActiveCell)
    If Selection.Offset(1, 0) = 0 And Selection.Offset(2, 0) = 0 Then
     
    Range(ActiveCell, ActiveCell.Offset(2, 0)).EntireRow.Copy
    ActiveSheet.Next.Select
    Range("A10000").End(xlUp).Offset(1, 0).Select
    Selection.EntireRow.Select
     Selection.Insert Shift:=xlDown
    ActiveCell.Offset(0, 11).Select
    ActiveCell = -ActiveCell.Offset(1, -1) - ActiveCell.Offset(2, -1)
      
    ActiveSheet.Previous.Select
    ActiveCell = ActiveCell + ActiveCell.Offset(1, -1) + ActiveCell.Offset(2, -1)
    Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(2, 0)).EntireRow.Delete
    End If
    Selection.Offset(1, 0).Select
    Loop
End Sub

Sub Conciliação_parcelaunica()
      Range("M2").Select
    Do Until IsEmpty(ActiveCell)
    If Selection = Selection.Offset(1, 0) Then
     
    Range(ActiveCell, ActiveCell.Offset(1, 0)).EntireRow.Cut
    ActiveSheet.Next.Select
    Range("A10000").End(xlUp).Offset(1, 0).Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown
    ActiveSheet.Previous.Select
    Range(ActiveCell, ActiveCell.Offset(1, 0)).EntireRow.Delete
    Selection.Offset(-1, 0).Select
     End If
    Selection.Offset(1, 0).Select
    Loop
      
   End Sub
   Sub Acabamento()
      ' Acabamento
      
      Rows("1:1").Select
    Selection.AutoFilter
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    ActiveSheet.Next.Select
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    Columns("E:I").Select
    Selection.EntireColumn.Hidden = True
    ActiveSheet.Previous.Select
    Columns("E:I").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Sub


