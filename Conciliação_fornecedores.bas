Attribute VB_Name = "Conciliação_fornecedores"
Sub Conciliação_fornecedores()
Organizaçãoinicial
Conciliação_2parcelas
Conciliação_parcelaunica
Acabamento
End Sub

Sub Organizaçãoinicial()
Columns("D:J").Select
    Selection.EntireColumn.Hidden = False
    ActiveSheet.Next.Select
    Columns("D:J").Select
    Selection.EntireColumn.Hidden = False
    ActiveSheet.Previous.Select
 Range("M2").Select
 
 'Formula de midle
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]<>0,MID(RC[-3],10,7),MID(RC[-3],16,7))"
    

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
