Attribute VB_Name = "Conciliação_entradas"
Sub Conciliação_entradas()
Organizaçãoinicial
Conciliação_3parcelas
conciliação_2parcelas
conciliação_parcelaunica
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
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],SEARCH(""3"",RC[-3]),5)"
    

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

Sub Conciliação_3parcelas()
Range("K2").Select
Do Until IsEmpty(ActiveCell)
    If Selection.Offset(1, 0) = 0 And Selection.Offset(2, 0) = 0 And Selection.Offset(3, 0) = 0 Then
     
     
    Range(ActiveCell, ActiveCell.Offset(3, 0)).EntireRow.Copy
    ActiveSheet.Next.Select
    Range("A10000").End(xlUp).Offset(1, 0).Select
    Selection.EntireRow.Select
     Selection.Insert Shift:=xlDown
    ActiveSheet.Previous.Select
    Range(ActiveCell, ActiveCell.Offset(3, 0)).EntireRow.Delete
    Selection.Offset(1, 0).Select
    End If
    Selection.Offset(1, 0).Select
    Loop
End Sub
Sub conciliação_2parcelas()

    Range("K2").Select
  
    Do Until IsEmpty(ActiveCell)
    If Selection.Offset(1, 0) = 0 And Selection.Offset(2, 0) = 0 And Selection + Selection.Offset(1, 1) + Selection.Offset(2, 1) <> 0 Then
    
    Range(ActiveCell, ActiveCell.Offset(2, 0)).EntireRow.Copy
    ActiveSheet.Next.Select
    Range("A10000").End(xlUp).Offset(1, 0).Select
    Selection.EntireRow.Select
     Selection.Insert Shift:=xlDown
    ActiveCell.Offset(0, 10).Select
    ActiveCell = -ActiveCell.Offset(1, 1) - ActiveCell.Offset(2, 1)
      
    ActiveSheet.Previous.Select
    ActiveCell = ActiveCell + ActiveCell.Offset(1, 1) + ActiveCell.Offset(2, 1)
    Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(2, 0)).EntireRow.Delete
    End If
    
    If Selection.Offset(1, 0) = 0 And Selection.Offset(2, 0) = 0 And -Selection.Offset(1, 1) - Selection.Offset(2, 1) = Selection Then
    Range(ActiveCell, ActiveCell.Offset(2, 0)).EntireRow.Copy
    ActiveSheet.Next.Select
    Range("A10000").End(xlUp).Offset(1, 0).Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown
    ActiveSheet.Previous.Select
    Range(ActiveCell, ActiveCell.Offset(2, 0)).EntireRow.Delete
    Selection.Offset(-1, 0).Select
    End If
    
    Selection.Offset(1, 0).Select
    Loop
End Sub
Sub conciliação_parcelaunica()

      Range("M2").Select

     
    If Selection = Selection.Offset(1, 0) And Selection.Offset(0, -2) + Selection.Offset(1, -1) <> 0 Then
    Range(ActiveCell, ActiveCell.Offset(1, 0)).EntireRow.Copy
    ActiveSheet.Next.Select
    Range("A10000").End(xlUp).Offset(1, 0).Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(0, 10).Select
    ActiveCell = -ActiveCell.Offset(1, 1)
    
    ActiveSheet.Previous.Select
    ActiveCell.Offset(0, -2) = ActiveCell.Offset(0, -2) + ActiveCell.Offset(1, -1)
    ActiveCell.Offset(1, 0).EntireRow.Delete
    End If
    
        Do Until IsEmpty(ActiveCell)
    If Selection = Selection.Offset(1, 0) And Selection.Offset(0, -2) + Selection.Offset(1, -1) = 0 Then
     
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
Sub acabamento()
      ' Acabamento
      
    Rows("1:1").Select
    Selection.AutoFilter
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    ActiveSheet.Next.Select
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("D:J").Select
    Selection.EntireColumn.Hidden = True
    ActiveSheet.Previous.Select
    Columns("D:J").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Sub
