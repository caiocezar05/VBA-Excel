Attribute VB_Name = "Conciliação_bancária"
Sub Import_tesouraria()
    Dim arq As Variant
    Dim Vcel As String
    Dim L As String
   Dim vln As String
    
Application.ScreenUpdating = False
    Sheets("Tesouraria").Select
    'Configura a leitura do arquivo
arq = Application.GetOpenFilename("Arquivo de Retorno (*txt*), *.*", Title:="Escolha o arquivo a ser importado", MultiSelect:=False)
    
    
    If arq = "" Or arq = False Then
    MsgBox "Você deveria ter escolhido algum arquivo..", vbOKOnly
    Exit Sub
    End If
    
    i = FreeFile
    
    'Abre o arquivo para leitura
    Open arq For Input As #i
    L = 1
    vln = 2

    'Lê o conteúdo do arquivo linha a linha
    Do While Not EOF(i)
        Line Input #i, L
       
        Cells(vln, 1) = L

        vln = vln + 1
    Loop
 

 ' Formatação
    
    
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(5, 1), Array(16, 4), Array(26, 9), Array(36, 1), _
        Array(80, 1), Array(96, 1), Array(98, 9)), TrailingMinusNumbers:=True
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeConstants, 22).Select
    Selection.EntireRow.Delete

    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "tipo"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "data"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Descrição"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "valor"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "t"
    Range("D1").Select
    Sheets("Capa").Select
    'Fecha o arquivo
    Close (i)
     MsgBox "pronto"
End Sub

Sub import_contábil()

    Dim arq As String
    Dim w As Worksheet


arq = Application.GetOpenFilename("Arquivo de Retorno (*xls*), *.*", Title:="Escolha o arquivo a ser importado", MultiSelect:=False)


        If arq = "" Then
        MsgBox "Você deveria ter escolhido um arquivo...", vbOKOnly, "Processo abortado"
        Exit Sub
    End If


Set w = Sheets("contábil")

Application.Workbooks.Open (arq)
    ActiveSheet.Range("A1").CurrentRegion.Select
    Selection.Copy Destination:=w.Cells(1, 1)
    
    Application.DisplayAlerts = False
    
        ActiveWorkbook.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    
     Sheets("Capa").Select

End Sub
Sub import_extrato()
    Dim arq As Variant
    Dim Vcel As String
    Dim L As String

    
    
Application.ScreenUpdating = False
   Set w = Sheets("Extrato")
   w.Select
    'Configura a leitura do arquivo
arq = Application.GetOpenFilename("Arquivo de Retorno (*txt*), *.*", Title:="Escolha o arquivo a ser importado", MultiSelect:=False)
    
    If arq = "" Or arq = False Then
    MsgBox "Você deveria ter escolhido algum arquivo..", vbOKOnly
    Exit Sub
    End If
    
    i = FreeFile
    
    'Abre o arquivo para leitura
    Open arq For Input As #i
    
    L = 4
    vln = 2

    'Lê o conteúdo do arquivo linha a linha

    Do While Not EOF(i)
       
       Line Input #i, L


        
        If Mid(L, 1, 1) = "=" Then
        MsgBox "Tente tirar a primeira linha do arquivo txt do extrato que você está tentando importar, as vezes o excel não reconhece o '===='"
     Exit Sub
     End If
     
        Cells(vln, 1) = L

        vln = vln + 1


    Loop

Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 4), Array(13, 9), Array(57, 1), Array(89, 9), Array(114, 1), _
        Array(129, 1), Array(131, 9)), TrailingMinusNumbers:=True
    
    
    Selection.SpecialCells(xlCellTypeConstants, 22).Select
    Selection.EntireRow.Delete
    Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Data"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Descrição"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Valor"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "t"
    Range("E1").Select
    Sheets("Capa").Select
    'Fecha o arquivo
    Close (i)
     MsgBox "pronto"

End Sub

Sub Formatação_chave()
Attribute Formatação_chave.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Chave do extrato
    Sheets("Extrato").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Chave"
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A2").FormulaR1C1 = "=RC[1]&RC[3]&RC[4]"
    Range("A2").Copy
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, -1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'
' Coluna extrato e contábil da tesouraria
'
    Sheets("Tesouraria").Select
    Range("F1") = "Contábil"
    Range("G1") = "Extrato"
 
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]=""Transação"",IF(RC[-1]=""D"",-RC[-2],RC[-2]),0)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""Extrato"",IF(RC[-2]=""D"",-RC[-3],RC[-3]),0)"
    
    Range("F2:G2").Copy
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("F:G").EntireColumn.AutoFit

' Chaves tesouraria
'

'
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("C1") = "Chave Ex"
    Range("D1") = "Chave Cn"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]&RC[3]&RC[4]"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[4]"
    Range("C2:D2").Copy
    
    Range("E3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, -2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'
' Format contábil
'

'
    Sheets("Contábil").Select
    Range("A:E,G:K,N:N").Select
    Selection.Delete Shift:=xlToLeft
    
' Chave contábil
'

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1") = "Chave"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]&RC[2]"
    Range("B2").Copy
    Range("C3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, -1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("D:D").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("A:A").EntireColumn.AutoFit
    Sheets("Capa").Select
    
    MsgBox "Pronto, verifique se está tudo bem bem nas planilhas, se estão todos com chave. caso esteja, prossiga para a função >conciliar<"
End Sub
Sub Filtros()
Attribute Filtros.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filtro extrato
'

'
    Sheets("Extrato").Select
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],Tesouraria!C3:C6,4,FALSE)"
    Selection.Copy
    Range("E3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Range("A1:F1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$F$10000").AutoFilter Field:=6, Criteria1:="#N/D"
    Range("B:E").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    Sheets("conciliação").Select
    Range("B:E").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'
' filtro tesouraria
'

    Sheets("Tesouraria").Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],Contábil!C2:C4,3,FALSE)"
    Selection.Copy
    Range("I3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1:J1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$10000").AutoFilter Field:=10, Criteria1:="#N/D"
    ActiveSheet.Range("$A$1:$J$10000").AutoFilter Field:=1, Criteria1:="transação"
    
    Range("B1:B10000").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    Sheets("conciliação").Select
    
    Range("B2").Select
    Selection.End(xlDown).Offset(3, 0).Select
    ActiveSheet.Paste
    
    Sheets("Tesouraria").Select
    Range("E1:E10000").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    
    Sheets("conciliação").Select
    Range("C2").Select
    Selection.End(xlDown).Offset(3, 0).Select
    ActiveSheet.Paste
    
    Sheets("Tesouraria").Select
    Range("H1:H10000").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    
    Sheets("conciliação").Select
    Range("D2").Select
    Selection.End(xlDown).Offset(3, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


' Filtro contábil
'

'
    Sheets("Contábil").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],Tesouraria!C4:C8,5,FALSE)"
    Selection.Copy
    Range("D3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1:J1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$10000").AutoFilter Field:=5, Criteria1:="#N/D"
    
    Range("A1:A10000").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    Sheets("conciliação").Select
    
    Range("B10000").Select
    Selection.End(xlUp).Offset(3, 0).Select
    ActiveSheet.Paste
    
     Sheets("Contábil").Select
    Range("C1:D10000").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    Sheets("conciliação").Select
    
    Range("C10000").Select
    Selection.End(xlUp).Offset(3, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    

    Columns("B:E").EntireColumn.AutoFit
    Range("A1").Select
     Sheets("Capa").Select
End Sub

Sub Clear()
'

    Sheets(Array("Extrato", "Tesouraria", "Contábil", "conciliação")).Delete
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Extrato"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Tesouraria"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Contábil"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Conciliação"
    Sheets("Capa").Select
    End Sub

