Attribute VB_Name = "Módulo1"

Sub import_BB()
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

Sub import_bradesco()

Worksheets("Extrato").Activate
Range("A5:A5").Select

        If Selection = "" Then
        Sheets("Capa").Select
        MsgBox "Poxa vida, vamos fazer uma forcinha e importar manualmenet esse extrato? Abra ele, copie e cole na sheet extrato e depois clica em mim...", vbOKOnly, "Processo abortado"
        Exit Sub
    End If


    Range("C:C,F:F").Select
    Selection.Delete Shift:=xlToLeft

    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeConstants, 2).Select
    Selection.EntireRow.Delete
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "DATA"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Histórico"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Valor"
    Range("D1").Select

    Columns("C:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
    
     Sheets("Capa").Select

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
Sub Formatação_chave()




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
    Range("A2").FormulaR1C1 = "=RC[1]&RC[3]"
    Range("A2").Copy
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, -1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

' Format contábil

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
'
' Filtro extrato
'

'
    Sheets("Extrato").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],contábil!C2:C4,3,FALSE)"
    Selection.Copy
    Range("D3").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Range("A1:F1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$10000").AutoFilter Field:=5, Criteria1:="#N/D"
    Range("B:E").Select
    Selection.SpecialCells(xlCellTypeVisible).Copy
    Sheets("conciliação").Select
    Range("B:E").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


' Filtro contábil
'

'
    Sheets("Contábil").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],Extrato!C1:C4,4,FALSE)"
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
  
End Sub

Sub Clear()
'

    Sheets(Array("Extrato", "Contábil", "conciliação")).Delete
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Extrato"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Contábil"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Conciliação"
    Sheets("Capa").Select
    End Sub


