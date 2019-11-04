Attribute VB_Name = "Conciliação_bancária"
Public Sub LeArquivoTexto()
    Dim Arq As Variant
    Dim Vcel As String
    Dim L As String

    
Application.ScreenUpdating = False
    Sheets("Tesouraria").Select
    'Configura a leitura do arquivo
Arq = Application.GetOpenFilename("Arquivo de Retorno (*txt*), *.*", Title:="Escolha o arquivo a ser importado", MultiSelect:=False)
    
    If Arq = "" Or Arq = False Then
    MsgBox "Você deveria ter escolhido algum arquivo..", vbOKOnly
    Exit Sub
    End If
    
    i = FreeFile
    
    'Abre o arquivo para leitura
    Open Arq For Input As #i
    L = 1
    Vln = 2

    'Lê o conteúdo do arquivo linha a linha
    Do While Not EOF(i)
        Line Input #i, L
       
        tipo = Mid(L, 6, 11)
        Data = Mid(L, 17, 10)
        desc = Mid(L, 37, 44)
        Vl = Mid(L, 81, 16)
        tipo2 = Mid(L, 97, 2)
        
        Cells(Vln, 1) = tipo
        Cells(Vln, 2) = Data
        Cells(Vln, 3) = desc
        Cells(Vln, 4) = Vl
        Cells(Vln, 5) = tipo2
        Vln = Vln + 1
    Loop
 

 ' Formatação
'

'
    Columns("D:D").Select
        Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.SpecialCells(xlCellTypeConstants, 22).Select
    Selection.EntireRow.Delete
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
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
        "=IF(RC[-5]=""transação"",IF(RC[-1]=""D"",-RC[-2],RC[-2]),0)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""extrato"",IF(RC[-2]=""D"",-RC[-3],RC[-3]),0)"
    
    Range("F2:G2").Copy
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
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
    Range("A1").Select
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
    ActiveSheet.Range("$A$1:$F$10000").AutoFilter Field:=6, Criteria1:="#N/A"
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
    ActiveSheet.Range("$A$1:$J$10000").AutoFilter Field:=10, Criteria1:="#N/A"
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
    ActiveSheet.Range("$A$1:$E$10000").AutoFilter Field:=5, Criteria1:="#N/A"
    
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
    
End Sub

Sub Salvar_resultado()
'

    Sheets(Array("Extrato", "Tesouraria", "Contábil", "conciliação")).Move
    ChDir "C:\Users\caio.santos\Desktop\projeto CESP\Conciliação de março"
    Windows("Conciliação.xlsx").Activate
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Extrato"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Tesouraria"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Contábil"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Conciliação"
    Sheets("Capa").Select
    ActiveWorkbook.Save
    End Sub

