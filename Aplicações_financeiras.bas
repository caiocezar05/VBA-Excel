Attribute VB_Name = "Aplicações_financeiras"
Sub Aplicações_financeiras()
    
    Dim arq As String
    Dim w As Worksheet
    Dim Wb As Workbook
    
    
    Set Wb = ThisWorkbook

arq = Application.GetOpenFilename("Arquivo de Retorno (*xls*), *.*", Title:="Escolha o arquivo a ser importado", MultiSelect:=False)


    If arq = "" Then
        MsgBox "Você deveria ter escolhido um arquivo...", vbOKOnly, "Processo abortado"
        Exit Sub
    End If

Set w = Sheets("C.2.1")

  Cells.Find(What:= _
        "Documentação suporte: Extratos aplicações CBD", After _
        :=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows _
        , SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Offset(2#).EntireRow.Select
  
  

    Range(Selection, Selection.Offset(1000, 0)).Delete

Application.Workbooks.Open (arq)

' Separar as informações necessárias
Sheets(1).Select
 
Do

    Cells.Find(What:="Número da Operação: ", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Selection.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[-9],SEARCH(""Número da Operação: "",RC[-9]),43)"
    
    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Cut
        Selection.Offset(0, -1).Select
    Application.DisplayAlerts = False
    ActiveSheet.Paste
    
    Range("B1,B2,A3:I3,A5:I5").EntireRow.Delete
  
      Cells.Find(What:="Transação efetuada com sucesso por:", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
  Range(Selection, Selection.Offset(4, 0)).EntireRow.Delete
  

    
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
    ActiveSheet.Name = "Extratos de Aplicações"
   
   For I = 2 To Sheets.Count
        Set xRg = Sheets(1).UsedRange
        If I > 2 Then
            Set xRg = Sheets(1).Cells(xRg.Rows.Count + 1, 1)
        End If
        Sheets(I).Activate
        ActiveSheet.UsedRange.Copy xRg
    Next
    
    Worksheets("Extratos de aplicações").Select
    
    
For Each Worksheet In Worksheets
If ActiveSheet.Name = "Extratos de Aplicações" Then
ActiveSheet.Next.Select
End If
Application.DisplayAlerts = False
ActiveSheet.Delete
Next
    Columns("C:I").Select
    Selection.Style = "Comma"
    Selection.WrapText = False
    Columns("H:I").Select
    Selection.UnMerge
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
    Selection.EntireColumn.AutoFit
    Columns("a:a").Insert
    
  
        
    
    Sheets(1).UsedRange.EntireRow.Copy
     Windows("Projeto cálculo de aplicações.xlsx").Activate
     Range("a1").Select
    
    
     Cells.Find(What:= _
        "Documentação suporte: Extratos aplicações CBD", After _
        :=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows _
        , SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Offset(2, 0).EntireRow.Select

        Selection.Insert
    
    Range("a1").Select
    
MsgBox "Pronto, agora coloque o número de contrato na coluna 'A' na frente de todas as linhas de referência, caso contrário o Sumif não vai funcionar. Lembre-se tambem de fechar a planilha com os contratos de CDB. valeu!. Lembre tambem que as aplicações automáticas deverão ser digitadas manualmente"
  
    

    
    
End Sub


