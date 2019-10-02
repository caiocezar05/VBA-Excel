Attribute VB_Name = "Extrair_num_MDA"
Sub Extract_doc_texts()
Dim MSW As Word.Application
Dim arq As String
Dim Doc As Document
Dim RG As Range
Dim EndPage As Integer
Dim StartPage As Integer


StartPage = 25
EndPage = 90 'página final que será varrido

arq = Application.GetOpenFilename

If arq = "" Then
MsgBox "escolhe um arquivo ai, vei!"
Exit Sub
End If

Set MSW = New Word.Application
MSW.Visible = True 'habilite essa função caso queira acompanhar o doc
MSW.Documents.Open (arq)


Set Doc = MSW.ActiveDocument
Cells(2, 2).Select
Set oRange = Doc.Range
    
    With oRange.Find
        .Text = "R$[1-9]@[!a-z][!a-z] [mb]illion" 'mude o padrão, caso necessário.
        '.Text = "R$*[bm]illion"
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = True
        While oRange.Find.Execute = True
            oRange.Select
            myPara = Doc.Range(0, MSW.Selection.Paragraphs(1).Range.End).Paragraphs.Count
           
          
           Do Until ActiveCell = Empty
           ActiveCell.Offset(1, 0).Activate
           Loop
            
            If Doc.Paragraphs(myPara).Range <> ActiveCell.Offset(-1, 0).Value And Doc.Paragraphs(myPara).Range.Information(wdActiveEndPageNumber) > StartPage Then
            ActiveCell = Doc.Paragraphs(myPara).Range
            ActiveCell.Offset(0, -1).Value = Doc.Paragraphs(myPara).Range.Information(wdActiveEndPageNumber)
            End If
            
            If Doc.Paragraphs(myPara).Range.Information(wdActiveEndPageNumber) > EndPage Then
            MSW.Quit
            Set MSW = Nothing
            Set Doc = Nothing
            Rows("1:1").Delete
            Format_MDA
            Exit Sub
            End If
            
            oRange.Collapse wdCollapseEnd
            

        Wend
    End With
Rows("1:1").Delete

    MSW.Quit
'Release object references

Set MSW = Nothing
Set Doc = Nothing

Format_MDA

End Sub
Sub Extract_tables_doc()
Dim MSW As Word.Application
Dim arq As String
Dim Doc As Document
Dim TBL As Table
Dim StartPage As Integer
Dim EndPage As Integer
Dim exl As Workbook
Dim key As String

'a key é tipo pra selecionar um tipo de tabela, nesse caso
'eu quero todas as que tiverem ligação ou coluna como o primriro trimestre 'june 30'.
'caso vc queira pegar todas as tabelas sem dinstinção, então deixa a key com espeço: " "
key = "June 30"
Set exl = ThisWorkbook


'Mude a pagina que vc quer começar a capturar as tabelas
StartPage = 13 'selecione a pagina onde a macro começará a busca
EndPage = 150 'selecione a pagina final da busca...


arq = Application.GetOpenFilename

If arq = "" Then
MsgBox "escolhe um arquivo ai, vei!"
Exit Sub
End If

Set MSW = New Word.Application
MSW.Visible = True
MSW.Documents.Open (arq)

Set Doc = MSW.ActiveDocument
Set TBLs = Doc.Tables
u = 2

  For Each t In TBLs
  npage = t.Range.Information(wdActiveEndPageNumber)
  If t.Range.Find.Execute(key) = True And npage > StartPage And npage < EndPage Then
    t.Range.Copy
    exl.Sheets.Add After:=ActiveSheet
    Columns("A:A").ColumnWidth = 45.29
    exl.ActiveSheet.Paste
    
    On Error Resume Next
    ActiveSheet.Name = "OM page - " & npage
    
    End If
  Next
 

'Release object references

Set MSW = Nothing
Set Doc = Nothing
MSW.Quit




End Sub
Sub Format_MDA()
Attribute Format_MDA.VB_ProcData.VB_Invoke_Func = " \n14"

    Columns("B:B").ColumnWidth = 100
    Columns("B:B").WrapText = True

 

    Range("C1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[-1],SEARCH(""R$"",RC2),SEARCH(""illion"",RC2,SEARCH(""R$"",RC2))-SEARCH(""R$"",RC2)+6)"
    
    Range(Selection.Offset(0, 1), Selection.Offset(0, 9)).Select
    Selection.FormulaR1C1 = _
        "=MID(RC2,SEARCH(""R$"",RC2,SEARCH(RC[-1],RC2)+2),SEARCH(""illion"",RC2,SEARCH(""R$"",RC2,SEARCH(RC[-1],RC2)+2))-SEARCH(""R$"",RC2,SEARCH(RC[-1],RC2)+2)+6)"
   
    Range("C1:K1").Copy
    
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.SpecialCells(xlCellTypeFormulas, 16).Select
    Selection.ClearContents


' Formatar Macro

   
Range("B1").Select

For Each tx In Range(Selection, Selection.End(xlDown))
    Num = tx.Offset(0, 1)
    n = 1
    Do While Num <> ""
    
    Vposicao = InStr(tx, Num)
    Vtamanho = Len(Num)
    
'formatar a fonte
   With tx.Characters(Start:=Vposicao, Length:=Vtamanho).Font
        .Name = "Calibri"
        .FontStyle = "Negrito"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    n = n + 1
    Num = tx.Offset(0, n)
    Loop
 Next tx
'format borders
    Range("A1:M1").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Borders
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    

   'head table
    Rows("1:1").Insert Shift:=xlDown
    Rows("1:1").Insert Shift:=xlDown
    
    Range("A2").Value = "Page"
    Range("B2").Value = "Texto"
    Range("C2").Value = "VL 1"
    Range("D2").Value = "VL 2"
    Range("E2").Value = "VL 3"
    Range("F2").Value = "VL 4"
    Range("G2").Value = "VL 5"
    Range("H2").Value = "VL 6"
    Range("I2").Value = "VL 7"
    Range("J2").Value = "VL 8"
    Range("K2").Value = "VL 9"
    Range("L2").Value = "VL 10"
    Range("M2").Value = "VL 11"
    Range("N2").Value = "VL 12"
    

    Columns("A:A").Insert Shift:=xlRight
    Columns("A:A").ColumnWidth = 1
        
    Columns("B:O").AutoFit
    Range("D:O, B:B").Select
    With Selection
    
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
    Range("B2:O2").Select
    Selection.Font.Bold = True
    
     With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    'Range("D3").Select
    'ActiveWindow.FreezePanes = True habilite se vc quiser congelar o painel

        

End Sub

Sub format_python()


Range("A1").Select

For Each tx In Range(Selection, Selection.End(xlDown))
    Num = tx.Offset(0, 1)
    n = 1
    Do While Num <> ""
    
    Vposicao = InStr(tx, Num)
    Vtamanho = Len(Num)
    
'formatar a fonte
   With tx.Characters(Start:=Vposicao, Length:=Vtamanho).Font
        .Name = "Calibri"
        .FontStyle = "Negrito"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    n = n + 1
    Num = tx.Offset(0, n)
    Loop
 Next tx
 
 'format borders
    Columns("A:A").ColumnWidth = 110
    Columns("A:A").WrapText = True
    Range("A1:M1").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Borders
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Rows("1:1").Insert Shift:=xlDown
    Rows("1:1").Insert Shift:=xlDown
    Columns("A:A").Insert Shift:=xlRight
    Columns("A:A").ColumnWidth = 1
    
   'head table
    Range("B2").Value = "Texto"
    Range("C2").Value = "VL 1"
    Range("D2").Value = "VL 2"
    Range("E2").Value = "VL 3"
    Range("F2").Value = "VL 4"
    Range("G2").Value = "VL 5"
    Range("H2").Value = "VL 6"
    Range("I2").Value = "VL 7"
    Range("J2").Value = "VL 8"
    Range("K2").Value = "VL 9"
    Range("L2").Value = "VL 10"
    Range("M2").Value = "VL 11"
    Range("N2").Value = "VL 12"
    
    Columns("B:N").AutoFit
    Columns("C:N").Select
    With Selection
    
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
    Range("B2:N2").Select
    Selection.Font.Bold = True
    
     With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("C3").Select
    'ActiveWindow.FreezePanes = True habilite se vc quiser congelar o painel
    
End Sub

