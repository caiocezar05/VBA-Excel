Attribute VB_Name = "Extrair_num_MDA"
Sub Extract_CircleUptexts()
Dim MSW As Word.Application
Dim arq As String
Dim Doc As Document
Dim RG As Range
Dim EndPage As Integer
Dim StartPage As Integer


StartPage = 10
EndPage = 100 'página final que será varrido

arq = Application.GetOpenFilename

If arq = "" Then
MsgBox "escolhe um arquivo!"
Exit Sub
End If

Set MSW = New Word.Application
MSW.Visible = True
MSW.Documents.Open (arq)


Set Doc = MSW.ActiveDocument
Cells(2, 2).Select
Set oRange = Doc.Range
    
    With oRange.Find
        .Text = "R$*million" 'mude o padrão, caso necessário.
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
            ActiveCell.Offset(0, 5).Value = Doc.Paragraphs(myPara).Range.Information(wdActiveEndPageNumber)
            End If
            
            If Doc.Paragraphs(myPara).Range.Information(wdActiveEndPageNumber) > EndPage Then
            MSW.Quit
            Set MSW = Nothing
            Set Doc = Nothing
            Columns("A:A").Delete
            Rows("1:1").Delete
            Format_MDA
            Exit Sub
            End If
            
            oRange.Collapse wdCollapseEnd
            

        Wend
    End With
Columns("A:A").Delete
Rows("1:1").Delete


    MSW.Quit
'Release object references

Set MSW = Nothing
Set Doc = Nothing

Format_MDA

End Sub
Sub newgg()
Dim MSW As Word.Application
Dim arq As String
Dim Doc As Document
Dim RG As Range
Dim EndPage As Integer
Dim StartPage As Integer


StartPage = 10
EndPage = 100 '

arq = Application.GetOpenFilename

If arq = "" Then
MsgBox "escolhe um arquivo ai, vei!"
Exit Sub
End If

Set MSW = New Word.Application
MSW.Documents.Open (arq)


Set Doc = MSW.ActiveDocument
Set oRange = Doc.Range

    Columns("A:A").Insert
    Cells(2, 2).Select

Do While Selection <> Empty

findtext = Left(ActiveCell, 200)
    With oRange.Find
        .Text = findtext
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = True
        On Error Resume Next
        oRange.Find.Execute
            oRange.Select
            myPara = Doc.Range(0, MSW.Selection.Paragraphs(1).Range.End).Paragraphs.Count

            ActiveCell.Offset(0, -1).Value = Doc.Paragraphs(myPara).Range.Information(wdActiveEndPageNumber)
       
       
            oRange.Collapse wdCollapseEnd

    End With
    Selection.Offset(1, 0).Select
Loop
MSW.Quit
'Release object references

Set MSW = Nothing
Set Doc = Nothing

End Sub


Sub doc_tables()
Dim MSW As Word.Application
Dim arq As String
Dim Doc As Document
Dim TBL As Table
Dim StartPage As Integer
Dim EndPage As Integer
Dim exl As Workbook
Dim key As String

'___a key é tipo pra selecionar um tipo de tabela, nesse caso eu quero todas as que tiverem ligação ou coluna como o primriro trimestre "março
key = " "
Set exl = ThisWorkbook

'Mude a pagina que vc quer começar a capturar as tabelas
StartPage = 20 'da pra mudar depois
EndPage = 50 'depois da pra trocar...


arq = Application.GetOpenFilename

If arq = "" Then
MsgBox "escolha um arquivo!"
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

    Columns("A:A").ColumnWidth = 100
    Columns("A:A").WrapText = True

 
' extrair Macro
'

'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC1,SEARCH(""R$"",RC[-1])+2,SEARCH("" "",RC1,(SEARCH(""R$"",RC[-1]))+2)-(SEARCH(""R$"",RC[-1])+2))"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC1,SEARCH(""R$"",RC1,SEARCH(RC[-1],RC1))+2,(SEARCH("" "",RC1,(SEARCH(""R$"",RC1,SEARCH(RC[-1],RC1)))+2))-(SEARCH(""R$"",RC1,SEARCH(RC[-1],RC1))+2))"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC1,SEARCH(""R$"",RC1,SEARCH(RC[-1],RC1))+2,(SEARCH("" "",RC1,(SEARCH(""R$"",RC1,SEARCH(RC[-1],RC1)))+2))-(SEARCH(""R$"",RC1,SEARCH(RC[-1],RC1))+2))"
    Range("B1:D1").Copy
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.SpecialCells(xlCellTypeFormulas, 16).Select
    Selection.ClearContents


' Formatar Macro
'

    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Columns("F:F").Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Rows("1:1").Insert Shift:=xlDown
    Columns("A:A").ColumnWidth = 1
    Range("B2").Value = "PG"
    Range("C2").Value = "Texto"
    Range("D2").Value = "Valor 1"
    Range("E2").Value = "Valor 2"
    Range("F2").Value = "Valor 3"
    
    Range("B3:F3").Select
    Range(Selection, Selection.End(xlDown)).Select

    With Selection.Borders
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
        End With
            
    Range("B2:F2").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .IndentLevel = 0
        .ReadingOrder = xlContext
        
    End With

    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    

' Copy_cola Macro
'

'
    Columns("D:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ' Numer_format Macro
'

'
    Range("D3:F78").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Range("B2").Select
    

' Macro2 Macro
Dim Vcel As Integer
Dim Vposicao As Integer
Dim Vfinal As Integer
Dim Vtamanho As Integer


Range("C3").Select
Vcel = 3

Do
    Vposicao = InStr(1, ActiveCell, "R$")
    If Vposicao > 0 Then
    Vfinal = InStr(Vposicao, ActiveCell, "million") + 8
    Vtamanho = Vfinal - Vposicao
    
    
    With ActiveCell.Characters(Start:=Vposicao, Length:=Vtamanho).Font
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
    End If
    
     Vposicao = InStr(Vfinal, ActiveCell, "R$")
     If Vposicao > 0 Then
    Vfinal = InStr(Vposicao, ActiveCell, "million") + 8
    Vtamanho = Vfinal - Vposicao
    

   With ActiveCell.Characters(Start:=Vposicao, Length:=Vtamanho).Font
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
    End If
    Vposicao = InStr(Vfinal, ActiveCell, "R$")
     If Vposicao > 0 Then
    Vfinal = InStr(Vposicao, ActiveCell, "million") + 8
    Vtamanho = Vfinal - Vposicao
    

   With ActiveCell.Characters(Start:=Vposicao, Length:=Vtamanho).Font
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
    End If
    
    Vcel = Vcel + 1
    Range("C" & Vcel).Select
    
 Loop Until Len(ActiveCell) = 0

' Macro2 Macro

Range("C3").Select
Vposição = 0
Vcel = 3
 
Do

    Vposicao = InStr(1, ActiveCell, "%")
    If Vposicao > 0 Then
    Vinicial = InStrRev(ActiveCell, " ", Vposicao)
    Vtamanho = Vposicao - Vinicial
    
    With ActiveCell.Characters(Start:=Vinicial, Length:=Vtamanho + 1).Font
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
    End If
    
    Vposicao = InStr(Vposicao + 1, ActiveCell, "%")
    If Vposicao > 0 Then
    Vinicial = InStrRev(ActiveCell, " ", Vposicao)
    Vtamanho = Vposicao - Vinicial
    

   With ActiveCell.Characters(Start:=Vinicial, Length:=Vtamanho + 1).Font
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
    End If
        Vposicao = InStr(Vposicao + 1, ActiveCell, "%")
    If Vposicao > 0 Then
    Vinicial = InStrRev(ActiveCell, " ", Vposicao)
    Vtamanho = Vposicao - Vinicial
    

   With ActiveCell.Characters(Start:=Vinicial, Length:=Vtamanho + 1).Font
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
    End If
    Vcel = Vcel + 1
    Range("C" & Vcel).Select
    
    Loop Until Len(ActiveCell) = 0
        

End Sub

