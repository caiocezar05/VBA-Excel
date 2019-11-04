Attribute VB_Name = "RDA_Convênios"

Sub Inserir_informações()
Attribute Inserir_informações.VB_ProcData.VB_Invoke_Func = " \n14"
'
' importação
'
    Range("A1").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\caio.santos\Desktop\text.txt", Destination:=Range("$A$1"))
        .Name = "text"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = ""
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Previous.Select
    
    'Seleção
Range("A1").Select

'
    Cells.Find(What:="Identificação do Projeto:", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B2").Select

    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("B2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(25, 1)), TrailingMinusNumbers:=True
    ActiveSheet.Previous.Select

'
' instituição
'
    Cells.Find(What:="Instituição", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Selection.Offset(1, 0).Select
    Selection.Cut
    ActiveSheet.Next.Select
    
    Range("B3").Select
    ActiveSheet.Paste
    ActiveSheet.Previous.Select

' Início e fim do projeto
'

'
    Cells.Find(What:= _
        "Data de Início do Projeto Data do fim do Projeto UF de Execução do Projeto", _
        After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:= _
        xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False) _
        .Activate
        
    Range(Selection, Selection.Offset(1, 0)).Select
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B5").Select
    ActiveSheet.Paste
    Range("B5").Select
    Selection.TextToColumns Destination:=Range("B5"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(25, 1), Array(48, 1)), TrailingMinusNumbers _
        :=True
    Range("B6").Select
    Selection.TextToColumns Destination:=Range("B6"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1), Array(21, 1)), TrailingMinusNumbers _
        :=True
    Range("B5:D6").Select
     Selection.Replace What:=" do projeto", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveSheet.Previous.Select

'
' coordenador do projeto
'

'
    Cells.Find(What:="Coordenador ou Responsável,", After:=ActiveCell, LookIn _
        :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    
    Range(Selection, Selection.Offset(1, 0)).Select
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B8").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.TextToColumns Destination:=Range("B8"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(26, 9)), TrailingMinusNumbers:=True
    

    

    
    Range("C9").Select
        ActiveCell.FormulaR1C1 = _
        "=LEFT(RC[-1],SEARCH("" "",RC[-1],(SEARCH("" "",RC[-1])) +1 ))"
    
    Selection.Copy
    Selection.Offset(0, -1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Offset(0, 1).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
'cabeçalhos
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "Viagens"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = " Obras Civis"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "Material de Consumopara Protótipo"
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "Equipamentos e Acessórios, Bens de Informática"
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "Treinamento"
    Range("C13").Select
    ActiveCell.FormulaR1C1 = " Software"
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "Material de Consumo"
    Range("E13").Select
    ActiveCell.FormulaR1C1 = "Equipamentos e" & Chr(10) & "Acessórios, Outros"
    Range("B15").Select
    ActiveCell.FormulaR1C1 = "Custo Incorrido pela" & Chr(10) & "Instituição"
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "Outros Correlatos: rateio de infra-estrutura da Instituição"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "Outros Correlatos"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "Livros/Periódicos"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = " Serviços Técnicos de" & Chr(10) & "Terceiros - Outros"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = " Serviços Técnicos de" & Chr(10) & "Terceiros - Outros"
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "Serviços Técnicos de Terceiros -Tecnológicos"
    Range("E17").Select
    ActiveCell.FormulaR1C1 = "Total de dispêndios"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "RH"
    Range("B20").Select
    ActiveSheet.Previous.Select
    
'primeira linha
    Range("A1").Select
        Cells.Find(What:="art 25", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Selection.ClearContents
    Selection.Offset(1, 0).Select
    Selection.Cut

    ActiveSheet.Next.Select
    Range("B12").Select
    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("B12"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), TrailingMinusNumbers:= _
        True
    ActiveSheet.Previous.Select
    
'segunda linha
 Range("A1").Select
    Cells.Find(What:="art 25", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Selection.ClearContents
    Selection.Offset(1, 0).Select
    Selection.Cut

    ActiveSheet.Next.Select
    Range("B14").Select
    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("B14"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), TrailingMinusNumbers:= _
        True
    ActiveSheet.Previous.Select
    
    ' Terceira linha
'

            Cells.Find(What:="Outros Correlatos", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
        
        '*mudar para proxima usada
    Selection.Offset(2, 0).Select
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B16").Select
    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("B16"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
        ActiveSheet.Previous.Select
        
        'quarta linha
        
   Cells.Find(What:="Total de dispêndios", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
        
        '*mudar para proxima usada
    Selection.Offset(2, 0).Select
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B18").Select
    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("B18"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
        ActiveSheet.Previous.Select
        
'        RH
'


    Cells.Find(What:="Valor (R$) ", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B20").Select
    ActiveSheet.Paste
    
 
   
    Selection.TextToColumns Destination:=Range("B20"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 9), Array(2, 9), Array(3, 9), Array(4, 9), Array(5, 9), Array(6, 9), _
        Array(7, 1)), TrailingMinusNumbers:=True
    
      'Outros
    ActiveSheet.Previous.Select
        Cells.Find(What:= _
        "Valor Total Repassado", After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Range(Selection, Selection.Offset(6, 0)).Select
    Selection.Cut
    ActiveSheet.Next.Select
    Range("B22").Select
    ActiveSheet.Paste
    ActiveSheet.Previous.Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete

'
' Acabamento
'

' Fazer os números ficarem em colunas

    Cells.Select
    Range("B22:H40").Activate
        Selection.WrapText = False

    Range("B5:D6").Select
    Selection.Copy
    Range("H5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B8:B9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H9").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B11:E12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H11").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B13:E14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B15:D16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H19").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B17:E18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H22").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B19:B20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H26").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B22:B28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H28").Select
    ActiveSheet.Paste
    Range("B5:G34").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("B2").Select

' Deletar o que estiver zerado
'

'
    Selection.TextToColumns Destination:=Range("B28:B34"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

    Range("C11:C34").Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("B2").Select

' Fit collumns
'

'
    Columns("A:A").Select
    Selection.ColumnWidth = 1
    Columns("B:B").ColumnWidth = 35.57
    Columns("B:B").ColumnWidth = 42
    Range("B1:C1").Select
    Range("C1").Activate
    Columns("B:B").ColumnWidth = 48.57
    Columns("B:B").ColumnWidth = 51.43
    Columns("B:B").ColumnWidth = 42
    Columns("C:C").EntireColumn.AutoFit
    Range("B2").Select
    
    ' Acabamento final Ticks
''
    Rows("1:1").Select
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
    Selection.Insert Shift:=x1Down
'
    Range("B1:B10").Select
      Selection.HorizontalAlignment = xlRight
    With Selection.Font
        .Color = -10477568
        .Bold = True

        .TintAndShade = 0
    End With


    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Cliente:"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Escopo:"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Data base:"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Objetivo:"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Procedimentos:"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "Conclusão:"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Huawei do Brasil LTDA"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Auditoria RDA"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "12/31/2017"
    Range("C5:f5").Select
    ActiveCell.FormulaR1C1 = _
        "Obter o confronto das informações contidas no RDA e no contrato das instituições."
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("C7:f7").Select
    ActiveCell.FormulaR1C1 = _
        "Realizamos a leitura do contrato firmado (Plano de Trabalho) com a instituição conveniada e confrontamos com as informações do RDA. Utilizamos com suporte o controle análitico fornecido pela Huawei."
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("C9:f10").Select
    ActiveCell.FormulaR1C1 = _
        "Comparando as informações declaradas na RDA sobre os convênios com a documentação suporte e/ou controles internos da Huawei, não há pontos relevantes a ressalvar, com excessão do exposto da sheet 8. Resumo."
With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With

'
    Range("11:11").Select
    Selection.Insert Shift:=x1Down
    Range("B12").Select
    Range(Selection, Selection.Offset(0, 3)).Select
     With Selection.Font
        .Name = "Bookshelf Symbol 7"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
        Selection.HorizontalAlignment = xlCenter
        Selection.Font.Bold = True
     With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "d"
    ActiveCell.Offset(0, 1).Select
    Selection = "d"
    ActiveCell.Offset(0, 1).Select
    Selection = "p"
    ActiveCell.Offset(0, 1).Select
    Selection = "o"
    Range(Selection.Offset(1, 0), Selection.Offset(2, 0)).Select
    Selection.EntireRow.Select
    Selection.Font.Bold = True

 ' bordas
    Range("B150").Select
Selection.End(xlUp).Select
Range("B13", Selection.Offset(0, 3)).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
        Range("B12:E12").Select

    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    
' ticks e legendas
    Range("B150").Select
    Selection.End(xlUp).Offset(2, 0).Select
    Selection = "Tick Marks"
    Selection.HorizontalAlignment = xlRight
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Bold = True
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range(Selection.Offset(1, 0), Selection.Offset(3, 0)).Select
    

    With Selection.Font
        .Name = "Bookshelf Symbol 7"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
        Selection.HorizontalAlignment = xlRight
        Selection.Font.Bold = True
    
    
    ActiveCell.FormulaR1C1 = "d"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Segundo RDA"
            Selection.Font.Bold = True
    ActiveCell.Offset(1, -1).Select
    
        ActiveCell.FormulaR1C1 = "p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Segundo contrato"
    Selection.Font.Bold = True
    ActiveCell.Offset(1, -1).Select
    
ActiveCell.FormulaR1C1 = "o"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Segundo planilha de pagamento"
    Selection.Font.Bold = True
    ActiveCell.Offset(1, -1).Select
    
    Range("B2").Select
    'cereja do bolo
    ActiveWindow.DisplayGridlines = False
End Sub


