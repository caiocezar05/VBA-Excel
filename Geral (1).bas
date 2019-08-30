Attribute VB_Name = "CBRs_resume"
Sub DotoAllSheets()

'Faz um comando ser aplicado a todas as sheets de um workbook
Do

'----------------------------------------------------------------
'Coloque aqui a formula ou o comando que vc espera que
'seja aplicado para todas as as Sheets
'----------------------------------------------------------------

If ActiveSheet.Index <> Sheets.Count Then
ActiveSheet.Next.Select
Else
Exit Do
End If
Loop

End Sub
Sub ResumeSheets()
'Combinar tudo em uma sheet:
'Esse comando fará com que os dados de várias sheets sejam copiados para apenas uma

    Dim I As Long
    Dim xRg As Range
    On Error Resume Next
    Worksheets.Add Sheets(1)
    ActiveSheet.Name = "Resumo" 'mude o nome da sheet de resumo caso prefira


    
   For I = 2 To Sheets.Count
        Set xRg = Sheets(1).UsedRange
        If I > 2 Then
            Set xRg = Sheets(1).Cells(xRg.Rows.Count + 1, 1)
        End If
        Sheets(I).Activate
        ActiveSheet.UsedRange.Copy xRg
    Next
    
    Worksheets("Resumo").Select
    
    
'essa parte do código apagará todas as outras sheets após feito o resumo,
'caso vc queira deixar as demais planilhas, tire essas linhas

 For Each Worksheet In Worksheets
If ActiveSheet.Name = "Resumo" Then
ActiveSheet.Next.Select
End If
Application.DisplayAlerts = False
ActiveSheet.Delete
Next
End Sub

Sub SplitData()
'Esse comando separará em várias sheets os dados da primeira coluna da sheet selecionada
'exemplo: se tiver um razão com várias contas, ele criará uma aba para cada conta e dividirá os dados
     
             Set xSht = ActiveSheet
    On Error Resume Next
    xRCount = xSht.Cells(xSht.Rows.Count, 1).End(xlUp).Row
    xTitle = "A1:C1"
    xTRrow = xSht.Range(xTitle).Cells(1).Row
    For I = 2 To xRCount
        Call xCol.Add(xSht.Cells(I, 1).Text, xSht.Cells(I, 1).Text)
    Next
    xSUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    For I = 1 To xCol.Count
        Call xSht.Range(xTitle).AutoFilter(1, CStr(xCol.Item(I)))
        Set xNSht = Nothing
        Set xNSht = Worksheets(CStr(xCol.Item(I)))
        If xNSht Is Nothing Then
            Set xNSht = Worksheets.Add(, Sheets(Sheets.Count))
            xNSht.Name = CStr(xCol.Item(I))
        Else
            xNSht.Move , Sheets(Sheets.Count)
        End If
        xSht.Range("A" & xTRrow & ":A" & xRCount).EntireRow.Copy xNSht.Range("A1")
        xNSht.Columns.AutoFit
    Next
    xSht.AutoFilterMode = False
    xSht.Activate
    Application.ScreenUpdating = xSUpdate

End Sub
