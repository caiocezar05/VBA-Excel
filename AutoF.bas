Attribute VB_Name = "Módulo1"
Sub SalvarAba()
'para isso funcionar, é preciso que as duas planilhas estejam abertas


Dim Wb As Workbook
Dim Wt As Workbook
Dim nome As String
Dim id As String
Dim Ln As Integer

Set Wb = ThisWorkbook

'---- essa é a planilha que contem o template de cadas
Windows("formulario CSC.xlsx").Activate
'_____________________________________________________________


'essa linha irá ativar a planilha de origem dos dados
Set Wt = ActiveWorkbook
Wb.Activate
'______________________________________________________________

'o template começará a ser preenchido a partir da linha 2, por isso é necessário ter um cabeçalho
Range("A2").Select
Do Until Selection = Empty
Ln = ActiveCell.Row


'aqui escolheremos os campos que serão preenchidos, onde do lado esquerdo Wt é a lanilha de cadastro que receberá os dados e na direita
'como Wb será de onde a macro puxará os dados
Wt.Sheets(1).Range("B2") = Wb.Sheets(1).Cells(Ln, 4)
Wt.Sheets(1).Range("B13") = Wb.Sheets(1).Cells(Ln, 5)
Wt.Sheets(1).Range("B15") = Wb.Sheets(1).Cells(Ln, 9)
Wt.Sheets(1).Range("B16") = Wb.Sheets(1).Cells(Ln, 10)
Wt.Sheets(1).Range("B18") = Wb.Sheets(1).Cells(Ln, 11)
Wt.Sheets(1).Range("B25") = Wb.Sheets(1).Cells(Ln, 2)


id = Wb.Sheets(1).Cells(Ln, 1)

Wt.Sheets(1).Name = id

'Aqui escolheremos o diretório onde as fichas de cadastro serão salvas
Wt.SaveAs "e:\e05774\Desktop\SAP\Exemplo\" & id & ".xls"

Selection.Offset(1, 0).Select
Loop

End Sub
