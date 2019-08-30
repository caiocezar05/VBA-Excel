Attribute VB_Name = "M�dulo1"
Sub SalvarAba()
'para isso funcionar, � preciso que as duas planilhas estejam abertas e que essa macro esteja na


Dim Wb As Workbook
Dim Wt As Workbook
Dim nome As String
Dim id As String
Dim Ln As Integer

Set Wb = ThisWorkbook

'---- essa � a planilha que contem o template de cadas
Windows("formulario CSC.xlsx").Activate
'_____________________________________________________________


'essa linha ir� ativar a planilha de origem dos dados
Set Wt = ActiveWorkbook
Wb.Activate
'______________________________________________________________

'o template come�ar� a ser preenchido a partir da linha 2, por isso � necess�rio ter um cabe�alho
Range("A2").Select
Do Until Selection = Empty
Ln = ActiveCell.Row


'aqui escolheremos os campos que ser�o preenchidos, onde do lado esquerdo Wt � a lanilha de cadastro que receber� os dados e na direita
'como Wb ser� de onde a macro puxar� os dados
Wt.Sheets(1).Range("B2") = Wb.Sheets(1).Cells(Ln, 4)
Wt.Sheets(1).Range("B13") = Wb.Sheets(1).Cells(Ln, 5)
Wt.Sheets(1).Range("B15") = Wb.Sheets(1).Cells(Ln, 9)
Wt.Sheets(1).Range("B16") = Wb.Sheets(1).Cells(Ln, 10)
Wt.Sheets(1).Range("B18") = Wb.Sheets(1).Cells(Ln, 11)
Wt.Sheets(1).Range("B25") = Wb.Sheets(1).Cells(Ln, 2)


id = Wb.Sheets(1).Cells(Ln, 1)

Wt.Sheets(1).Name = id

'Aqui escolheremos o diret�rio onde as fichas de cadastro ser�o salvas
Wt.SaveAs "e:\e05774\Desktop\SAP\Exemplo\" & id & ".xls"

Selection.Offset(1, 0).Select
Loop

End Sub
