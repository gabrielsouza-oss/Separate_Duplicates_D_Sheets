Attribute VB_Name = "M�dulo1"
Public Sub Separate_Duplicates()

Application.ScreenUpdating = False

Dim contador() As Long
Dim Contadores(300) As Long ' Set the Max Number of Sheets to be added
Dim Rng As Range
Dim ws2 As Worksheet
Dim ws As Worksheet
Dim lLastRow As Long
Dim lColumn As Long

Col = Application.InputBox("Enter the Relative Positional of Key Column on Dataframe", Type:=1)
Columns_To_Copy = Application.InputBox("Enter the Amount of Fields to be copied on the Left of the Key Column", Type:=1)

Set ws = ThisWorkbook.Sheets(1)
lLastRow = ws.Cells(Rows.Count, Col).End(xlUp).Row
lColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
ReDim contador(lLastRow)

contador(0) = 0

'Ordenando os Dados de forma Ascendente de acordo com a KEY

ws.Range(Cells(1, 1).Address, Cells(lLastRow, lColumn).Address).Sort key1:=ws.Range(Cells(1, Col).Address, Cells(lLastRow, Col).Address), order1:=xlAscending, Header:=xlYes
   
'Adicionando o Numero de Sheets necessarias para separar os dados

For i = 2 To lLastRow
    If ws.Cells(i, Col).Text = ws.Cells(i + 1, Col).Text And ws.Cells(i - 1, Col).Text = ws.Cells(i, Col).Text Then
        contador(i) = contador(i - 1) + 1
    Else
        contador(i) = 1
    End If
Next i

For i = 2 To lLastRow
    maximo = contador(i)
    For j = 2 To lLastRow
        If maximo < contador(j) Then
            maximo = contador(j)
        End If
    Next j
Next i

If maximo = 0 Then
    MsgBox ("No Duplicate Keys Where Found, Exiting")
    Exit Sub
End If

MsgBox (maximo)

ActiveWorkbook.Sheets.Add Count:=maximo, After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

paste_count = 2

'Separando os Dados

For i = 2 To lLastRow
    If ws.Cells(i, Col).Text = ws.Cells(i + 1, Col).Text Then
        Set ws2 = ThisWorkbook.Sheets(contador(i) + 1)
        
        'Copiando o Header da Tabela
        
        Set Rng = ws.Range(Cells(1, 1).Address, Cells(1, lColumn).Address)
        Rng.Copy Destination:=ws2.Range(Cells(1, 1).Address, Cells(1, lColumn).Address)
        
        'Isolando as Duplicatas em outras Sheets
        
        Set Rng = ws.Range(Cells(i + 1, Col).Address, Cells(i + 1, Col - Columns_To_Copy).Address)
        Rng.Copy Destination:=ws2.Range(Cells(paste_count, Col).Address, Cells(paste_count, Col - Columns_To_Copy).Address)
        paste_count = paste_count + 1
        
    End If
Next i

'Subindo as linhas das Sheets
For i = 2 To ActiveWorkbook.Worksheets.Count
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(i)
    ws.Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    On Error GoTo 0
Next i


'Copiando o Dataframe Original
Set ws = ThisWorkbook.Sheets(1)
ws.Copy After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

'Finalizando a Remoção das Duplicatas do DF Original
ws.Range(Cells(1, 1).Address, Cells(lLastRow, lColumn).Address).RemoveDuplicates Columns:=Array(Col), Header:=xlYes

Application.ScreenUpdating = True

End Sub
