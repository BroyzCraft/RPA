Attribute VB_Name = "interior_organizar_rotas"
Sub organizar()

    Dim macro As String
    Dim dadosNome As String
    Dim dados As Variant

    macro = ActiveWorkbook.Name
    
    'procura e abre o arquivo de dados
    MsgBox ("Selecione a planilha com os dados da roteirização: ")
    dados = Application.GetOpenFilename
    If dados = False Then Exit Sub
    Workbooks.Open dados, , True
    dadosNome = ActiveWorkbook.Name
    
    'reseta os dados
    Windows(macro).Activate
    Sheets("interior_organizar_rotas").Select
    Range("A1").Select
    If ActiveSheet.FilterMode Then
        Range("A1").AutoFilter
    End If
    Cells.Select
    Selection.Delete shift:=xlUp
    
    'coleta os dados novos
    Windows(dadosNome).Activate
    Cells.Select
    Selection.Copy
    Windows(macro).Activate
    Sheets("interior_organizar_rotas").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    'Realiza as formatações.
    Range("B:D,F:F,J:AH,AJ:AJ").Select
    Selection.Delete shift:=xlToLeft
    Range("A1").Select
    Selection.CurrentRegion.Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$F$500"), , xlYes).Name = _
        "Tabela1"
    Columns("A:A").ColumnWidth = 25
    Columns("B:B").ColumnWidth = 6
    Columns("C:C").ColumnWidth = 38
    Columns("D:D").ColumnWidth = 60
    Columns("F:F").ColumnWidth = 20
    ActiveSheet.ListObjects("Tabela1").ShowTotals = True
    ActiveSheet.ListObjects("Tabela1").TableStyle = "TableStyleMedium1"
    Range("Tabela1[[#Totals],[PESO (KG)]]").Select
    ActiveSheet.ListObjects("Tabela1").ListColumns("PESO (KG)").TotalsCalculation _
        = xlTotalsCalculationSum
    Range("A1").Select
    
    Windows(dadosNome).Close
    
End Sub


