Attribute VB_Name = "inteiror_imprimir_cortes"
Sub imprimir()

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
    Sheets("interior_imprimir_cortes").Select
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
    Sheets("interior_imprimir_cortes").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    'Realiza as formatações.
    Range("C:C,D:D,F:F").Select
    Selection.Delete shift:=xlToLeft
    Columns("G:AE").Select
    Selection.Delete shift:=xlToLeft
    Columns("H:J").Select
    Selection.Delete shift:=xlToLeft
    Columns("E").Select
    Selection.Delete shift:=xlToLeft
    Range("F7").Select
    Selection.CurrentRegion.Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$F$1000"), , xlYes).Name = _
        "Tabela1"
    Range("Tabela1_1[#All]").Select
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=6, Criteria1:= _
        "=191*", Operator:=xlAnd
    ActiveSheet.ListObjects("Tabela1").TableStyle = "TableStyleLight8"
    Columns("A:A").ColumnWidth = 30
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 4
    Columns("D:D").ColumnWidth = 40
    Columns("E:E").ColumnWidth = 8
    Columns("F:F").ColumnWidth = 19
    ActiveSheet.ListObjects("Tabela1").ShowTotals = True
    Range("Tabela1_1[[#Totals],[PESO (KG)]]").Select
    ActiveSheet.ListObjects("Tabela1").ListColumns("PESO (KG)").TotalsCalculation _
        = xlTotalsCalculationSum
    Range("A1").Select
    Selection.CurrentRegion.Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
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
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
        
    aux = MsgBox("Imprimir? ", vbYesNo)
    
    If aux = vbYes Then
        Range("Tabela1_1[#All]").Select
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Preview:=True
    End If
    
    Windows(dadosNome).Close
    
End Sub


