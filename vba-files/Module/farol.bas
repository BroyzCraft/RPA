Attribute VB_Name = "farol"
Sub importar()
    
    Application.DisplayAlerts = False
    
    Dim macro As String
    Dim dadosNome As String
    Dim dados As Variant
    
    'reseta os dados
    macro = ActiveWorkbook.Name
    Windows(macro).Activate
    Sheets("farol-dados").Select
    Range("A1").Select
    If ActiveSheet.FilterMode Then
        Range("A1").AutoFilter
    End If
    Cells.Select
    Selection.Delete shift:=xlUp
    
    'procura e abre o arquivo de dados
    MsgBox ("Selecione a planilha com os dados: ")
    dados = Application.GetOpenFilename
    If dados = False Then Exit Sub
    Workbooks.Open dados, , True
    dadosNome = ActiveWorkbook.Name
    
    Windows(dadosNome).Activate
    Cells.Copy
    Windows(macro).Activate
    Sheets("farol-dados").Select
    Range("A1").PasteSpecial xlPasteAll
    Windows(dadosNome).Close False
    
    'ajuste
    Sheets("farol-resumo").Select
    Range("A1").Select
    ActiveSheet.PivotTables("STATUS DE ROTA").PivotCache.Refresh
    ActiveSheet.PivotTables("STATUS DE ENTREGA").PivotCache.Refresh
    Columns("C:F").Select
    Selection.ColumnWidth = 13
    Columns("I:L").Select
    Selection.ColumnWidth = 13
    Range("A49").Select
    
End Sub
