Attribute VB_Name = "otif"
Sub AtualizarDados()
    
    Application.DisplayAlerts = False
    Dim data
    Sheets("otif-dados").Visible = True
    Sheets("otif-menu").Visible = True
    Sheets("otif-resumo").Visible = True
    Sheets("otif-consolidado").Visible = True
    Sheets("otif-filhos").Visible = True
    
    Sheets("otif-dados").Select
    Range("C:Z").Delete
    data = Format(Date, "ddmmyyyy")
    Sheets("otif-dados").Select
    Range("A1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    ActiveSheet.ListObjects("otif_remessas_2").Range.AutoFilter Field:=1, Criteria1:="=*" & data & "*", Operator:=xlAnd
    Columns("B:B").Select
    Selection.Copy
    Range("F1").Select
    ActiveSheet.Paste
    ActiveSheet.ListObjects("otif_remessas_2").Range.AutoFilter Field:=1
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("F1:Z1").Select
    Selection.FormulaR1C1 = "=COUNTA(R[1]C:R[99999]C)"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[21])"
    Sheets("otif-menu").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "='otif-dados'!R[-1]C[3]"
    Range("B3").Select
    data = Format(Date, "dd/mm/yyyy")
    ActiveSheet.PivotTables("otif_consolidado").PivotCache.Refresh
    ActiveSheet.PivotTables("otif_consolidado").PivotFields("DATA").ClearAllFilters
    ActiveSheet.PivotTables("otif_consolidado").PivotFields("DATA").PivotFilters. _
        Add2 Type:=xlSpecificDate, Value1:=data
    Application.DisplayAlerts = True
    data = Format(Date, "dd.mm.yyyy")
    ActiveSheet.PivotTables("otif_filhos").PivotCache.Refresh
    ActiveSheet.PivotTables("otif_filhos").PivotFields("Data").ClearAllFilters
    ActiveSheet.PivotTables("otif_filhos").PivotFields("Data").PivotFilters.Add2 _
        Type:=xlCaptionEquals, Value1:=data
    
End Sub

Sub gerarBackup()
    
    Dim plan As String
    Dim macro As String
    plan = "10. Fechamento diario - Outubro.xlsx"
    macro = "RPAs - Bruno.xlsm"
    Workbooks.Open ("\\Ecfs1\leo\Logistica\Transporte\1_TRANSPORTES\Controle de Diario\FECHAMENTO GERAL\FECHAMENTOS 2021\Fechamento On time + In Full\" & plan)
    
    Workbooks(macro).Activate
    Sheets("otif-resumo").Select
    Cells.Copy
    Workbooks(plan).Activate
    Sheets("otif-resumo").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Workbooks(macro).Activate
    Sheets("otif-consolidado").Select
    Cells.Copy
    Workbooks(plan).Activate
    Sheets("otif-consolidado").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Workbooks(macro).Activate
    Sheets("otif-filhos").Select
    Cells.Copy
    Workbooks(plan).Activate
    Sheets("otif-filhos").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Workbooks(plan).Close (True)
    
End Sub
