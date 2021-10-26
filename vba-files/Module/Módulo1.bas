Attribute VB_Name = "Módulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("A1").Select
    ActiveSheet.PivotTables("STATUS DE ROTA").PivotCache.Refresh
    ActiveSheet.PivotTables("STATUS DE ENTREGA").PivotCache.Refresh
    Columns("C:F").Select
    Selection.ColumnWidth = 15
    Columns("I:L").Select
    Selection.ColumnWidth = 15
    
End Sub
