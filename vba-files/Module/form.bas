Attribute VB_Name = "form"
Sub abreform()

    'fechar abas planilhas
    Dim quant As Integer
    Dim x As Integer
    quant = ActiveWorkbook.Worksheets.Count
    For x = 1 To quant
        If ActiveWorkbook.Worksheets(x).Name <> "Menu" Then
            ActiveWorkbook.Worksheets(x).Visible = False
        End If
    Next x
    
    'abre menu
    form_macros.Show

End Sub
