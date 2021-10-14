VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

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

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ActiveWorkbook.Close False
    
End Sub
