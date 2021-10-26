VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Count = 1 Then
            
            Application.Visible = False
    
            'fechar abas planilhas
            Dim quant As Integer
            Dim x As Integer
            quant = ActiveWorkbook.Worksheets.Count
            For x = 1 To quant
                If ActiveWorkbook.Worksheets(x).Name <> "Menu" Then
                    ActiveWorkbook.Worksheets(x).Visible = False
                End If
            Next x
            
            Call form_macros.Show
            
    End If
End Sub

Private Sub Worksheet_Activate()
    
    Application.Visible = False
    
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
