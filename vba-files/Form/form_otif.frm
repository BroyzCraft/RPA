VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_otif 
   Caption         =   "OTIF"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "form_otif.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_otif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    form_otif.Hide
    form_macros.Show
    Sheets("otif-dados").Visible = False
    Sheets("otif-menu").Visible = False
    Sheets("otif-resumo").Visible = False
    Sheets("otif-consolidado").Visible = False
    Sheets("otif-filhos").Visible = False
    Application.Visible = False
    
End Sub

Private Sub CommandButton10_Click()
    
    Sheets("otif-dados").Visible = True
    Sheets("otif-menu").Visible = True
    Sheets("otif-resumo").Visible = True
    Sheets("otif-consolidado").Visible = True
    Sheets("otif-filhos").Visible = True
    
    'Worksheets(Array("otif-resumo", "otif-consolidado", "otif-filhos")).Copy
    'With ActiveWorkbook
    '     .SaveAs Filename:=Environ("USERPROFILE") & "\Desktop\OTIF.xlsx", FileFormat:=xlOpenXMLWorkbook
    '     .Close SaveChanges:=True
    'End With
    
    otif.gerarBackup
    Shell "C:\WINDOWS\explorer.exe """ & "\\Ecfs1\leo\Logistica\Transporte\1_TRANSPORTES\Controle de Diario\FECHAMENTO GERAL\FECHAMENTOS 2021\Fechamento On time + In Full\" & "", vbNormalFocus
    
End Sub

Private Sub CommandButton12_Click()
    
    otif.coletarInformacoes
    
End Sub

Private Sub CommandButton8_Click()
    
    MsgBox ("Preencha com as informações de reentrega e devolução")
    Sheets("otif-consolidado").Visible = True
    Sheets("otif-consolidado").Select
    
End Sub

Private Sub CommandButton9_Click()
    
    MsgBox ("Preencha com as informações de pedidos filhos")
    Sheets("otif-filhos").Visible = True
    Sheets("otif-filhos").Select

End Sub

Private Sub CommandButton11_Click()
    
    MsgBox ("Aguarde a finalização e preencha o 'otif-resumo' com os dados que serão apresentados")
    MsgBox ("Após a atualização, gere a planilha do OTIF nesse menu")
    otif.AtualizarDados

End Sub


