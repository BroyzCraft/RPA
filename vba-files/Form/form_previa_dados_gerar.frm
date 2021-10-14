VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_previa_dados_gerar 
   Caption         =   "Dados de Usuario"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "form_previa_dados_gerar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_previa_dados_gerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    previa.gerar
    
End Sub

Private Sub datafim_Change()
    
    If datafim.SelStart = 2 Then datafim.SelText = "/"
    If datafim.SelStart = 5 Then datafim.SelText = "/"
    
End Sub

Private Sub datainicio_Change()

    If datainicio.SelStart = 2 Then datainicio.SelText = "/"
    If datainicio.SelStart = 5 Then datainicio.SelText = "/"
    
End Sub
