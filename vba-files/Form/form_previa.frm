VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_previa 
   Caption         =   "Prévia"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "form_previa.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_previa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    form_previa.Hide
    form_macros.Show
    Sheets("previa").Visible = False
    Sheets("previa-sp").Visible = False
    Sheets("previa-retira").Visible = False
    Sheets("previa-loja").Visible = False
    Sheets("previa-rj").Visible = False
    Application.Visible = False

End Sub

Private Sub CommandButton8_Click()
    
    Sheets("previa").Visible = True
    Sheets("previa").Select
    Sheets("previa-sp").Visible = True
    Sheets("previa-retira").Visible = True
    Sheets("previa-loja").Visible = True
    Sheets("previa-rj").Visible = True
    form_previa.Hide
    form_previa_dados_gerar.Show

End Sub

Private Sub CommandButton9_Click()
    
    Sheets("previa").Visible = True
    Sheets("previa").Select
    Sheets("previa-sp").Visible = True
    Sheets("previa-retira").Visible = True
    Sheets("previa-loja").Visible = True
    Sheets("previa-rj").Visible = True
    form_previa.Hide
    form_previa_dados_atualizar.Show
    
End Sub
