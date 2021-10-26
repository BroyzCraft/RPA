VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_macros 
   Caption         =   "Menu de Macros"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "form_macros.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_macros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    Application.Visible = True
    form_macros.Hide
    form_roteirizacao_interior.Show

End Sub

Private Sub CommandButton2_Click()
    
    Application.Visible = True
    Sheets("otif-dados").Visible = True
    Sheets("otif-menu").Visible = True
    Sheets("otif-resumo").Visible = True
    Sheets("otif-consolidado").Visible = True
    Sheets("otif-filhos").Visible = True
    form_macros.Hide
    form_otif.Show
    
End Sub

Private Sub CommandButton3_Click()
    
    Application.Visible = True
    form_macros.Hide
    form_previa.Show

End Sub

Private Sub CommandButton5_Click()
    
    Application.Visible = True
    Sheets("farol-resumo").Visible = True
    Sheets("farol-dados").Visible = True
    farol.importar
    Sheets("farol-resumo").Select
    
End Sub

Private Sub CommandButton6_Click()

    Application.Visible = True
    form_rj.Show
    form_macros.Hide
    Sheets("rj-menu").Visible = True
    Sheets("rj-controle").Visible = True
    Sheets("rj-capa-corte").Visible = True
    Sheets("rj-menu").Select
    
End Sub
