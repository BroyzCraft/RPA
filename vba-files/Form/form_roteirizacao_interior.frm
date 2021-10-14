VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_roteirizacao_interior 
   Caption         =   "Roteirização - Interior"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "form_roteirizacao_interior.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_roteirizacao_interior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    form_roteirizacao_interior.Hide
    form_macros.Show
    Sheets("interior_imprimir_cortes").Visible = False
    Sheets("interior_imprimir_cortes").Visible = False
    
End Sub

Private Sub CommandButton8_Click()
    
    Sheets("interior_organizar_rotas").Visible = True
    interior_organizar_rotas.organizar

End Sub

Private Sub CommandButton9_Click()
    
    Sheets("interior_imprimir_cortes").Visible = True
    inteiror_imprimir_cortes.imprimir
    
End Sub
