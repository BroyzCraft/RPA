VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_rj 
   Caption         =   "Roteirização RJ"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "form_rj.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_rj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    form_rj.Hide
    form_macros.Show
    Sheets("rj-menu").Visible = False
    Sheets("rj-controle").Visible = False
    Sheets("rj-capa-corte").Visible = False
    Application.Visible = False
    
End Sub

Private Sub CommandButton10_Click()

End Sub

Private Sub CommandButton11_Click()
    rj.imprimirControle
End Sub

Private Sub CommandButton8_Click()
    rj.apagar
End Sub

Private Sub CommandButton9_Click()
     rj.imprimirCortes
End Sub
