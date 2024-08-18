VERSION 5.00
Begin VB.Form frmCadastroClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11895
End
Attribute VB_Name = "frmCadastroClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    ' Centraliza o formulário MDI Child dentro do MDI Form
    Me.Left = (frmMDIPrincipal.ScaleWidth - Me.Width) \ 2
    Me.Top = (frmMDIPrincipal.ScaleHeight - Me.Height) \ 2
    
End Sub

