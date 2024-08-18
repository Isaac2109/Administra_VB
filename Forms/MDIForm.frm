VERSION 5.00
Begin VB.MDIForm frmMDIPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Administra"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   15810
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu menuCadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu cadastroClientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu cadastroFornecedores 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu cadastroPrestadores 
         Caption         =   "&Prestadores"
      End
      Begin VB.Menu cadastroProdutos 
         Caption         =   "&Produtos"
      End
   End
End
Attribute VB_Name = "frmMDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cadastroClientes_Click()

    frmCadastroClientes.Show

End Sub

Private Sub cadastroProdutos_Click()

    FrmCadastroProdutos.Show

End Sub

