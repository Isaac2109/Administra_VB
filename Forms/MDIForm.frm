VERSION 5.00
Begin VB.MDIForm frmMDIPrincipall 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm"
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
   End
End
Attribute VB_Name = "frmMDIPrincipall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

    'Conexao com Banco de Dados
    cnnBD.ConnectionString = "Provider=SQLOLEDB;Data Source=ISAAC-PC\SQLEXPRESS,1433;Initial Catalog=Administra_VB;Password=Lrsiazevedo2023@;User ID=sa;"
    cnnBD.Open
    
    MsgBox " Conexão efetuada com sucesso. Seja Bem Vindo!! "

End Sub
