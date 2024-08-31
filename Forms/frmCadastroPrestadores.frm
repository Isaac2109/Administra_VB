VERSION 5.00
Begin VB.Form frmCadastroPrestadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro De Prestadores"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   9150
   Begin VB.TextBox txtEstado 
      Height          =   360
      Left            =   7680
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      Height          =   360
      Left            =   5880
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtNascimento 
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox cbmPessoa 
      Height          =   360
      ItemData        =   "frmCadastroPrestadores.frx":0000
      Left            =   7800
      List            =   "frmCadastroPrestadores.frx":000A
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtCodCliente 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtNomeCliente 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.TextBox txtCPF_CNPJ 
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtTelefone 
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtCEP 
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtEndereco 
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtCidade 
      Height          =   360
      Left            =   5640
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblEstado 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   7680
      TabIndex        =   21
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNascimento 
      Caption         =   "Nascimento:"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblPessoa 
      Caption         =   "Pessoa:"
      Height          =   255
      Left            =   7800
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblCodigoCliente 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblNomeCliente 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCPF_CNPJ 
      Caption         =   "CPF/CNPJ:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTelefone 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblCEP 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblEndereço 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCidade 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "frmCadastroPrestadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    ' Centraliza o formulário MDI Child dentro do MDI Form
    Me.Left = (frmMDIPrincipal.ScaleWidth - Me.Width) \ 2
    Me.Top = (frmMDIPrincipal.ScaleHeight - Me.Height) \ 2
    
End Sub


