VERSION 5.00
Begin VB.Form FrmCadastroProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7950
   Begin VB.CommandButton btnGravar 
      Caption         =   "V"
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnAlterar 
      Caption         =   "X"
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtEstoque 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton btnBuscarMarca 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox cbmSituacao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCadastroProdutos.frx":0000
      Left            =   6600
      List            =   "frmCadastroProdutos.frx":000A
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtObservacoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   12
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtPrecoSaida 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtPrecoEntrada 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtNomeGrupo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtCodGrupo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtNomeMarca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtCodMarca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtNomeProduto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox txtCodProduto 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblEstoque 
      Caption         =   "Estoque:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblSituacao 
      Caption         =   "Situa��o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Observa��es:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblPrecoSaida 
      Caption         =   "Pre�o Sa�da:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblPrecoEntrada 
      Caption         =   "Pre�o Entrada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblGrupo 
      Caption         =   "Grupo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblMarca 
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblNomeProduto 
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblCodigoProduto 
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "FrmCadastroProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inclusao As Boolean
Dim rsProdutos As New Recordset
Dim rsProdutosAberto As Boolean

Private Sub Form_Load()
    ' Centraliza o formul�rio MDI Child dentro do MDI Form
    Me.Left = (frmMDIPrincipal.ScaleWidth - Me.Width) \ 2
    Me.Top = (frmMDIPrincipal.ScaleHeight - Me.Height) \ 2
    
    inclusao = True
    rsProdutosAberto = False
    
    rsProdutos.Open "SELECT * FROM Produtos", ConexaoBD, adOpenDynamic, adLockOptimistic
    rsProdutosAberto = True
    
    rsProdutos.MoveLast
    
    txtCodProduto = rsProdutos("Codigo") + 1
    txtEstoque = 0
    
    rsProdutos.Close
    rsProdutosAberto = False
End Sub

Private Sub btnAlterar_Click()
    inclusao = False
End Sub

Private Sub btnGravar_Click()
    If gravarDados = True Then
        limparDados
    End If
End Sub

Private Function gravarDados() As Boolean
On Error GoTo Trataerro

    ' Verifica se o Recordset est� aberto antes de tentar abrir novamente
    If Not rsProdutosAberto Then
        rsProdutos.Open "SELECT * FROM Produtos", ConexaoBD, adOpenDynamic, adLockOptimistic
        rsProdutosAberto = True
    End If
    
    If inclusao = True Then
        rsProdutos.AddNew
    End If
    
    rsProdutos("Codigo") = VazioToNull(txtCodProduto)
    rsProdutos("Nome") = VazioToNull(txtNomeProduto)
    rsProdutos("Situacao") = VazioToNull(cbmSituacao)
    rsProdutos("CodigoMarca") = VazioToNull(txtCodMarca)
    rsProdutos("CodigoGrupo") = VazioToNull(txtCodGrupo)
    rsProdutos("PrecoEntrada") = VazioToNull(txtPrecoEntrada)
    rsProdutos("PrecoSaida") = VazioToNull(txtPrecoSaida)
    rsProdutos("Estoque") = VazioToNull(txtEstoque)
    rsProdutos("Observacoes") = VazioToNull(txtObservacoes)
    rsProdutos.Update
    
    rsProdutos.Close
    rsProdutosAberto = False
    gravarDados = True
    
    Exit Function
    
Trataerro:
    MsgBox "Erro nos dados informados", vbInformation, "ERRO"
    rsProdutos.CancelUpdate
    gravarDados = False
End Function


Private Sub limparDados()
    txtCodProduto = txtCodProduto + 1
    txtNomeProduto.Text = Empty
    cbmSituacao.Text = Empty
    txtCodMarca.Text = Empty
    txtCodGrupo.Text = Empty
    txtPrecoEntrada.Text = Empty
    txtEstoque.Text = 0
    txtPrecoSaida.Text = Empty
    txtObservacoes.Text = Empty
End Sub







