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
   Begin VB.CommandButton btnNovo 
      Caption         =   "N"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnInicio 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnAnterior 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnAvan�ar 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnFinal 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnPesquisarProduto 
      Caption         =   "B"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnGravar 
      Caption         =   "G"
      Height          =   375
      Left            =   4560
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
   Begin VB.CommandButton btnPesquisarGrupo 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton btnPesquisarMarca 
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
      Text            =   "0,00"
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
      Text            =   "0,00"
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
   Begin VB.Label lblObservacoes 
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
Dim rsProdutos As New Recordset
Dim rsMarca As New Recordset
Dim rsGrupo As New Recordset
Dim rsProdutosAberto As Boolean
Dim sql As String
Dim inclusao As Boolean

Private Sub Form_Load()
    ' Centraliza o formul�rio MDI Child dentro do MDI Form
    Me.Left = (frmMDIPrincipal.ScaleWidth - Me.Width) \ 2
    Me.Top = (frmMDIPrincipal.ScaleHeight - Me.Height) \ 2
    
    inclusao = True
    
    'CHAMAR FUN��O DE INCLUSAO
    incluir
    
End Sub

Private Sub btnAlterar_Click()
    
    inclusao = False
    
    btnPesquisarProduto.Enabled = True
    btnNovo.Enabled = True
    btnAvan�ar.Enabled = True
    btnFinal.Enabled = True
    btnAnterior.Enabled = True
    btnInicio.Enabled = True
    
    alteracao
        
End Sub

Private Sub btnNovo_Click()

    inclusao = True
        
    btnPesquisarProduto.Enabled = False
    btnNovo.Enabled = False
    btnAvan�ar.Enabled = False
    btnFinal.Enabled = False
    btnAnterior.Enabled = False
    btnInicio.Enabled = False
    
    limparCampos
    
    incluir
    
End Sub

Private Sub btnInicio_Click()

    rsProdutos.MoveFirst
    preencherCampos

End Sub

Private Sub btnAnterior_Click()

    If rsProdutos("Codigo") <> 0 Then
        rsProdutos.MovePrevious
        preencherCampos
    End If

End Sub

Private Sub btnAvan�ar_Click()

    rsProdutos.MoveNext
    If rsProdutos.EOF = True Then
        rsProdutos.MovePrevious
    End If
    preencherCampos

End Sub

Private Sub btnFinal_Click()

    rsProdutos.MoveLast
    preencherCampos

End Sub

Private Sub btnGravar_Click()
    If gravarDados = True Then
        limparCampos
        incluir
    End If
End Sub

Private Sub incluir()

    If rsProdutosAberto = True Then
        rsProdutos.Close
    End If
    
    'Setar Pr�ximo C�digo do Novo Produto
    rsProdutos.Open "SELECT * FROM Produtos", ConexaoBD, adOpenDynamic, adLockOptimistic
    rsProdutosAberto = True
    
    rsProdutos.MoveLast
    txtCodProduto = rsProdutos("Codigo") + 1
    txtEstoque = 0

End Sub

Private Sub alteracao()

    If rsProdutosAberto = True Then
        rsProdutos.Close
    End If
    
    sql = "SELECT Prod.Codigo, Prod.Nome, CodigoMarca, Marcas.Nome NomeMarca, " & _
          "CodigoGrupo, Grupos.Nome NomeGrupo, PrecoEntrada, PrecoSaida, Estoque, Situacao, Observacoes " & _
          "FROM Produtos Prod " & _
          "LEFT JOIN Marcas on Prod.CodigoMarca = Marcas.Codigo " & _
          "LEFT JOIN Grupos on Prod.CodigoGrupo = Grupos.Codigo "

    rsProdutos.Open sql, ConexaoBD, adOpenDynamic, adLockOptimistic
    rsProdutosAberto = True
    
    rsProdutos.MoveFirst
    preencherCampos
    
End Sub

'Gravar RecordSet
Private Function gravarDados() As Boolean
On Error GoTo Trataerro

    If rsProdutosAberto = True Then
        rsProdutos.Close
    End If

    rsProdutos.Open "SELECT * FROM Produtos", ConexaoBD, adOpenDynamic, adLockOptimistic
    rsProdutosAberto = True
    
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
    
    gravarDados = True
    
    Exit Function
    
Trataerro:
    MsgBox "Erro nos dados informados", vbInformation, "ERRO"
    rsProdutos.CancelUpdate
    gravarDados = False
End Function

'Limpar dados dos txtBox
Private Sub limparCampos()
    txtCodProduto = txtCodProduto + 1
    txtNomeProduto.Text = Empty
    cbmSituacao.Text = Empty
    txtCodMarca.Text = Empty
    txtNomeMarca.Text = Empty
    txtCodGrupo.Text = Empty
    txtNomeGrupo.Text = Empty
    txtPrecoEntrada.Text = Empty
    txtEstoque = 0
    txtPrecoSaida.Text = Empty
    txtObservacoes.Text = Empty
End Sub

Private Sub preencherCampos()

    txtCodProduto = NullToVazio(rsProdutos("Codigo"))
    txtNomeProduto = NullToVazio(rsProdutos("Nome"))
    cbmSituacao = NullToVazio(rsProdutos("Situacao"))
    txtCodMarca = NullToVazio(rsProdutos("CodigoMarca"))
    txtNomeMarca = NullToVazio(rsProdutos("NomeMarca"))
    txtCodGrupo = NullToVazio(rsProdutos("CodigoGrupo"))
    txtNomeGrupo = NullToVazio(rsProdutos("NomeGrupo"))
    
    ' Formata o valor do banco para exibir com 2 casas decimais
    txtPrecoEntrada = Format(rsProdutos("PrecoEntrada"), "#,##0.00")
    txtPrecoSaida = Format(rsProdutos("PrecoSaida"), "#,##0.00")
    
    txtEstoque = NullToVazio(rsProdutos("Estoque"))
    txtObservacoes = NullToVazio(rsProdutos("Observacoes"))

End Sub


'Busca Grupo Pelo Codigo
Private Sub txtCodGrupo_LostFocus()

    If txtCodGrupo = "" Then
        txtNomeGrupo = ""
    ElseIf txtCodGrupo <> "" And IsNumeric(txtCodGrupo) Then
    
        rsGrupo.Open "SELECT * FROM Grupos WHERE Codigo = " & txtCodGrupo, ConexaoBD, adOpenForwardOnly, adLockOptimistic
        
        If rsGrupo.EOF <> True Then
            txtNomeGrupo = rsGrupo("Nome")
        Else
            txtNomeGrupo = ""
        End If
        
        rsGrupo.Close
    End If

End Sub

'Busca Marca Pelo Codigo
Private Sub txtCodMarca_LostFocus()

    If txtCodMarca = "" Then
        txtNomeMarca = ""
    ElseIf txtCodMarca <> "" And IsNumeric(txtCodMarca) Then
        
        rsMarca.Open "SELECT * FROM Marcas WHERE Codigo = " & txtCodMarca, ConexaoBD, adOpenForwardOnly, adLockOptimistic
        
        If rsMarca.EOF <> True Then
            txtNomeMarca = rsMarca("Nome")
        Else
            txtNomeMarca = ""
        End If
        
        rsMarca.Close
    End If

End Sub

'FORMATA��O DE VALOR
Private Sub txtPrecoEntrada_Change()
    Dim valor As String
    Dim pos As Integer
    
    ' Apenas aplique a formata��o se o campo n�o estiver vazio
    If txtPrecoEntrada <> "" Then
    
        ' Remove formata��o pr�via
        valor = Replace(txtPrecoEntrada, ",", "")
        valor = Replace(valor, ".", "")
    
        If IsNumeric(valor) Then
            ' Converte o valor para duas casas decimais
            valor = Format(Val(valor) / 100, "#,##0.00")
            
            ' Preserva a posi��o do cursor
            pos = Len(txtPrecoEntrada) - txtPrecoEntrada.SelStart
            txtPrecoEntrada = valor
            txtPrecoEntrada.SelStart = Len(txtPrecoEntrada) - pos
        ElseIf valor <> "" Then
            ' Caso digite um valor n�o num�rico
            MsgBox "Digite apenas n�meros.", vbExclamation
            txtPrecoEntrada = "0,00"
        End If
        
    End If
End Sub

'FORMATA��O DE VALOR
Private Sub txtPrecoSaida_Change()
    Dim valor As String
    Dim pos As Integer
    
    If txtPrecoSaida <> "" Then
    
        ' Remove formata��o pr�via
        valor = Replace(txtPrecoSaida, ",", "")
        valor = Replace(valor, ".", "")
    
        If IsNumeric(valor) Then
            ' Converte o valor para duas casas decimais
            valor = Format(Val(valor) / 100, "#,##0.00")
            
            ' Preserva a posi��o do cursor
            pos = Len(txtPrecoSaida) - txtPrecoSaida.SelStart
            txtPrecoSaida = valor
            txtPrecoSaida.SelStart = Len(txtPrecoSaida) - pos
        ElseIf valor <> "" Then
            ' Caso digite um valor n�o num�rico
            MsgBox "Digite apenas n�meros.", vbExclamation
            txtPrecoSaida = "0,00"
        End If
        
    End If
End Sub

'Pesquisar Marca
Private Sub btnPesquisarMarca_Click()

    frmPesquisar.TabelaBD = "Marcas"
    frmPesquisar.ColunaBD = "Nome"
    frmPesquisar.Form = "CadastroProdutos"
    frmPesquisar.PreencherCampo = "Marca"
    
    frmPesquisar.Show

End Sub

'Pesquisar Grupo
Private Sub btnPesquisarGrupo_Click()

    frmPesquisar.TabelaBD = "Grupos"
    frmPesquisar.ColunaBD = "Nome"
    frmPesquisar.Form = "CadastroProdutos"
    frmPesquisar.PreencherCampo = "Grupo"
    
    frmPesquisar.Show

End Sub

