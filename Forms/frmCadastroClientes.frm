VERSION 5.00
Begin VB.Form frmCadastroClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9105
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
   ScaleHeight     =   2730
   ScaleWidth      =   9105
   Begin VB.CommandButton btnAlterar 
      Caption         =   "X"
      Height          =   375
      Left            =   8520
      TabIndex        =   29
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnGravar 
      Caption         =   "G"
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnPesquisarCliente 
      Caption         =   "B"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnFinal 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnAvançar 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnAnterior 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnInicio 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnNovo 
      Caption         =   "N"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtEstado 
      Height          =   360
      Left            =   7680
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      Height          =   360
      Left            =   5880
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtNascimento 
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox cbmPessoa 
      Height          =   360
      ItemData        =   "frmCadastroClientes.frx":0000
      Left            =   7800
      List            =   "frmCadastroClientes.frx":000A
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtCodCliente 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtNomeCliente 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   6495
   End
   Begin VB.TextBox txtCPF_CNPJ 
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtTelefone 
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtCEP 
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtEndereco 
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox txtCidade 
      Height          =   360
      Left            =   5640
      TabIndex        =   9
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblEstado 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   7680
      TabIndex        =   21
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNascimento 
      Caption         =   "Nascimento:"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblPessoa 
      Caption         =   "Pessoa:"
      Height          =   255
      Left            =   7800
      TabIndex        =   18
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblCodigoCliente 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblNomeCliente 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblCPF_CNPJ 
      Caption         =   "CPF/CNPJ:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblTelefone 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblCEP 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblEndereço 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblCidade 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
End
Attribute VB_Name = "frmCadastroClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClientes As New Recordset
Dim rsClientesAberto As Boolean
Dim sql As String
Dim inclusao As Boolean

Private Sub Form_Load()
    ' Centraliza o formulário MDI Child dentro do MDI Form
    Me.Left = (frmMDIPrincipal.ScaleWidth - Me.Width) \ 2
    Me.Top = (frmMDIPrincipal.ScaleHeight - Me.Height) \ 2
    
    inclusao = True
    
    'CHAMAR FUNÇÃO DE INCLUSAO
    incluir
    
End Sub

Private Sub btnAlterar_Click()
    
    inclusao = False
    
    btnPesquisarCliente.Enabled = True
    btnNovo.Enabled = True
    btnAvançar.Enabled = True
    btnFinal.Enabled = True
    btnAnterior.Enabled = True
    btnInicio.Enabled = True
    
    alteracao
        
End Sub

Private Sub btnNovo_Click()

    inclusao = True
    
    limparCampos
    incluir
    
End Sub

Private Sub btnInicio_Click()

    rsClientes.MoveFirst
    preencherCampos

End Sub

Private Sub btnAnterior_Click()

    If rsClientes("Codigo") <> 0 Then
        rsClientes.MovePrevious
        preencherCampos
    End If

End Sub

Private Sub btnAvançar_Click()

    rsClientes.MoveNext
    If rsClientes.EOF = True Then
        rsClientes.MovePrevious
    End If
    preencherCampos

End Sub

Private Sub btnFinal_Click()

    rsClientes.MoveLast
    preencherCampos

End Sub

Private Sub btnGravar_Click()
    If gravarDados = True Then
        inclusao = True
        
        limparCampos
        incluir
    End If
End Sub

Private Sub incluir()

    'DESATIVAR BOTÕES DE ALTERAÇÃO
    btnPesquisarCliente.Enabled = False
    btnNovo.Enabled = False
    btnAvançar.Enabled = False
    btnFinal.Enabled = False
    btnAnterior.Enabled = False
    btnInicio.Enabled = False

    If rsClientesAberto = True Then
        rsClientes.Close
    End If
    
    'Setar Próximo Código do Novo Cliente
    rsClientes.Open "SELECT * FROM Clientes", ConexaoBD, adOpenDynamic, adLockOptimistic
    rsClientesAberto = True
    
    If rsClientes.EOF = True Then
        txtCodCliente = 0
    Else
        rsClientes.MoveLast
        txtCodCliente = rsClientes("Codigo") + 1
    End If
        
    txtEstoque = 0

End Sub

Private Sub alteracao()

    If rsClientesAberto = True Then
        rsClientes.Close
    End If
    
    sql = "SELECT * FROM Clientes"

    rsClientes.Open sql, ConexaoBD, adOpenDynamic, adLockOptimistic
    rsClientesAberto = True
    
    If rsClientes.EOF = True Then
        incluir
        Exit Sub
    End If
    rsClientes.MoveFirst
    preencherCampos
    
End Sub

'Gravar RecordSet
Private Function gravarDados() As Boolean
On Error GoTo Trataerro

    If rsClientesAberto = True Then
        rsClientes.Close
    End If
    
    If inclusao = True Then
        rsClientes.Open "SELECT * FROM Clientes", ConexaoBD, adOpenDynamic, adLockOptimistic
        rsClientesAberto = True
        rsClientes.AddNew
    Else
        rsClientes.Open "SELECT * FROM Clientes WHERE Codigo = " & txtCodCliente, ConexaoBD, adOpenDynamic, adLockOptimistic
    End If
    
    rsClientes("Codigo") = VazioToNull(txtCodCliente)
    rsClientes("Nome") = VazioToNull(txtNomeCliente)
    rsClientes("Pessoa") = VazioToNull(cbmPessoa)
    rsClientes("CPF_CNPJ") = VazioToNull(txtCPF_CNPJ)
    rsClientes("Telefone") = VazioToNull(txtTelefone)
    rsClientes("DataDeNascimento") = VazioToNull(txtNascimento)
    rsClientes("Email") = VazioToNull(txtEmail)
    rsClientes("CEP") = VazioToNull(txtCEP)
    rsClientes("Endereco") = VazioToNull(txtEndereco)
    rsClientes("Cidade") = VazioToNull(txtCidade)
    rsClientes("Estado") = VazioToNull(txtEstado)
    
    rsClientes.Update
    
    gravarDados = True
    
    Exit Function
    
Trataerro:
    MsgBox "Erro nos dados informados", vbInformation, "ERRO"
    rsClientes.CancelUpdate
    gravarDados = False
End Function

'Limpar dados dos txtBox
Public Sub limparCampos()
    txtNomeCliente = ""
    cbmPessoa = ""
    txtCPF_CNPJ = ""
    txtTelefone = ""
    txtNascimento = ""
    txtEmail = ""
    txtCEP = ""
    txtEndereco = ""
    txtCidade = ""
    txtEstado = ""
End Sub

Private Sub preencherCampos()

    txtCodCliente = NullToVazio(rsClientes("Codigo"))
    txtNomeCliente = NullToVazio(rsClientes("Nome"))
    cbmPessoa = NullToVazio(rsClientes("Pessoa"))
    txtCPF_CNPJ = NullToVazio(rsClientes("CPF_CNPJ"))
    txtTelefone = NullToVazio(rsClientes("Telefone"))
    
    'FORMATAÇÃO DE DATA
    If NullToVazio(rsClientes("DataDeNascimento")) <> "" Then
        txtNascimento = Format(CDate(rsClientes("DataDeNascimento")), "dd/mm/yyyy")
    Else
        txtNascimento = NullToVazio(rsClientes("DataDeNascimento"))
    End If
    
    txtEmail = NullToVazio(rsClientes("Email"))
    txtCEP = NullToVazio(rsClientes("CEP"))
    txtEndereco = NullToVazio(rsClientes("Endereco"))
    txtCidade = NullToVazio(rsClientes("Cidade"))
    txtEstado = NullToVazio(rsClientes("Estado"))

End Sub

Private Sub btnPesquisarCliente_Click()
 
    frmPesquisar.TabelaBD = "Clientes"
    frmPesquisar.ColunaBD = "Nome"
    frmPesquisar.Form = "CadastroClientes"
    
    frmPesquisar.Show
 
End Sub


