VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPesquisar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6870
   Begin MSAdodcLib.Adodc adoPesquisa 
      Height          =   375
      Left            =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtgridPesquisa 
      Bindings        =   "frmPesquisar.frx":0000
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   6075,213
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnPesquisar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtPesquisa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmPesquisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TabelaBD As String
Public ColunaBD As String
Public Form As String
Public PreencherCampo As String

Private Sub Form_Load()

    ' Centraliza o formulário MDI Child dentro do MDI Form
    Me.Left = (frmMDIPrincipal.ScaleWidth - Me.Width) \ 2
    Me.Top = (frmMDIPrincipal.ScaleHeight - Me.Height) \ 2

End Sub

Private Sub btnPesquisar_Click()

    Dim sql As String
    
    sql = "SELECT * " & _
            "FROM " & TabelaBD & _
            " WHERE " & ColunaBD & " LIKE '%" & txtPesquisa & "%'"

    With adoPesquisa
        .UserName = "sa"
        .Password = "Lrsiazevedo2023@"
        .ConnectionString = ConexaoBD
        .RecordSource = sql
        .Refresh
    End With

End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        btnPesquisar.SetFocus
    End If

End Sub

Private Sub dtgridPesquisa_DblClick()

    If Form = "CadastroProdutos" Then
    
        If TabelaBD = "Grupos" Then
        
            FrmCadastroProdutos.txtCodGrupo = adoPesquisa.Recordset("Codigo")
            FrmCadastroProdutos.txtNomeGrupo = adoPesquisa.Recordset("Nome")
            
        ElseIf TabelaBD = "Marcas" Then
        
            FrmCadastroProdutos.txtCodMarca = adoPesquisa.Recordset("Codigo")
            FrmCadastroProdutos.txtNomeMarca = adoPesquisa.Recordset("Codigo")
            
        ElseIf TabelaBD = "Produtos" Then
        
            FrmCadastroProdutos.limparCampos
            FrmCadastroProdutos.txtCodProduto = NullToVazio(adoPesquisa.Recordset("Codigo"))
            FrmCadastroProdutos.txtNomeProduto = NullToVazio(adoPesquisa.Recordset("Nome"))
            FrmCadastroProdutos.cbmSituacao = NullToVazio(adoPesquisa.Recordset("Situacao"))
            
            FrmCadastroProdutos.txtCodMarca = NullToVazio(adoPesquisa.Recordset("CodigoMarca"))
            FrmCadastroProdutos.txtCodMarca_LostFocus
            
            FrmCadastroProdutos.txtCodGrupo = NullToVazio(adoPesquisa.Recordset("CodigoGrupo"))
            FrmCadastroProdutos.txtCodGrupo_LostFocus
            
            ' Formata o valor do banco para exibir com 2 casas decimais
            FrmCadastroProdutos.txtPrecoEntrada = Format(adoPesquisa.Recordset("PrecoEntrada"), "#,##0.00")
            FrmCadastroProdutos.txtPrecoSaida = Format(adoPesquisa.Recordset("PrecoSaida"), "#,##0.00")
            
            FrmCadastroProdutos.txtEstoque = NullToVazio(adoPesquisa.Recordset("Estoque"))
            FrmCadastroProdutos.txtObservacoes = NullToVazio(adoPesquisa.Recordset("Observacoes"))
            
        End If
    
    ElseIf Form = "CadastroClientes" Then
    
        frmCadastroClientes.limparCampos
        frmCadastroClientes.txtCodCliente = NullToVazio(adoPesquisa.Recordset("Codigo"))
        frmCadastroClientes.txtNomeCliente = NullToVazio(adoPesquisa.Recordset("Nome"))
        frmCadastroClientes.cbmPessoa = NullToVazio(adoPesquisa.Recordset("Pessoa"))
        frmCadastroClientes.txtCPF_CNPJ = NullToVazio(adoPesquisa.Recordset("CPF_CNPJ"))
        frmCadastroClientes.txtTelefone = NullToVazio(adoPesquisa.Recordset("Telefone"))
        frmCadastroClientes.txtNascimento = NullToVazio(adoPesquisa.Recordset("DataDeNascimento"))
        frmCadastroClientes.txtEmail = NullToVazio(adoPesquisa.Recordset("Email"))
        frmCadastroClientes.txtCEP = NullToVazio(adoPesquisa.Recordset("CEP"))
        frmCadastroClientes.txtEndereco = NullToVazio(adoPesquisa.Recordset("Endereco"))
        frmCadastroClientes.txtCidade = NullToVazio(adoPesquisa.Recordset("Cidade"))
        frmCadastroClientes.txtEstado = NullToVazio(adoPesquisa.Recordset("Estado"))
        
    End If
    
    Unload Me

End Sub





