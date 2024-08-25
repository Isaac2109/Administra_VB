VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAlterarSenha 
      Caption         =   "Alterar Senha"
      Height          =   435
      Left            =   7440
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton btnEntrar 
      Caption         =   "Ok"
      Height          =   435
      Left            =   5520
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox picAdministra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   4500
      TabIndex        =   7
      Top             =   120
      Width           =   4500
   End
   Begin VB.TextBox txtSenha 
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
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lblSlogan 
      Caption         =   "Administra: Potencializando Seu Negócio com Eficiência."
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
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label lblSenha 
      Caption         =   "Senha:"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuário:"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------Key Press
Private Sub txtSenha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        btnEntrar_Click
    End If

End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtSenha.SetFocus
    End If
    
End Sub

Private Sub btnCancelar_Click()

    Unload Me

End Sub

Private Sub btnEntrar_Click()

    Dim rsUsuario As New Recordset
    Dim Query As String
    Dim Login As String
    Dim Senha As String
    
    Login = txtUsuario
    Senha = txtSenha
    
    Query = "Select * from Usuarios Where Login = '" & Login & "' and Senha = '" & Senha & "'"
    
    rsUsuario.Open Query, ConexaoBD, adOpenForwardOnly, adLockOptimistic
    
    If rsUsuario.EOF = True Then
        MsgBox "Usuario ou Senha Incorreta!", vbOKOnly, "Atenção!"
        txtSenha.SetFocus
    Else
        Unload Me
        frmMDIPrincipal.Show
    End If
    
End Sub
