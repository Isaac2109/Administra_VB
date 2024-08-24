Attribute VB_Name = "Modulo"
Public ConexaoBD As New Connection
Public StringConexao As String

Sub Main()

    If ConectaBD = True Then
        frmLogin.Show
    End If
    
End Sub

Function ConectaBD() As Boolean
On Error GoTo Trataerro

    'Conexão com Banco de Dados
    StringConexao = "Provider=SQLOLEDB;Data Source=ISAAC-PC\SQLEXPRESS,1433;Initial Catalog=Administra_VB;Password=Lrsiazevedo2023@;User ID=sa;"
    
    ConexaoBD.ConnectionString = StringConexao
    ConexaoBD.Open
    
    ' Verificar o estado da conexão
    If ConexaoBD.State = 1 Then
        ' Se a conexão estiver aberta, retornar True
        ConectaBD = True
    End If
    
    Exit Function
    
Trataerro:
    MsgBox "Conexão Com Banco De Dados não foi realizada corretamente!     Por Favor Verifique seu Banco De Dados", vbOKOnly, "Erro com Banco De Dados"
    End
    
End Function
