Attribute VB_Name = "Modulo"
Public cnnBD As New Connection
Public StringConexao As String

Sub Main()

    'Conexão com Banco de Dados
    StringConexao = "Provider=SQLOLEDB;Data Source=ISAAC-PC\SQLEXPRESS,1433;Initial Catalog=Administra_VB;Password=Lrsiazevedo2023@;User ID=sa;"
    
    cnnBD.ConnectionString = StringConexao
    cnnBD.Open
    
    frmMDIPrincipal.Show
    
End Sub
