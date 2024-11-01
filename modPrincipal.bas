Attribute VB_Name = "modPrincipal"
Option Explicit

Public conexao As New ADODB.Connection
Public usuarioLogado As Long


Sub Main()
    If Not conectarDB() Then
        MsgBox "Não foi possível conectar ao banco de dados! Verifique a configuração e tente novamente.", vbCritical
        End
    Else
        MsgBox "Conexão com banco de dados estabelecida com sucesso!"
    End If
    FrmLogin.Show
End Sub

Private Function conectarDB() As Boolean
    On Error GoTo erro_conecta

    conexao.Open "Driver={MySQL ODBC 8.0 ANSI Driver};" & _
                 "Server=192.168.1.200;" & _
                 "Port=3306;" & _
                 "Database=vb6_pdv;" & _
                 "User=marcos;" & _
                 "Password=1234;" & _
                 "Option=3;"
    conectarDB = True
    Exit Function

erro_conecta:
    conectarDB = False
    MsgBox "Erro ao conectar ao banco de dados: " & Err.Description, vbCritical
End Function

Sub FecharConexao()
    If conexao.State = adStateOpen Then
        conexao.Close
        Set conexao = Nothing
    End If
End Sub

