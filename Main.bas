Attribute VB_Name = "modMain"
Option Explicit

Public conexao As New ADODB.Connection

Sub Main()

    If Not conectarDB Then
        MsgBox "Sem conex�o com o banco de dados.", vbExclamation
        End
    End If
    
    Principal.Show
    
End Sub

Public Function conectarDB() As Boolean

    On Error GoTo falhaConexao
    
    ' Driver   = Vers�o do ODBC instalado na m�quina
    ' Server   = Nome do server utilizado para iniciar o banco
    ' Port     = Porta utilizada para iniciar o banco
    ' Database = Nome do banco de dados
    ' User     = Usu�rio de acesso ao banco
    ' Password = Senha de acesso ao banco
    ' Option   = Se refere � configura��o do tipo de cursor (Manter valor)
    conexao.Open "Driver={MySQL ODBC 8.0 ANSI Driver};" & _
                 "Server=localhost;" & _
                 "Port=3306;" & _
                 "Database=testeVB6;" & _
                 "User=root;" & _
                 "Password=TesteVB6;" & _
                 "Option=3;"
                 
    conectarDB = True
    Exit Function
                 
falhaConexao:
    conectarDB = False
                 
End Function
