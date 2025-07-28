Attribute VB_Name = "modGlobal"
Option Explicit

' Declara variáveis que podem ser usadas em qualquer lugar do código
Public xIDDigitado                  As Integer
Public xCondicionaisEditar          As String
Public xCondicionaisConsultar       As String
Public xRsGlobal                    As ADODB.Recordset

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       VerificarID
' Description:       Controla condicionais para ID_Transacao
' Parameters :       sMetodo (String) - De onde vem a chamada para consultar o ID
'                    sID (TextBox)    - Qual ID deve ser utilizado para consulta
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Public Function VerificarID(ByVal sMetodo As String, ByVal sID As TextBox) As Boolean
    
    Dim comando     As New ADODB.Command
    Dim rs          As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    VerificarID = False
    
    ' Verifica se valor inserido é numérico
    If Not IsNumeric(sID.Text) Then
        VerificarID = True
        MsgBox "ID da transação é inválido", vbExclamation
        sID.Text = ""
        sID.SetFocus
        Exit Function
        
    End If
    
    ' Liga conexão com o bd
    comando.ActiveConnection = conexao
    
    ' Verifica se já existe alguma transação com o ID informado
    comando.CommandText = "SELECT * FROM transacoes WHERE ID_Transacao = '" & sID.Text & "'"
    Set rs = comando.Execute
    
    If Not rs.EOF Then ' Se não der EndOfFile significa que encontrou dados
    
        Select Case sMetodo
        
            Case "Inserir"
                VerificarID = True
                MsgBox "ID inserido já existe na base de dados", vbExclamation
                sID.Text = ""
                sID.SetFocus
                Exit Function
                
            Case "Consultar", "Editar", "Excluir"
                Exit Function

        End Select
        
    End If
    
    rs.Close
    
End Function

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       VerificarNumCartao
' Description:       Controla condicionais para Numero_Cartao
' Created by :       Guilherme Rodrigues
' Parameters :       sNumCartao (TextBox) - Controle de número do cartão inserido pelo usuário
'--------------------------------------------------------------------------------
Public Function VerificarNumCartao(ByVal sNumCartao As TextBox) As Boolean
    
    VerificarNumCartao = False
    
    If Not IsNumeric(sNumCartao) Then
        VerificarNumCartao = True
        MsgBox "Número do cartão digitado é inválido", vbExclamation
        sNumCartao.Text = ""
        sNumCartao.SetFocus
        Exit Function
        
    End If

End Function

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       VerificarValor
' Description:       Controla condicionais para Valor_Transacao
' Created by :       Guilherme Rodrigues
' Parameters :       sValor (TextBox) - Controle de valor inserido pelo usuário
'--------------------------------------------------------------------------------
Public Function VerificarValor(ByVal sValor As TextBox) As Boolean
    
    Dim sValorComVirgula    As String
    
    VerificarValor = False
    
    If Not IsNumeric(sValor.Text) Then
        VerificarValor = True
        MsgBox "Valor digitado é inválido", vbExclamation
        sValor.SetFocus
        Exit Function
        
    End If
    
    sValorComVirgula = InStr(1, sValor.Text, ".")
    
    If sValorComVirgula > 0 Then
      sValor.Text = Replace(sValor.Text, ".", ",")
      
    End If

End Function

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       VerificarData
' Description:       Controla condicionais para Data_Transacao
' Created by :       Guilherme Rodrigues
' Parameters :       sData (TextBox) - Controle de data inserida pelo usuário
'--------------------------------------------------------------------------------
Public Function VerificarData(ByVal sData As TextBox) As Boolean
    
    VerificarData = False
    
    If Len(sData.Text) = 8 And InStr(sData.Text, "/") = 0 Then
        sData.Text = Left(sData.Text, 2) & "/" & Mid(sData.Text, 3, 2) & "/" & Right(sData.Text, 4)
    
    End If
    
    ' Verifica se data é válida
    If Not IsDate(sData.Text) Then
        VerificarData = True
        MsgBox "Data inserida é inválida", vbExclamation
        Exit Function
        
    End If
    
End Function
