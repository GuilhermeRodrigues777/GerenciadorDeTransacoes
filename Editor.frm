VERSION 5.00
Begin VB.Form Editor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesc_Editar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2123
      TabIndex        =   11
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtData_Editar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2123
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txtValor_Editar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2123
      TabIndex        =   7
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtNumCartao_Editar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2123
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtID_Editar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2123
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "VOLTAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3533
      TabIndex        =   1
      Top             =   5160
      Width           =   1995
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "SALVAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1133
      TabIndex        =   0
      Top             =   5160
      Width           =   1995
   End
   Begin VB.Label lblSucessoDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4725
   End
   Begin VB.Label lblDadosDa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DADOS DA TRANSAÇÃO QUE DESEJA ALTERAR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   465
      TabIndex        =   12
      Top             =   360
      Width           =   5805
   End
   Begin VB.Label lblDescriçãoDa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição da transação"
      Height          =   285
      Left            =   2130
      TabIndex        =   10
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label lblDataDa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data da transação"
      Height          =   285
      Left            =   2130
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblValorDa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor da transação"
      Height          =   285
      Left            =   2130
      TabIndex        =   6
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblNumCartao_Editar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número do cartão"
      Height          =   285
      Left            =   2130
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblIDTransação_Editar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID da transação"
      Height          =   285
      Left            =   2123
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       Form_Load
' Description:       Realiza ações ao iniciar a tela de edição
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub Form_Load()

    Dim comando         As New ADODB.Command
    
    Set rs = New ADODB.Recordset
    
    ' Liga conexão com o bd
    comando.ActiveConnection = conexao
    
    ' Consulta quais são os dados retornados com o ID inserido
    comando.CommandText = "SELECT * FROM transacoes WHERE " & xCondicionaisEditar
    Set rs = comando.Execute
    
    ' Se não der EndOfFile significa que encontrou dados
    ' Mostra na tela quais são os dados retornados
    If Not rs.EOF Then
        txtID_Editar.Text = rs("ID_Transacao")
        txtNumCartao_Editar.Text = rs("Numero_Cartao")
        txtValor_Editar.Text = rs("Valor_Transacao")
        txtData_Editar.Text = Format(rs("Data_Transacao"), "dd/mm/yyyy")
        txtDesc_Editar.Text = rs("Descricao")
    
    Else
        MsgBox "Não foi encontrado registro os dados inseridos", vbExclamation
        Exit Sub
        
    End If
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       cmdSalvar_Click
' Description:       Faz o botão 'cmdSalvar' salvar os dados no bd
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub cmdSalvar_Click()

    Dim comando     As New ADODB.Command
    
    If txtID_Editar <> xIDDigitado Then
        MsgBox "Campo ID não pode ser alterado", vbExclamation
        txtID_Editar = xIDDigitado
        txtID_Editar.SetFocus
        Exit Sub
    
    End If
    
    If txtID_Editar.Text = "" Or _
    txtNumCartao_Editar.Text = "" Or _
    txtValor_Editar.Text = "" Or _
    txtData_Editar.Text = "" Or _
    txtDesc_Editar.Text = "" Then
        MsgBox "Todos os campos devem ser preenchidos", vbExclamation
        Exit Sub
        
    End If
    
    If txtNumCartao_Editar.Text <> "" Then
        If VerificarNumCartao(txtNumCartao_Editar) Then Exit Sub
        
    End If
    
    If txtValor_Editar.Text <> "" Then
        If VerificarValor(txtValor_Editar) Then Exit Sub
    
    End If
    
    If txtData_Editar.Text <> "" Then
        If VerificarData(txtData_Editar) Then Exit Sub
    
    End If
    
    ' Liga conexão com o bd
    comando.ActiveConnection = conexao
    
    ' Faz a atualização dos dados no bd
    comando.CommandText = "UPDATE transacoes SET ID_Transacao = ?, Numero_Cartao = ?, Valor_Transacao = ?, Data_Transacao = ?, Descricao = ? WHERE ID_Transacao = ?"
    
    With comando.Parameters
        .Append comando.CreateParameter("ID_Transacao", adInteger, adParamInput, 10, txtID_Editar.Text)
        .Append comando.CreateParameter("Numero_Cartao", adInteger, adParamInput, 19, txtNumCartao_Editar.Text)
        .Append comando.CreateParameter("Valor_Transacao", adVarChar, adParamInput, 30, txtValor_Editar.Text)
        .Append comando.CreateParameter("Data_Transacao", adDate, adParamInput, 10, (Format$(CDate(txtData_Editar.Text), "dd-mm-yyyy")))
        .Append comando.CreateParameter("Descricao", adVarChar, adParamInput, 30, txtDesc_Editar.Text)
        .Append comando.CreateParameter("ID_Transacao_Where", adInteger, adParamInput, 10, txtID_Editar.Text)
    End With
        
    comando.Execute
    
    Unload Me
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       cmdVoltar_Click
' Description:       Faz o botão 'cmdVoltar' fechar a janela de edição
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub cmdVoltar_Click()
    
    Unload Me
    
End Sub
