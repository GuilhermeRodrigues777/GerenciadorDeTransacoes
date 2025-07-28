VERSION 5.00
Begin VB.Form Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Transações com Cartão"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerSucessoEdit 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   9075
      Top             =   4920
   End
   Begin VB.Timer timerSucessoInsert 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   10275
      Top             =   4920
   End
   Begin VB.Timer timerSucessoDelete 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   9675
      Top             =   4920
   End
   Begin VB.CommandButton btConsultar 
      Caption         =   "CONSULTAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   8400
      TabIndex        =   9
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton btExcluir 
      Caption         =   "EXCLUIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   8400
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton btEditar 
      Caption         =   "EDITAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   8400
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton btInserir 
      Caption         =   "INSERIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   8400
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   7455
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2985
         MaxLength       =   10
         TabIndex        =   1
         Top             =   390
         Width           =   4000
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2985
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2880
         Width           =   4000
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2985
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2256
         Width           =   4000
      End
      Begin VB.TextBox txtValor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2985
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1634
         Width           =   4000
      End
      Begin VB.TextBox txtNumCartao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2985
         MaxLength       =   19
         TabIndex        =   2
         Top             =   1012
         Width           =   4000
      End
      Begin VB.Label lblLabel1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID Transação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   15
         Top             =   390
         Width           =   2500
      End
      Begin VB.Label lblLabel2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número do cartão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   14
         Top             =   1012
         Width           =   2500
      End
      Begin VB.Label lblLabel5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   13
         Top             =   2880
         Width           =   2500
      End
      Begin VB.Label lblLabel4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data da transação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   12
         Top             =   2256
         Width           =   2500
      End
      Begin VB.Label lblLabel3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da transação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   11
         Top             =   1634
         Width           =   2500
      End
   End
   Begin VB.Label lblSucessoEdit 
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
      Left            =   3240
      TabIndex        =   19
      Top             =   4920
      Width           =   4725
   End
   Begin VB.Label lblSucessoDelete 
      Alignment       =   2  'Center
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
      Left            =   3203
      TabIndex        =   18
      Top             =   4920
      Width           =   4725
   End
   Begin VB.Label Label1 
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
      TabIndex        =   17
      Top             =   0
      Width           =   4725
   End
   Begin VB.Label lblSucessoInsert 
      Alignment       =   2  'Center
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
      Left            =   3203
      TabIndex        =   16
      Top             =   4920
      Width           =   4725
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000011&
      BackStyle       =   0  'Transparent
      Caption         =   "GERENCIADOR DE TRANSAÇÕES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2618
      TabIndex        =   10
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       btInserir_Click
' Description:       Verifica requisitos necessários e realiza inserção dos dados no bd
' Created by :       Guilherme Rodrigues
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub btInserir_Click(Index As Integer)

    Dim comando             As New ADODB.Command
    Dim sIDRetornado        As Boolean
    
    If txtID.Text = "" Or _
    txtNumCartao.Text = "" Or _
    txtValor.Text = "" Or _
    txtData.Text = "" Or _
    txtDesc.Text = "" Then
        MsgBox "Todos os campos devem ser preenchidos", vbExclamation
        Exit Sub
        
    End If
    
    ' Valida se valores de entrada são vazios
    If txtID.Text = "" Then
        MsgBox "ID da transação não informado", vbExclamation
        txtID.SetFocus
        Exit Sub
        
    End If
    
    If txtNumCartao.Text = "" Then
        MsgBox "Número do cartão não informado", vbExclamation
        txtNumCartao.SetFocus
        Exit Sub
        
    End If
    
    If txtValor.Text = "" Then
        MsgBox "Valor da transação não informado", vbExclamation
        txtValor.SetFocus
        Exit Sub
        
    End If
    
    If txtData.Text = "" Then
        MsgBox "Data da transação não informada", vbExclamation
        txtData.SetFocus
        Exit Sub
        
    End If
    
    If txtDesc.Text = "" Then
        MsgBox "Descrição da transação não informada", vbExclamation
        txtDesc.SetFocus
        Exit Sub
        
    End If
    
    ' Valida se valores de entrada são válidos
    If VerificarID("Inserir", txtID) Then Exit Sub
    
    If VerificarNumCartao(txtNumCartao) Then Exit Sub
    
    If VerificarValor(txtValor) Then Exit Sub
    
    If VerificarData(txtData) Then Exit Sub
    
    ' Liga conexão com o bd
    comando.ActiveConnection = conexao
    
    ' Insere os dados no bd
    comando.CommandText = "INSERT INTO transacoes (ID_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao) VALUES (?, ?, ?, ?, ?)"
    
    With comando.Parameters
        .Append comando.CreateParameter("ID_Transacao", adInteger, adParamInput, 10, txtID.Text)
        .Append comando.CreateParameter("Numero_Cartao", adVarChar, adParamInput, 19, txtNumCartao.Text)
        .Append comando.CreateParameter("Valor_Transacao", adVarChar, adParamInput, 30, txtValor.Text)
        .Append comando.CreateParameter("Data_Transacao", adDate, adParamInput, 10, (Format$(CDate(txtData.Text), "dd-mm-yyyy")))
        .Append comando.CreateParameter("Descricao", adVarChar, adParamInput, 30, txtDesc.Text)
    End With
        
    comando.Execute
    
    lblSucessoInsert.Caption = "Registro inserido com sucesso."
    txtID.Text = ""
    txtNumCartao.Text = ""
    txtValor.Text = ""
    txtData.Text = ""
    txtDesc.Text = ""
    lblSucessoInsert.ForeColor = RGB(0, 100, 0)
    lblSucessoInsert.Visible = True
    timerSucessoInsert.Enabled = True
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       btExcluir_Click
' Description:       Verifica requisitos necessários e realiza exclusão dos dados no bd
' Created by :       Guilherme Rodrigues
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub btExcluir_Click(Index As Integer)
    
    Dim comando             As New ADODB.Command
    Dim rs                  As ADODB.Recordset
    Dim sResposta           As VbMsgBoxResult
    Dim sID                 As Integer
    Dim sNumCartao          As Long
    Dim sValor              As Double
    Dim sValorComPonto      As Double
    Dim sData               As Date
    Dim sDesc               As String
    Dim sCondicionais       As String
    
    Set rs = New ADODB.Recordset
    
    sID = Empty
    sNumCartao = Empty
    sValor = Empty
    sData = "1/1/0001"
    sDesc = Empty
    sCondicionais = ""
    
    If txtID.Text = "" Then
        MsgBox "ID da transação precisa ser informado", vbExclamation
        txtID.SetFocus
        Exit Sub
    End If
    
    ' Liga conexão com o bd
    comando.ActiveConnection = conexao
    
    If txtID.Text <> "" Then
        If VerificarID("Excluir", txtID) Then Exit Sub
        sID = txtID.Text
        
    End If
    
    If txtNumCartao.Text <> "" Then
        If VerificarNumCartao(txtNumCartao) Then Exit Sub
            
        ' Verifica se já existe alguma transação com o número de cartão informado
        comando.CommandText = "SELECT * FROM transacoes WHERE Numero_Cartao = '" & txtNumCartao.Text & "'"
        Set rs = comando.Execute
        
        If Not rs.EOF Then ' Se não der EndOfFile significa que encontrou dados
            sNumCartao = txtNumCartao.Text
        
        Else
            MsgBox "Não foi encontrado registro com esse número de cartão", vbExclamation
            Exit Sub
            
        End If
        
        rs.Close
        
    End If
    
    If txtValor.Text <> "" Then
        If VerificarValor(txtValor) Then Exit Sub
        
        ' Verifica se já existe alguma transação com o valor informado
        comando.CommandText = "SELECT * FROM transacoes WHERE Valor_Transacao = '" & txtValor.Text & "'"
        Set rs = comando.Execute
        
        If Not rs.EOF Then ' Se não der EndOfFile significa que encontrou dados
            sValor = txtValor.Text
        
        Else
            MsgBox "Não foi encontrado registro com esse valor", vbExclamation
            Exit Sub
            
        End If
        
        rs.Close
        
    End If
    
    If txtData.Text <> "" Then
        If VerificarData(txtData) Then Exit Sub
        
        ' Verifica se já existe alguma transação com o valor informado
        comando.CommandText = "SELECT * FROM transacoes WHERE Data = '" & txtData.Text & "'"
        Set rs = comando.Execute
        
        If Not rs.EOF Then ' Se não der EndOfFile significa que encontrou dados
            sData = txtData.Text
        
        Else
            MsgBox "Não foi encontrado registro com esse valor", vbExclamation
            Exit Sub
            
        End If
        
    End If
    
    If txtDesc.Text <> "" Then
        
        ' Verifica se já existe alguma transação com o valor informado
        comando.CommandText = "SELECT * FROM transacoes WHERE Descricao = '" & txtDesc.Text & "'"
        Set rs = comando.Execute
        
        If Not rs.EOF Then ' Se não der EndOfFile significa que encontrou dados
            sDesc = txtDesc.Text
        
        Else
            MsgBox "Não foi encontrado registro com essa descrição", vbExclamation
            Exit Sub
            
        End If
        
    End If
    
    If IsNull(sID) And IsNull(sNumCartao) And _
    IsNull(sValor) And (IsNull(sData) Or sData = "1/1/0001") And _
    IsNull(sDesc) Then
        MsgBox "Nenhum campo preenchido", vbExclamation
        Exit Sub
        
    End If
    
    ' Pede uma confirmação da decisão de excluir os dados
    sResposta = MsgBox("Realmente deseja excluir os dados dessa transação?", vbYesNo + vbQuestion, "")
    
    If sResposta = vbYes Then
        
        If sID <> 0 Then
            sCondicionais = sCondicionais & "ID_Transacao = " & sID & " AND "
            
        End If
        
        If sNumCartao <> 0 Then
            sCondicionais = sCondicionais & "Numero_Cartao = '" & sNumCartao & "' AND "
            
        End If
        
        If sValor <> 0 Then
            sCondicionais = sCondicionais & "Valor_Transacao = " & sValor & " AND "
            
        End If
        
        If sData <> "1/1/0001" Then
            sCondicionais = sCondicionais & "Data_Transacao = #" & sData & "# AND "
            
        End If
        
        If sDesc <> "" Then
            sCondicionais = sCondicionais & "Descricao = '" & sDesc & "' AND "
            
        End If
        
        ' Remove o último "AND"
        If Right(sCondicionais, 5) = " AND " Then
            sCondicionais = Left(sCondicionais, Len(sCondicionais) - 5)
            
        End If
        
        If sCondicionais <> "" Then
            comando.CommandText = "DELETE FROM transacoes WHERE " & sCondicionais
            
        Else
            MsgBox "Nenhuma condicional informada para deletar registros", vbExclamation
            Exit Sub
            
        End If
        
        comando.Execute
        
        lblSucessoDelete.Caption = "Registro deletado com sucesso."
        txtID.Text = ""
        txtNumCartao.Text = ""
        txtValor.Text = ""
        txtData.Text = ""
        txtDesc.Text = ""
        lblSucessoDelete.ForeColor = RGB(0, 100, 0)
        lblSucessoDelete.Visible = True
        timerSucessoDelete.Enabled = True
        
    Else
        Exit Sub
        
    End If
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       btEditar_Click
' Description:       Verifica requisitos necessários e realiza edição dos dados no bd
' Created by :       Guilherme Rodrigues
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub btEditar_Click(Index As Integer)
    
    If txtID.Text = "" Then
        MsgBox "Pelo menos o campo ID precisa ser informado", vbExclamation
        Exit Sub
        
    End If
    
    If txtID.Text <> "" Then
        If VerificarID("Editar", txtID) Then Exit Sub
        sCondicionais = sCondicionais & "ID_Transacao = '" & txtID.Text & "' AND "

    End If
    
    If txtNumCartao.Text <> "" Then
        If VerificarNumCartao(txtNumCartao) Then Exit Sub
        sCondicionais = sCondicionais & "Numero_Cartao = '" & txtNumCartao.Text & "' AND "
        
    End If
    
    If txtValor.Text <> "" Then
        If VerificarValor(txtValor) Then Exit Sub
        sCondicionais = sCondicionais & "Valor_Transacao = '" & txtValor.Text & "' AND "
        
    End If
    
    If txtData.Text <> "" Then
        If VerificarData(txtData) Then Exit Sub
        sCondicionais = sCondicionais & "Data_Transacao = '" & Format(txtData.Text, "yyyy-mm-dd") & "' AND "
        
    End If
    
    If txtDesc.Text <> "" Then
        sCondicionais = sCondicionais & "Descricao = '" & txtDesc.Text & "' AND "
    
    End If
    
    ' Remove o último "AND"
    If Right(sCondicionais, 5) = " AND " Then
        sCondicionais = Left(sCondicionais, Len(sCondicionais) - 5)
        
    End If
    
    xIDDigitado = txtID.Text
    xCondicionaisEditar = sCondicionais
    
    Editor.Show
    
    lblSucessoEdit.Caption = "Registro alterado com sucesso."
    txtID.Text = ""
    txtNumCartao.Text = ""
    txtValor.Text = ""
    txtData.Text = ""
    txtDesc.Text = ""
    lblSucessoEdit.ForeColor = RGB(0, 100, 0)
    lblSucessoEdit.Visible = True
    timerSucessoEdit.Enabled = True

End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       btConsultar_Click
' Description:       Verifica requisitos necessários e realiza consulta dos dados no bd
' Created by :       Guilherme Rodrigues
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub btConsultar_Click(Index As Integer)

    Dim sCondicionais   As String
    
    sCondicionais = ""

    If txtID.Text = "" And _
    txtNumCartao.Text = "" And _
    txtValor.Text = "" And _
    txtData.Text = "" And _
    txtDesc.Text = "" Then
        MsgBox "Pelo menos 1 campo precisa ser informado", vbExclamation
        Exit Sub
        
    End If
    
    If txtID.Text <> "" Then
        If VerificarID("Consultar", txtID) Then Exit Sub
        sCondicionais = sCondicionais & "ID_Transacao = '" & txtID.Text & "' AND "
    
    End If
    
    If txtNumCartao.Text <> "" Then
        If VerificarNumCartao(txtNumCartao) Then Exit Sub
        sCondicionais = sCondicionais & "Numero_Cartao = '" & txtNumCartao.Text & "' AND "
        
    End If
    
    If txtValor.Text <> "" Then
        If VerificarValor(txtValor) Then Exit Sub
        sCondicionais = sCondicionais & "Valor_Transacao = '" & txtValor.Text & "' AND "
        
    End If
    
    If txtData.Text <> "" Then
        If VerificarData(txtData) Then Exit Sub
        sCondicionais = sCondicionais & "Data_Transacao = '" & Format(txtData.Text, "yyyy-mm-dd") & "' AND "
        
    End If
    
    If txtDesc.Text <> "" Then
        sCondicionais = sCondicionais & "Descricao = '" & txtDesc.Text & "' AND "
    
    End If
    
    ' Remove o último "AND"
    If Right(sCondicionais, 5) = " AND " Then
        sCondicionais = Left(sCondicionais, Len(sCondicionais) - 5)
        
    End If
    
    xCondicionaisConsultar = sCondicionais
    
    Consulta.Show
    
    txtID.Text = ""
    txtNumCartao.Text = ""
    txtValor.Text = ""
    txtData.Text = ""
    txtDesc.Text = ""

End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       timerSucessoInsert_Timer
' Description:       Controla propriedades da label 'lblSucessoInsert'
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub timerSucessoInsert_Timer()

    lblSucessoInsert.Visible = False
    timerSucessoInsert.Enabled = False
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       timerSucessoDelete_Timer
' Description:       Controla propriedades da label 'lblSucessoDelete'
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub timerSucessoDelete_Timer()

    lblSucessoDelete.Visible = False
    timerSucessoDelete.Enabled = False

End Sub

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       timerSucessoEdit_Timer
' Description:       Controla propriedades da label 'lblSucessoEdit'
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub timerSucessoEdit_Timer()

    lblSucessoEdit.Visible = False
    timerSucessoEdit.Enabled = False

End Sub
