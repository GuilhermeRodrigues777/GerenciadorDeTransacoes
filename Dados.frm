VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
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
   ScaleHeight     =   6120
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "VOLTAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3893
      MaskColor       =   &H00808080&
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       Form_Load
' Description:       Realiza ações ao iniciar a tela de consulta
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub Form_Load()

    Dim rs                  As ADODB.Recordset
    Dim conexao             As ADODB.Connection
    Dim sSQL                As String
    
    Set conexao = New ADODB.Connection
    conexao.Open "Driver={MySQL ODBC 8.0 ANSI Driver};" & _
                 "Server=localhost;" & _
                 "Port=3306;" & _
                 "Database=testeVB6;" & _
                 "User=root;" & _
                 "Password=TesteVB6;" & _
                 "Option=3;"
                 
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
     
    sSQL = "SELECT " & _
            "ID_Transacao, " & _
            "Numero_Cartao, " & _
            "Valor_Transacao, " & _
            "DATE_FORMAT(Data_Transacao, '%d/%m/%Y') AS DataFormatada, " & _
            "Descricao " & _
           "FROM transacoes WHERE " & xCondicionaisConsultar
    
    rs.Open sSQL, conexao, adOpenStatic, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "Nenhum registro encontrado para os filtros informados.", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
        
    End If
    
    Set DataGrid1.DataSource = rs

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
