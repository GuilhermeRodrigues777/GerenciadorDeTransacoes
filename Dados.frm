VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7200
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "EXPORTAR PARA EXCEL"
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
      Left            =   3180
      TabIndex        =   2
      Top             =   5160
      Width           =   3135
   End
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
      Left            =   720
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

    sSQL = "SELECT " & _
           "ID_Transacao, " & _
           "Numero_Cartao, " & _
           "Valor_Transacao, " & _
           "DATE_FORMAT(Data_Transacao, '%d/%m/%Y') AS DataFormatada, " & _
           "Descricao " & _
           "FROM transacoes WHERE " & xCondicionaisConsultar

    Set xRsGlobal = New ADODB.Recordset
    xRsGlobal.CursorLocation = adUseClient
    xRsGlobal.Open sSQL, conexao, adOpenStatic, adLockReadOnly

    If xRsGlobal.EOF Then
        MsgBox "Nenhum registro encontrado para os filtros informados", vbExclamation
        xRsGlobal.Close
        Set xRsGlobal = Nothing
        Exit Sub
        
    End If

    Set DataGrid1.DataSource = xRsGlobal
    
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

'--------------------------------------------------------------------------------
' Project    :       GestaoDeTransacoes
' Name       :       cmdExport_Click
' Description:       Faz o botão 'cmdExport' exportar os dados do DataGrid para um arquivo Excel
' Created by :       Guilherme Rodrigues
'--------------------------------------------------------------------------------
Private Sub cmdExport_Click()

    Dim exlApp          As Object
    Dim exlBook         As Object
    Dim exlSheet        As Object
    Dim i               As Integer
    Dim linha           As Integer
    Dim sCaminho        As String

    If xRsGlobal Is Nothing Or xRsGlobal.EOF Then
        MsgBox "Nenhum dado para exportar", vbExclamation
        Exit Sub
        
    End If

    CommonDialog1.CancelError = True
    On Error GoTo Fim

    CommonDialog1.DialogTitle = "Salvar planilha Excel"
    CommonDialog1.Filter = "Arquivos Excel (*.xlsx)|*.xlsx"
    CommonDialog1.DefaultExt = "xlsx"
    CommonDialog1.ShowSave

    sCaminho = CommonDialog1.FileName

    ' Starta o Excel
    Set exlApp = CreateObject("Excel.Application")
    Set exlBook = exlApp.Workbooks.Add
    Set exlSheet = exlBook.Sheets(1)

    For i = 0 To xRsGlobal.Fields.Count - 1
        exlSheet.Cells(1, i + 1).Value = xRsGlobal.Fields(i).Name
    Next i

    ' Preenche os dados no arquivo
    linha = 2
    xRsGlobal.MoveFirst
    
    Do While Not xRsGlobal.EOF
    
        For i = 0 To xRsGlobal.Fields.Count - 1
            exlSheet.Cells(linha, i + 1).Value = xRsGlobal.Fields(i).Value
        Next i
        
        linha = linha + 1
        xRsGlobal.MoveNext
        
    Loop

    exlSheet.Columns.AutoFit

    ' Salva no local escolhido
    exlBook.SaveAs sCaminho
    MsgBox "Arquivo salvo com sucesso em:" & vbCrLf & sCaminho, vbInformation

    exlBook.Close False
    exlApp.Quit
    Set exlSheet = Nothing
    Set exlBook = Nothing
    Set exlApp = Nothing
    
    Exit Sub

Fim:
    If Err.Number = 32755 Then
        MsgBox "Exportação cancelada.", vbInformation
        
    Else
        MsgBox "Erro ao salvar: " & Err.Description, vbCritical
        
    End If
    
End Sub
