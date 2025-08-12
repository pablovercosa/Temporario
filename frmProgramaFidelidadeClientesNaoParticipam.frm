VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProgramaFidelidadeClientesNaoParticipam 
   Caption         =   " Programa Fidelidade x Clientes que não participam"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgramaFidelidadeClientesNaoParticipam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_cnpj 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   120
      TabIndex        =   6
      Top             =   390
      Width           =   2295
   End
   Begin VB.TextBox txt_nomeEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   2460
      TabIndex        =   5
      Top             =   390
      Width           =   6255
   End
   Begin VB.CommandButton cmd_novo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vincular Cliente (Para NÃO participar)"
      Height          =   430
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   4245
   End
   Begin VB.CommandButton cmd_excluir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Desvincular"
      Height          =   430
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   4245
   End
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   430
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5190
      Width           =   8595
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   430
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5700
      Width           =   8595
   End
   Begin VB.ComboBox cmb_clientes 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4680
      Width           =   7785
   End
   Begin MSFlexGridLib.MSFlexGrid grade_programas 
      Height          =   2805
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   12648447
      BackColorSel    =   12648384
      ForeColorSel    =   0
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Clientes que não participam do Programa de Fidelidade"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   930
      Width           =   4665
   End
   Begin VB.Label Label9 
      Caption         =   "Empresa/Filial"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   150
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Clientes"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4710
      Width           =   825
   End
End
Attribute VB_Name = "frmProgramaFidelidadeClientesNaoParticipam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s_Funcao As String

Private Sub cmd_cancelar_Click()
  grade_programas.Enabled = True
  cmd_excluir.Enabled = True
  cmd_novo.Enabled = True
  cmd_salvar.Enabled = False
  cmb_clientes.Enabled = False

  cmb_clientes.ListIndex = -1
  s_Funcao = ""
End Sub

Private Sub cmd_excluir_Click()
On Error GoTo Erro

  Dim lCodCliente As Long
  Dim lnResponse As Long
  Dim sSql As String

  If grade_programas.RowSel < 1 Then
    MsgBox "Selecione um registro na grade.", vbInformation
    Exit Sub
  End If

  lnResponse = MsgBox("Deseja realmente desvincular o registro?", vbYesNo, "Atenção")
  If lnResponse = vbNo Then
    Exit Sub
  End If

  lCodCliente = grade_programas.TextMatrix(grade_programas.RowSel, 1)
  
  'db_SQLSERVER.Execute "Delete from [ProgramaFidelidade_ClienteNaoParticipa] where CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_Cliente] = " & lCodCliente, dbSeeChanges
  
  sSql = "Delete from [ProgramaFidelidade_ClienteNaoParticipa] where CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_Cliente] = " & lCodCliente
  Dim cmd As New ADODB.Command
  cmd.ActiveConnection = gDB_SQLSERVER
  cmd.CommandText = sSql
  cmd.CommandType = adCmdText
  cmd.Execute
  Set cmd = Nothing

  MsgBox "Registro desvinculado com sucesso.", vbInformation, "Sucesso"
  
  cmd_cancelar_Click
  carregarGrade
  
  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar o registro...Detalhes do Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub cmd_novo_Click()
  grade_programas.Enabled = False
  cmd_excluir.Enabled = False
  cmb_clientes.Enabled = True
  cmd_salvar.Enabled = True

  cmb_clientes.ListIndex = -1

  s_Funcao = "INCLUIR"
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim arrOp() As String

  If s_Funcao = "INCLUIR" Then

    If cmb_clientes.Text = "" Then
      MsgBox "Informe um Cliente", vbInformation, "Atenção"
      Exit Sub
    End If

    arrOp = Split(cmb_clientes.Text, " - ")

    sSql = "Insert into ProgramaFidelidade_ClienteNaoParticipa (CNPJ,Cd_Cliente) "
    sSql = sSql & " VALUES ('" & gCNPJ_CPFControleDeLicencaWebApi & "'," & arrOp(0) & ") "
    
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = gDB_SQLSERVER
    cmd.CommandText = sSql
    cmd.CommandType = adCmdText
    cmd.Execute
    Set cmd = Nothing

'    db_SQLSERVER.Execute sSql, dbSeeChanges

'    If db_SQLSERVER.RecordsAffected > 0 Then
'      MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"
'    Else
'      MsgBox "Registro já existe na base ou algum parâmetro que vc digitou esta inconsistente", vbInformation, "Atenção"
'    End If
    cmd_cancelar_Click
    carregarGrade
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao tentar salvar o registro...Detalhes do Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
On Error GoTo Erro:
  grade_programas.ColWidth(0) = 10
  grade_programas.ColWidth(1) = 1200
  grade_programas.ColWidth(2) = 6000
  
  grade_programas.Row = 0
  grade_programas.TextMatrix(0, 1) = "Código"
  grade_programas.TextMatrix(0, 2) = "Nome do Cliente"
  
  txt_cnpj.Text = gCNPJ_CPFControleDeLicencaWebApi
  txt_nomeEmpresa.Text = gNomeEmpresaFilial
  
  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER
  
  carregarGrade
  
  Dim rsClientes As Recordset
  Set rsClientes = db.OpenRecordset("Select Código, Nome from [Cli_For] order by Nome ", dbOpenDynaset)
  
  While Not rsClientes.EOF
    cmb_clientes.AddItem rsClientes.Fields(0).Value & " - " & rsClientes.Fields(1).Value

    rsClientes.MoveNext
  Wend
  rsClientes.Close
  Set rsClientes = Nothing

  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub

Private Sub carregarGrade()
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  
  grade_programas.Rows = 1

  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM ProgramaFidelidade_ClienteNaoParticipa "
  strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"

  'Set rsPrograma = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma.Open strSQL, gDB_SQLSERVER

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF

      grade_programas.AddItem 0 & vbTab & rsPrograma.Fields("Cd_Cliente").Value

      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub
