VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProgramaFidelidadeCNPJGrupos 
   Caption         =   " Programa Fidelidade x CNPJs Participantes"
   ClientHeight    =   6195
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
   Icon            =   "frmProgramaFidelidadeCNPJGrupos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_cnpjVincular 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   660
      TabIndex        =   10
      Top             =   4650
      Width           =   2295
   End
   Begin VB.TextBox txt_cnpj 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   120
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   390
      Width           =   6255
   End
   Begin VB.CommandButton cmd_novo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vincular CNPJ"
      Height          =   430
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   4245
   End
   Begin VB.CommandButton cmd_excluir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Desvincular"
      Height          =   430
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   5190
      Width           =   8595
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   430
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5700
      Width           =   8595
   End
   Begin MSFlexGridLib.MSFlexGrid grade_programas 
      Height          =   2805
      Left            =   120
      TabIndex        =   6
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
   Begin VB.Label lbl_informacao 
      Caption         =   "Atenção: CNPJ logado participa do Grupo prog.fidelidade CNPJ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1410
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Label Label4 
      Caption         =   "CNPJs vinculados ao Programa de Fidelidade"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   930
      Width           =   3525
   End
   Begin VB.Label Label9 
      Caption         =   "Empresa/Filial"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "CNPJ"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4710
      Width           =   525
   End
End
Attribute VB_Name = "frmProgramaFidelidadeCNPJGrupos"
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
  txt_cnpjVincular.Enabled = False
  txt_cnpjVincular.Text = ""

  s_Funcao = ""
End Sub

Private Sub cmd_excluir_Click()
On Error GoTo Erro

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

  sSql = "Delete from ProgramaFidelidade_empresaGrupo where CNPJ_Principal = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and CNPJ = '" & grade_programas.TextMatrix(grade_programas.RowSel, 2) & "'"
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
  txt_cnpjVincular.Enabled = True
  cmd_salvar.Enabled = True

  txt_cnpjVincular.Text = ""

  s_Funcao = "INCLUIR"
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim arrOp() As String

  If s_Funcao = "INCLUIR" Then

    If Len(LTrim(RTrim(txt_cnpjVincular.Text))) = 0 Then
      MsgBox "Informe um CNPJ Participante", vbInformation, "Atenção"
      txt_cnpjVincular.SetFocus
      Exit Sub
    End If

    sSql = "Insert into ProgramaFidelidade_empresaGrupo (CNPJ_principal, CNPJ) "
    sSql = sSql & " VALUES ('" & gCNPJ_CPFControleDeLicencaWebApi & "','" & txt_cnpjVincular.Text & "') "

    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = gDB_SQLSERVER
    cmd.CommandText = sSql
    cmd.CommandType = adCmdText
    cmd.Execute
    Set cmd = Nothing

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
  grade_programas.ColWidth(1) = 3000
  grade_programas.ColWidth(2) = 3000
  
  grade_programas.Row = 0
  grade_programas.TextMatrix(0, 1) = "CNPJ Principal do Programa"
  grade_programas.TextMatrix(0, 2) = "CNPJ Participante"
  
  txt_cnpj.Text = gCNPJ_CPFControleDeLicencaWebApi
  txt_nomeEmpresa.Text = gNomeEmpresaFilial
  
  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER
  
  ProgramaFidelidadeEmpresaGrupoValida
  
  carregarGrade
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub

Public Sub carregarGrade()
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  
  grade_programas.Rows = 1

  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String

  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Or gIndicadorProgramaFidelidadeCNPJPrincipal = 3 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresaGrupo "
    strSQL = strSQL & " WHERE CNPJ_principal = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresaGrupo "
    strSQL = strSQL & " WHERE CNPJ_principal = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "'"
    lbl_informacao.Caption = "Atenção: CNPJ logado participa do Grupo prog.fidelidade CNPJ " & gCNPJProgramaFidelidadeCNPJPrincipal
    lbl_informacao.Visible = True
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_salvar.Enabled = True
    cmd_cancelar.Enabled = False
  End If

  rsPrograma.Open strSQL, gDB_SQLSERVER

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF

      grade_programas.AddItem 0 & vbTab & rsPrograma.Fields("CNPJ_principal").Value & vbTab & rsPrograma.Fields("CNPJ").Value

      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub
