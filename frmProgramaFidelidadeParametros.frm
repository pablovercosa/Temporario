VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProgramaFidelidadeParametros 
   Caption         =   " Programa de Fidelidade (Pontuação)"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgramaFidelidadeParametros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   16800
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_nomePrograma 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   660
      MaxLength       =   70
      TabIndex        =   11
      Top             =   4935
      Width           =   5535
   End
   Begin VB.CheckBox chk_programaAtivo 
      Caption         =   "Programa Ativo"
      Enabled         =   0   'False
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   4440
      Width           =   1605
   End
   Begin VB.CommandButton cmd_calendarioDt3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11220
      Picture         =   "frmProgramaFidelidadeParametros.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4342
      Width           =   465
   End
   Begin VB.TextBox txt_cnpj 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   1260
      TabIndex        =   23
      Top             =   90
      Width           =   2295
   End
   Begin VB.TextBox txt_nomeEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   3600
      TabIndex        =   22
      Top             =   90
      Width           =   13125
   End
   Begin VB.CommandButton cmd_novo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Criar novo programa"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   8295
   End
   Begin VB.CommandButton cmd_alterar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Alterar programa"
      Height          =   430
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   8295
   End
   Begin VB.CommandButton cmd_salvar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5490
      Width           =   16665
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   430
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5940
      Width           =   16665
   End
   Begin VB.CheckBox chk_indicadorProgFidelidade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Participa do Programa de Fidelidade?"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   2985
   End
   Begin VB.TextBox txt_vlProgFidelidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   7380
      TabIndex        =   12
      Top             =   4935
      Width           =   1425
   End
   Begin VB.TextBox txt_iPontosProgFidelidade 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   10500
      TabIndex        =   16
      Text            =   "1"
      Top             =   4935
      Width           =   1275
   End
   Begin VB.TextBox txt_vlProgFidelidadeParaCadaPonto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15270
      TabIndex        =   13
      Top             =   4935
      Width           =   1425
   End
   Begin VB.CommandButton cmd_calendarioDt1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6660
      Picture         =   "frmProgramaFidelidadeParametros.frx":4F23C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4342
      Width           =   465
   End
   Begin VB.CommandButton cmd_calendarioDt2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16230
      Picture         =   "frmProgramaFidelidadeParametros.frx":4FB1E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4342
      Width           =   465
   End
   Begin MSMask.MaskEdBox msk_dtInicioDoPrograma 
      Height          =   375
      Left            =   5190
      TabIndex        =   5
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   4365
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox msk_dtTrocaPontosAte 
      Height          =   375
      Left            =   14760
      TabIndex        =   9
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   4365
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid grade_programas 
      Height          =   2805
      Left            =   60
      TabIndex        =   1
      Top             =   960
      Width           =   16665
      _ExtentX        =   29395
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   1
      Cols            =   10
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
   Begin MSMask.MaskEdBox msk_dtFimDoPrograma 
      Height          =   375
      Left            =   9750
      TabIndex        =   7
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   4365
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
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
      Left            =   3630
      TabIndex        =   27
      Top             =   570
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label Label4 
      Caption         =   "Nome"
      Height          =   225
      Left            =   90
      TabIndex        =   26
      Top             =   5025
      Width           =   555
   End
   Begin VB.Label Label10 
      Caption         =   "Data fim do programa"
      Height          =   225
      Left            =   7950
      TabIndex        =   25
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Label Label9 
      Caption         =   "Empresa/Filial"
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   150
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "A cada  R$"
      Height          =   225
      Left            =   6420
      TabIndex        =   21
      Top             =   5010
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "O cliente ganha                                  pontos"
      Height          =   225
      Left            =   9240
      TabIndex        =   20
      Top             =   5010
      Width           =   3315
   End
   Begin VB.Label Label3 
      Caption         =   "E cada ponto representa R$ "
      Height          =   225
      Left            =   13020
      TabIndex        =   19
      Top             =   5010
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Data início do programa"
      Height          =   225
      Left            =   3240
      TabIndex        =   18
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "O cliente troca pontos até"
      Height          =   225
      Left            =   12720
      TabIndex        =   17
      Top             =   4440
      Width           =   2025
   End
End
Attribute VB_Name = "frmProgramaFidelidadeParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim s_Funcao As String
Dim lCodPrograma As Long


Private Sub cmd_alterar_Click()

  If grade_programas.RowSel < 1 Then
    MsgBox "Selecione um registro na grade.", vbInformation
    Exit Sub
  End If
  
  grade_programas.Enabled = False
  cmd_novo.Enabled = False
  chk_programaAtivo.Enabled = True
  msk_dtInicioDoPrograma.Enabled = True
  cmd_calendarioDt1.Enabled = True
  msk_dtFimDoPrograma.Enabled = True
  cmd_calendarioDt3.Enabled = True
  msk_dtTrocaPontosAte.Enabled = True
  cmd_calendarioDt2.Enabled = True
  txt_nomePrograma.Enabled = True
  txt_vlProgFidelidade.Enabled = True
  txt_vlProgFidelidadeParaCadaPonto.Enabled = True
  
  cmd_salvar.Enabled = True
  
  grade_programas.TextMatrix(0, 5) = "Cliente troca pontos até"
  grade_programas.TextMatrix(0, 6) = "Cada 'X R$' em compras"
  grade_programas.TextMatrix(0, 7) = "O cliente ganha 'X pts'"
  grade_programas.TextMatrix(0, 8) = "Cada ponto representa 'X R$/Cents'"

  If grade_programas.TextMatrix(grade_programas.RowSel, 9) = "ATIVO" Then
    chk_programaAtivo.Value = 1
  Else
    chk_programaAtivo.Value = 0
  End If

  lCodPrograma = grade_programas.TextMatrix(grade_programas.RowSel, 1)
  txt_nomePrograma.Text = grade_programas.TextMatrix(grade_programas.RowSel, 2)
  msk_dtInicioDoPrograma.Text = grade_programas.TextMatrix(grade_programas.RowSel, 3)
  msk_dtFimDoPrograma.Text = grade_programas.TextMatrix(grade_programas.RowSel, 4)
  msk_dtTrocaPontosAte.Text = grade_programas.TextMatrix(grade_programas.RowSel, 5)
  txt_vlProgFidelidade.Text = grade_programas.TextMatrix(grade_programas.RowSel, 6)
  txt_vlProgFidelidadeParaCadaPonto.Text = grade_programas.TextMatrix(grade_programas.RowSel, 8)
  
  s_Funcao = "ALTERAR"
End Sub

Private Sub cmd_calendarioDt1_Click()
    msk_dtInicioDoPrograma.Text = frmCalendario.gsDateCalender(msk_dtInicioDoPrograma.Text)
End Sub

Private Sub cmd_calendarioDt2_Click()
    msk_dtTrocaPontosAte.Text = frmCalendario.gsDateCalender(msk_dtTrocaPontosAte.Text)
End Sub

Private Sub cmd_calendarioDt3_Click()
    msk_dtFimDoPrograma.Text = frmCalendario.gsDateCalender(msk_dtFimDoPrograma.Text)
End Sub

Private Sub cmd_cancelar_Click()
  grade_programas.Enabled = True
  cmd_alterar.Enabled = True
  cmd_novo.Enabled = True
  chk_programaAtivo.Enabled = False
  cmd_salvar.Enabled = True
  msk_dtInicioDoPrograma.Enabled = False
  cmd_calendarioDt1.Enabled = False
  msk_dtFimDoPrograma.Enabled = False
  cmd_calendarioDt3.Enabled = False
  msk_dtTrocaPontosAte.Enabled = False
  cmd_calendarioDt2.Enabled = False
  txt_nomePrograma.Enabled = False
  txt_vlProgFidelidade.Enabled = False
  txt_vlProgFidelidadeParaCadaPonto.Enabled = False

  chk_programaAtivo.Value = False
  msk_dtInicioDoPrograma.Text = "  /  /    "
  msk_dtFimDoPrograma.Text = "  /  /    "
  msk_dtTrocaPontosAte.Text = "  /  /    "
  txt_nomePrograma.Text = ""
  txt_vlProgFidelidade.Text = ""
  txt_iPontosProgFidelidade.Text = "1"
  txt_vlProgFidelidadeParaCadaPonto.Text = ""

  s_Funcao = ""
End Sub

Private Sub cmd_novo_Click()
  grade_programas.Enabled = False
  cmd_alterar.Enabled = False
  chk_programaAtivo.Enabled = True
  cmd_salvar.Enabled = True
  msk_dtInicioDoPrograma.Enabled = True
  cmd_calendarioDt1.Enabled = True
  msk_dtFimDoPrograma.Enabled = True
  cmd_calendarioDt3.Enabled = True
  msk_dtTrocaPontosAte.Enabled = True
  cmd_calendarioDt2.Enabled = True
  txt_nomePrograma.Enabled = True
  txt_vlProgFidelidade.Enabled = True
  txt_vlProgFidelidadeParaCadaPonto.Enabled = True

  chk_programaAtivo.Value = False
  msk_dtInicioDoPrograma.Text = "  /  /    "
  msk_dtFimDoPrograma.Text = "  /  /    "
  msk_dtTrocaPontosAte.Text = "  /  /    "
  txt_nomePrograma.Text = ""
  txt_vlProgFidelidade.Text = ""
  txt_iPontosProgFidelidade.Text = "1"
  txt_vlProgFidelidadeParaCadaPonto.Text = ""

  s_Funcao = "INCLUIR"
End Sub

Private Sub cmd_salvar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim rsPrograma As ADODB.Recordset
  
  If chk_indicadorProgFidelidade.Value = 1 Then
      sSql = "Update [Parâmetros Filial] set participaProgramaFidelidade=1 where Filial =" & gnCodFilial
      gParticipaProgramaFidelidade = 1
  Else
      sSql = "Update [Parâmetros Filial] set participaProgramaFidelidade=0 where Filial =" & gnCodFilial
      gParticipaProgramaFidelidade = 0
  End If
  db.Execute sSql
  
  If s_Funcao = "" Then
    MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"
  End If
  
  If s_Funcao = "ALTERAR" Then

    If chk_programaAtivo.Value = 1 Then
      ' Verificar se já tem algum outro programa de fidelidade que esta ativo, pois só pode haver UM PROGRAMA ATIVO

      sSql = "SELECT count(*) FROM ProgramaFidelidade_empresa "
      sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
      sSql = sSql & " AND Cd_status = 1 AND Cd_programa <> " & lCodPrograma

      Set rsPrograma = New ADODB.Recordset
      rsPrograma.Open sSql, gDB_SQLSERVER
'      Set rsPrograma = db_SQLSERVER.OpenRecordset(sSql, dbOpenDynaset, dbSeeChanges)

      If rsPrograma.Fields(0).Value = 1 Then    ' ATIVO
          MsgBox "Já existe um 'Programa de Fidelidade' ATIVO para esta empresa. Só pode haver UM PROGRAMA ATIVO.", vbInformation, "Atenção"
          rsPrograma.Close
          Set rsPrograma = Nothing
          Exit Sub
      End If
      rsPrograma.Close
      Set rsPrograma = Nothing
    End If

    sSql = "Update ProgramaFidelidade_empresa set Nm_programa='" & LTrim(RTrim(txt_nomePrograma.Text)) & "', "
    sSql = sSql & " Dt_IniPrograma=convert(datetime,'" & msk_dtInicioDoPrograma.Text & "', 103), "
    sSql = sSql & " Dt_FimPrograma=convert(datetime,'" & msk_dtFimDoPrograma.Text & "', 103), "
    sSql = sSql & " Dt_PrazoLimiteTrocaPontos=convert(datetime,'" & msk_dtTrocaPontosAte.Text & "', 103), "
    
    If chk_programaAtivo.Value = 1 Then
        sSql = sSql & " [Cd_status]=1, "
    Else
        sSql = sSql & " [Cd_status]=0, "
    End If

    sSql = sSql & " [Vl_ProgFidelidade]=" & Replace(txt_vlProgFidelidade.Text, ",", ".") & ", "
    sSql = sSql & " [Vl_ProgFidelidadeParaCadaPonto]=" & Replace(txt_vlProgFidelidadeParaCadaPonto.Text, ",", ".")
    
    sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
    sSql = sSql & " AND Cd_programa = " & lCodPrograma
    
''    ws_SQLSERVER.BeginTrans
''    db_SQLSERVER.Execute sSql, dbSeeChanges
''    ws_SQLSERVER.CommitTrans

    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = gDB_SQLSERVER
    cmd.CommandText = sSql
    cmd.CommandType = adCmdText
    cmd.Execute
    Set cmd = Nothing

    MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"

    cmd_cancelar_Click
    carregarGrade
  ElseIf s_Funcao = "INCLUIR" Then
  
    If chk_programaAtivo.Value = 1 Then
      ' Verificar se já tem algum outro programa de fidelidade que esta ativo, pois só pode haver UM PROGRAMA ATIVO
  
      sSql = "SELECT count(*) FROM ProgramaFidelidade_empresa "
      sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
      sSql = sSql & " AND [Cd_status] = 1 "
  
      'Set rsPrograma = db_SQLSERVER.OpenRecordset(sSql, dbOpenDynaset, dbSeeChanges)
      Set rsPrograma = New ADODB.Recordset
      rsPrograma.Open sSql, gDB_SQLSERVER
  
      If rsPrograma.Fields(0).Value = 1 Then    ' ATIVO
          MsgBox "Já existe um 'Programa de Fidelidade' ATIVO para esta empresa. Só pode haver UM PROGRAMA ATIVO.", vbInformation, "Atenção"
          rsPrograma.Close
          Set rsPrograma = Nothing
          Exit Sub
      End If
      rsPrograma.Close
      Set rsPrograma = Nothing
    End If
  
    If msk_dtInicioDoPrograma.Text = "  /  /    " Then
      MsgBox "Informe uma Data início do programa.", vbInformation, "Atenção"
      Exit Sub
    End If
    
    If msk_dtFimDoPrograma.Text = "  /  /    " Then
      MsgBox "Informe uma Data fim do programa.", vbInformation, "Atenção"
      Exit Sub
    End If
    
    If msk_dtTrocaPontosAte.Text = "  /  /    " Then
      MsgBox "Informe uma Data 'Troca Pontos Até' do programa.", vbInformation, "Atenção"
      Exit Sub
    End If
    
    If txt_nomePrograma.Text = "" Then
      MsgBox "Informe o nome do programa.", vbInformation, "Atenção"
      Exit Sub
    End If
    
    If txt_vlProgFidelidade.Text = "" Then
      MsgBox "Informe o valor R$ do programa.", vbInformation, "Atenção"
      Exit Sub
    End If
    
    If txt_vlProgFidelidadeParaCadaPonto.Text = "" Then
      MsgBox "Informe o valor R$ que quanto representa cada ponto do programa.", vbInformation, "Atenção"
      Exit Sub
    End If
    
    Dim iProgramaAtivo As Integer
    If chk_programaAtivo.Value = True Then
      iProgramaAtivo = 1
    Else
      iProgramaAtivo = 0
    End If
  
    sSql = "Insert into ProgramaFidelidade_empresa (Nm_programa,CNPJ,Dt_criacao,"
    sSql = sSql & " Dt_IniPrograma , Dt_FimPrograma, "
    sSql = sSql & " Dt_PrazoLimiteTrocaPontos,Cd_status,Vl_ProgFidelidade,Nm_PontosProgFidelidade,"
    sSql = sSql & " Vl_ProgFidelidadeParaCadaPonto) "
    sSql = sSql & " VALUES ('" & LTrim(RTrim(txt_nomePrograma.Text)) & "','" & gCNPJ_CPFControleDeLicencaWebApi & "', convert(datetime,'" & Now & "', 103), "
    sSql = sSql & " convert(datetime,'" & msk_dtInicioDoPrograma.Text & "', 103), convert(datetime,'" & msk_dtFimDoPrograma.Text & "', 103),"
    sSql = sSql & " convert(datetime,'" & msk_dtTrocaPontosAte.Text & "', 103)," & iProgramaAtivo & "," & Replace(txt_vlProgFidelidade.Text, ",", ".") & ",1," & Replace(txt_vlProgFidelidadeParaCadaPonto.Text, ",", ".") & ") "
    
    Dim cmd2 As New ADODB.Command
    cmd2.ActiveConnection = gDB_SQLSERVER
    cmd2.CommandText = sSql
    cmd2.CommandType = adCmdText
    cmd2.Execute
    Set cmd2 = Nothing

    'db_SQLSERVER.Execute sSql

    MsgBox "Registro salvo com sucesso", vbInformation, "Sucesso"

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
  grade_programas.ColWidth(1) = 700
  grade_programas.ColWidth(2) = 2980
  grade_programas.ColWidth(3) = 1200
  grade_programas.ColWidth(4) = 1200
  grade_programas.ColWidth(5) = 2000
  grade_programas.ColWidth(6) = 2000
  grade_programas.ColWidth(7) = 1770
  grade_programas.ColWidth(8) = 2800
  grade_programas.ColWidth(9) = 1700
  
  grade_programas.Row = 0
  grade_programas.TextMatrix(0, 1) = "Código"
  grade_programas.TextMatrix(0, 2) = "Nome do Programa Fidelidade"
  grade_programas.TextMatrix(0, 3) = "Data Início"
  grade_programas.TextMatrix(0, 4) = "Data Fim"
  grade_programas.TextMatrix(0, 5) = "Cliente troca pontos até"
  grade_programas.TextMatrix(0, 6) = "Cada 'X R$' em compras"
  grade_programas.TextMatrix(0, 7) = "O cliente ganha 'X pts'"
  grade_programas.TextMatrix(0, 8) = "Cada ponto representa 'X R$/Cents'"
  grade_programas.TextMatrix(0, 9) = "Status do programa"
  
  txt_cnpj.Text = gCNPJ_CPFControleDeLicencaWebApi
  txt_nomeEmpresa.Text = gNomeEmpresaFilial
  
  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER
  
  ProgramaFidelidadeEmpresaGrupoValida
  

  carregarGrade
  
  If gParticipaProgramaFidelidade = 1 Then
      chk_indicadorProgFidelidade.Value = 1
  Else
      chk_indicadorProgFidelidade.Value = 0
  End If
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Public Sub carregarGrade()
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  
  grade_programas.Rows = 1

  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String
  
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Or gIndicadorProgramaFidelidadeCNPJPrincipal = 3 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresa "
    strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresa "
    strSQL = strSQL & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "'"
    lbl_informacao.Caption = "Atenção: CNPJ logado participa do Grupo prog.fidelidade CNPJ " & gCNPJProgramaFidelidadeCNPJPrincipal
    lbl_informacao.Visible = True
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_salvar.Enabled = True
    cmd_cancelar.Enabled = False
  End If
 
  'Set rsPrograma = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma.Open strSQL, gDB_SQLSERVER

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF

      iStatus = rsPrograma.Fields("Cd_status").Value
      If iStatus = 1 Then
          sStatus = "ATIVO"
      Else
          sStatus = "INATIVO"
      End If
      grade_programas.AddItem 0 & vbTab & rsPrograma.Fields("Cd_programa").Value & vbTab & _
                      rsPrograma.Fields("Nm_programa").Value & vbTab & _
                      rsPrograma.Fields("Dt_IniPrograma").Value & vbTab & _
                      rsPrograma.Fields("Dt_FimPrograma").Value & vbTab & _
                      rsPrograma.Fields("Dt_PrazoLimiteTrocaPontos").Value & vbTab & _
                      Format(rsPrograma.Fields("Vl_ProgFidelidade").Value, FORMAT_VALUE) & vbTab & _
                      rsPrograma.Fields("Nm_PontosProgFidelidade").Value & vbTab & _
                      Format(rsPrograma.Fields("Vl_ProgFidelidadeParaCadaPonto").Value, FORMAT_VALUE) & vbTab & _
                      sStatus

      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub
