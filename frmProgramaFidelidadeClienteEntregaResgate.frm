VERSION 5.00
Begin VB.Form frmProgramaFidelidadeClienteEntregaResgate 
   Caption         =   " Programa Fidelidade x Cliente Entrega Resgate"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmProgramaFidelidadeClienteEntregaResgate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_resgatar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "RESGATAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1650
      Width           =   10995
   End
   Begin VB.TextBox txt_nomeCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1110
      Width           =   6855
   End
   Begin VB.TextBox txt_codigoGuidResgate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   630
      Width           =   4005
   End
   Begin VB.TextBox txt_cpf 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1110
      Width           =   2355
   End
   Begin VB.Label lbl_informacao 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   630
      TabIndex        =   7
      Top             =   2250
      Width           =   9915
   End
   Begin VB.Label Label6 
      Caption         =   "Área de RESGATE DE PONTOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   150
      Width           =   3165
   End
   Begin VB.Label Label2 
      Caption         =   "Código GUID Resgate"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   690
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CPF/CNPJ do Cliente"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   1170
      Width           =   1605
   End
End
Attribute VB_Name = "frmProgramaFidelidadeClienteEntregaResgate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsClientes As Recordset
  
Private Sub cmd_resgatar_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim rsPrograma As ADODB.Recordset
  Dim rsProgramaLanc As ADODB.Recordset
  Dim dtValidadeTroca As Date
  Dim lCdPrograma As Long
  Dim iStatusResgate As Integer
  
  lbl_informacao.ForeColor = &H80000012
  lbl_informacao.Caption = ""
  
''  If LTrim(RTrim(txt_cpf.Text)) = "" Then
''      MsgBox "Digite o CPF / CNPJ do cliente!", vbInformation, "Atenção"
''      txt_cpf.SetFocus
''      Exit Sub
''  End If
  
  If LTrim(RTrim(txt_codigoGuidResgate.Text)) = "" Then
      MsgBox "Digite o Código GUID Resgate!", vbInformation, "Atenção"
      txt_codigoGuidResgate.SetFocus
      Exit Sub
  End If
  
  ' Verificar se é possível fazer o RESGATE...Se esta dentro do prazo
  sSql = "SELECT * FROM [ProgramaFidelidade_lancamentos] "
''  sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [CPF_CGC_CLIENTE]='" & LTrim(RTrim(txt_cpf.Text)) & "' and "
  sSql = sSql & " WHERE Cd_guid_resgate='" & LTrim(RTrim(txt_codigoGuidResgate.Text)) & "' and Tp_lancamento = 2 "

  Set rsProgramaLanc = New ADODB.Recordset
  rsProgramaLanc.Open sSql, gDB_SQLSERVER
  
  If rsProgramaLanc.BOF And rsProgramaLanc.EOF Then
      ' Não existe um programa de fidelidade da empresa
      MsgBox "Este Carnê de Resgate NÃO existe em nossos registros! Verifique se as informações digitadas estão corretas.", vbInformation, "Atenção"
      lbl_informacao.ForeColor = &HFF&
      lbl_informacao.Caption = "Atenção: Este Carnê Fidelidade de Resgate " & LTrim(RTrim(txt_codigoGuidResgate.Text)) & " NÃO existe em nossos registros!"
      rsProgramaLanc.Close
      Set rsProgramaLanc = Nothing
      Exit Sub
  End If
  lCdPrograma = rsProgramaLanc.Fields("Cd_programa").Value
  txt_cpf.Text = rsProgramaLanc.Fields("CPF_CGC_CLIENTE").Value
  
  If IsNull(rsProgramaLanc.Fields("Status_guid_resgate").Value) = True Then
    iStatusResgate = 0
  Else
    iStatusResgate = rsProgramaLanc.Fields("Status_guid_resgate").Value
  End If
  rsProgramaLanc.Close
  Set rsProgramaLanc = Nothing
  
  If Not IsNull(iStatusResgate) And iStatusResgate = 1 Then
      MsgBox "Este Carnê Fidelidade já foi recebido/utilizado anteriormente! ESTA INVÁLIDO!", vbInformation, "Atenção"
      lbl_informacao.ForeColor = &HFF&
      lbl_informacao.Caption = "Atenção: Este Carnê Fidelidade de Resgate " & LTrim(RTrim(txt_codigoGuidResgate.Text)) & " já foi recebido/utilizado anteriormente! ESTA INVÁLIDO!"
      Exit Sub
  End If
  
  ProgramaFidelidadeEmpresaGrupoValida
  
  ' Verificar se é possível fazer o RESGATE...Se esta dentro do prazo
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then
      sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
      sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_programa]=" & lCdPrograma
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
      sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
      sSql = sSql & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "' and [Cd_programa]=" & lCdPrograma
  End If

  Set rsPrograma = New ADODB.Recordset
  rsPrograma.Open sSql, gDB_SQLSERVER
  
  If rsPrograma.BOF And rsPrograma.EOF Then
      ' Não existe um programa de fidelidade da empresa
      MsgBox "Não existe um programa de fidelidade desta empresa.", vbInformation, "Atenção"
      lbl_informacao.ForeColor = &HFF&
      lbl_informacao.Caption = "Atenção: Não existe um programa de fidelidade vinculado a esta empresa!"
      rsPrograma.Close
      Set rsPrograma = Nothing
      Exit Sub
  End If
  dtValidadeTroca = rsPrograma.Fields("Dt_PrazoLimiteTrocaPontos").Value
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  If Now > dtValidadeTroca Then
      MsgBox "Prazo para RESGATE esta expirado! CUPOM INVÁLIDO", vbInformation, "Atenção"
      lbl_informacao.ForeColor = &HFF&
      lbl_informacao.Caption = "Atenção: Prazo para Resgate deste Carnê Fidelidade " & LTrim(RTrim(txt_codigoGuidResgate.Text)) & " esta expirado! CUPOM INVÁLIDO!"
      Exit Sub
  End If
  
  'Atualizar lançamento no programa de fidelidade - RECEBIDO = 1
  sSql = "Update [ProgramaFidelidade_lancamentos] set [Status_guid_resgate] = 1, [Dt_recebido_guid_resgate]=convert(datetime,'" & Now & "', 103), [Cd_operador_recebido_guid_resgate]=" & gnUserCode
  sSql = sSql & " Where [CNPJ]='" & gCNPJ_CPFControleDeLicencaWebApi & "' and [CPF_CGC_CLIENTE]='" & LTrim(RTrim(txt_cpf.Text)) & "' and [Cd_guid_resgate]='" & LTrim(RTrim(txt_codigoGuidResgate.Text)) & "' "
  
  'db_SQLSERVER.Execute sSql
  Dim cmd As New ADODB.Command
  cmd.ActiveConnection = gDB_SQLSERVER
  cmd.CommandText = sSql
  cmd.CommandType = adCmdText
  cmd.Execute
  Set cmd = Nothing

  MsgBox "Carnê Fidelidade recebido com sucesso!", vbInformation, "Sucesso"
  lbl_informacao.Caption = "Parabéns: Carnê Fidelidade " & LTrim(RTrim(txt_codigoGuidResgate.Text)) & " recebido com SUCESSO!"
  
  Exit Sub
Erro:
  MsgBox "Erro realizando resgate Cód: " & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
On Error GoTo Erro:

  Set rsClientes = db.OpenRecordset("Select Código, Nome, CGC from [Cli_For] order by Nome ", dbOpenDynaset)

  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER

  Exit Sub
Erro:
  MsgBox "Erro realizando a carga da tela Cód: " & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsClientes.Close
  Set rsClientes = Nothing
  
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub

Private Sub txt_cpf_LostFocus()
  Dim sAuxNome As String
  
  If LTrim(RTrim(txt_cpf.Text)) = "" Then
      txt_nomeCliente.Text = ""
      Exit Sub
  End If
  
  txt_nomeCliente.Text = ""
  rsClientes.MoveFirst
  While Not rsClientes.EOF
  
    sAuxNome = ""
    If Not IsNull(rsClientes.Fields("CGC").Value) And rsClientes.Fields("CGC").Value <> "" Then
      sAuxNome = rsClientes.Fields("CGC").Value
      sAuxNome = Replace(sAuxNome, ".", "")
      sAuxNome = Replace(sAuxNome, "/", "")
      sAuxNome = Replace(sAuxNome, "-", "")
    End If
    
    If txt_cpf.Text = sAuxNome Then
        txt_nomeCliente.Text = rsClientes.Fields("Nome").Value
        rsClientes.MoveLast
    Else
        txt_nomeCliente.Text = "Cliente sem NOME no cadastro"
    End If
    rsClientes.MoveNext
  Wend
End Sub

