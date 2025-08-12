VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmProgramaFidelidadeResgatePontos_Rapido 
   Caption         =   " Programa Fidelidade x Resgate Pontos RÁPIDO"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgramaFidelidadeResgatePontos_Rapido.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10935
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opt_tipoImpressaoREL 
      Appearance      =   0  'Flat
      Caption         =   "Modelo Relatório"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2940
      TabIndex        =   26
      Top             =   4680
      Width           =   1605
   End
   Begin VB.OptionButton opt_tipoImpressaoTICKET 
      Appearance      =   0  'Flat
      Caption         =   "Modelo Ticket"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1380
      TabIndex        =   25
      Top             =   4680
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox picture_statusProcessamento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   150
      Picture         =   "frmProgramaFidelidadeResgatePontos_Rapido.frx":4E95A
      ScaleHeight     =   615
      ScaleWidth      =   855
      TabIndex        =   22
      Top             =   180
      Width           =   855
   End
   Begin VB.CommandButton cmd_receberResgate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "(Loja) Receber o Resgate dos pontos do cliente"
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
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4170
      Width           =   5265
   End
   Begin VB.CommandButton cmd_resgatar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Baixar pontos de Resgate para o cliente"
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
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2970
      Width           =   10695
   End
   Begin VB.CommandButton cmd_imprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir CUPOM"
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
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4170
      Width           =   5265
   End
   Begin VB.TextBox txt_totalpontos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   2670
      TabIndex        =   13
      Top             =   2070
      Width           =   1455
   End
   Begin VB.TextBox txt_resgatarPontos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   2070
      Width           =   1455
   End
   Begin VB.TextBox txt_totalReais 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   2670
      TabIndex        =   12
      Top             =   2490
      Width           =   1455
   End
   Begin VB.TextBox txt_saldoReaisParaResgate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmd_mostraSaldo 
      BackColor       =   &H80000000&
      Caption         =   "Mostra Saldo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2070
      Width           =   1545
   End
   Begin VB.TextBox txt_nomeCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   10
      Top             =   1260
      Width           =   6105
   End
   Begin VB.TextBox txt_codCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   2670
      TabIndex        =   9
      Top             =   1260
      Width           =   1845
   End
   Begin VB.TextBox txt_cpf 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   1260
      Width           =   2355
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   -30
      Top             =   5010
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      Caption         =   "Último Resgate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   24
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblUltimoResgate 
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
      Height          =   1125
      Left            =   450
      TabIndex        =   23
      Top             =   6300
      Width           =   9915
   End
   Begin VB.Label lbl_vendeMais 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   1080
      TabIndex        =   21
      Top             =   90
      Width           =   9735
   End
   Begin VB.Label lbl_guid 
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
      Height          =   495
      Left            =   450
      TabIndex        =   20
      Top             =   3510
      Width           =   9915
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
      Left            =   450
      TabIndex        =   19
      Top             =   5040
      Width           =   9915
   End
   Begin VB.Label Label6 
      Caption         =   "Área de RESGATE DE PONTOS"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   150
      TabIndex        =   18
      Top             =   1770
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Total de pontos acumulados"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   2130
      Width           =   2205
   End
   Begin VB.Label Label8 
      Caption         =   "Informar a qtde de pontos para RESGATE"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4530
      TabIndex        =   16
      Top             =   2130
      Width           =   3225
   End
   Begin VB.Label Label9 
      Caption         =   "Saldo em R$  acumulados"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   150
      TabIndex        =   15
      Top             =   2550
      Width           =   2085
   End
   Begin VB.Label Label10 
      Caption         =   "Saldo em R$  para RESGATE"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4530
      TabIndex        =   14
      Top             =   2580
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Nome"
      Height          =   225
      Left            =   4740
      TabIndex        =   7
      Top             =   1020
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Cód"
      Height          =   225
      Left            =   2700
      TabIndex        =   6
      Top             =   1020
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "CPF/CNPJ do Cliente"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   1020
      Width           =   1605
   End
End
Attribute VB_Name = "frmProgramaFidelidadeResgatePontos_Rapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lCodigoCliente As Long
Dim dVl_ProgFidelidadeParaCadaPonto As Double
Public vTotalDaVendaEmAndamento As Double
Dim dVl_ProgFidelidade As Double
Dim lCodPrograma As Long
Dim sNomePrograma As String
  
Private Sub cmd_imprimir_Click()
On Error GoTo Erro:
  Dim strNomeArq As String
  Dim sEndereco As String
  Dim sBairro As String
  Dim sCidadeEstado As String
  Dim sFone As String
  Dim sNomeCliente As String
  Dim iStatusResgate As Integer
  Dim sVl_SaldoEmReais As String

  If gCdGuidResgate = "" Then
    MsgBox "Realize o resgate antes de Imprimir o CUPOM", vbInformation, "Atenção"
    Exit Sub
  End If

  'Verificar se o CUPOM já foi utilizado
  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CNPJ,CPF_CGC_CLIENTE,Cd_programa,Cd_cliente,Dt_criacao,Vl_CompraCliente, "
  strSQL = strSQL & " Nm_PontosAdquiridos,Vl_SaldoEmReais,Tp_lancamento,Cd_operador,Cd_guid_resgate,"
  strSQL = strSQL & " Status_guid_resgate,Dt_recebido_guid_resgate,Cd_operador_recebido_guid_resgate "
  strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos "
  strSQL = strSQL & " WHERE Cd_guid_resgate = '" & gCdGuidResgate & "' "

  rsPrograma.Open strSQL, gDB_SQLSERVER

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  Else
    rsPrograma.Close
    Set rsPrograma = Nothing
    Exit Sub
  End If
  
  sVl_SaldoEmReais = rsPrograma.Fields("Vl_SaldoEmReais").Value
  
  If IsNull(rsPrograma.Fields("Status_guid_resgate").Value) = True Then
    iStatusResgate = 0
  Else
    iStatusResgate = rsPrograma.Fields("Status_guid_resgate").Value
  End If
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  If Not IsNull(iStatusResgate) And iStatusResgate = 1 Then
      MsgBox "Este Carnê Fidelidade já foi recebido/utilizado anteriormente! ESTA INVÁLIDO!", vbInformation, "Atenção"
      Exit Sub
  End If


  Dim rsParamFilial As Recordset
  Set rsParamFilial = db.OpenRecordset("Select * FROM [Parâmetros Filial] where Filial= " & gnCodFilial, dbOpenDynaset)
  
  If IsNull(rsParamFilial.Fields("Endereço").Value) Then
      sEndereco = "Endereço não cadastrado"
  Else
      sEndereco = rsParamFilial.Fields("Endereço").Value
  End If
  
  If IsNull(rsParamFilial.Fields("Bairro").Value) Then
      sEndereco = "Bairro não cadastrado"
  Else
      sBairro = rsParamFilial.Fields("Bairro").Value
  End If
  
  If IsNull(rsParamFilial.Fields("Cidade").Value) And IsNull(rsParamFilial.Fields("Estado").Value) Then
      sCidadeEstado = "Cidade/Estado não cadastrado"
  Else
      If IsNull(rsParamFilial.Fields("Cidade").Value) And Not IsNull(rsParamFilial.Fields("Estado").Value) Then
          sCidadeEstado = "Residente no estado do " & rsParamFilial.Fields("Estado").Value
      ElseIf Not IsNull(rsParamFilial.Fields("Cidade").Value) And IsNull(rsParamFilial.Fields("Estado").Value) Then
          sCidadeEstado = rsParamFilial.Fields("Cidade").Value
      Else
          sCidadeEstado = rsParamFilial.Fields("Cidade").Value & " - " & rsParamFilial.Fields("Estado").Value
      End If
  End If
  
  If IsNull(rsParamFilial.Fields("Fone").Value) Then
      sEndereco = "Fone não cadastrado"
  Else
      sBairro = rsParamFilial.Fields("Fone").Value
  End If
  
  rsParamFilial.Close
  Set rsParamFilial = Nothing

  Dim rsClientes As Recordset
  Set rsClientes = db.OpenRecordset("Select Código, Nome, CGC from [Cli_For] order by Nome ", dbOpenDynaset)
  Dim sAuxNome As String
  While Not rsClientes.EOF
  
    sAuxNome = ""
    If Not IsNull(rsClientes.Fields("CGC").Value) And rsClientes.Fields("CGC").Value <> "" Then
        sAuxNome = rsClientes.Fields("CGC").Value
        sAuxNome = Replace(sAuxNome, ".", "")
        sAuxNome = Replace(sAuxNome, "/", "")
        sAuxNome = Replace(sAuxNome, "-", "")
    End If
    
    If txt_cpf.Text = sAuxNome Then
        If Not IsNull(rsClientes.Fields("Nome").Value) Then
            sNomeCliente = rsClientes.Fields("Nome").Value
        Else
            sNomeCliente = "Cliente sem NOME no cadastro"
        End If

        rsClientes.MoveLast
    End If
    rsClientes.MoveNext
  Wend
  rsClientes.Close
  Set rsClientes = Nothing
  
  If opt_tipoImpressaoTICKET.Value = True Then
      strNomeArq = gsReportPath & "programaFidelidade01_45Col.rpt"
  Else
      strNomeArq = gsReportPath & "programaFidelidade01.rpt"
  End If
  
  If Dir(strNomeArq) = "" Then
    DisplayMsg "Arquivo """ & strNomeArq & """ não encontrado."
    Exit Sub
  End If
  
  CrystalReport1.DataFiles(0) = gsQuickDBFileName
  CrystalReport1.Destination = 0
  CrystalReport1.ReportFileName = strNomeArq
  CrystalReport1.ParameterFields(0) = "NomeEmpresa;" & gNomeEmpresaFilial & ";true"
  CrystalReport1.ParameterFields(1) = "EnderecoEmpresa;" & sEndereco & ";true"
  CrystalReport1.ParameterFields(2) = "NomeCliente;" & sNomeCliente & ";true"
  CrystalReport1.ParameterFields(3) = "CpfCnpjCliente;" & txt_cpf.Text & ";true"
  CrystalReport1.ParameterFields(4) = "NomeProgramaFidelidade;" & sNomePrograma & ";true"
  CrystalReport1.ParameterFields(5) = "DataResgate;" & Format(sEndereco, "dd/mm/yyyy") & ";true"
  CrystalReport1.ParameterFields(6) = "ValorResgate;" & sVl_SaldoEmReais & ";true"
  CrystalReport1.ParameterFields(7) = "EnderecoEmpresa2;" & sBairro & ", " & sCidadeEstado & ", " & sFone & "" & ";true"
  CrystalReport1.ParameterFields(8) = "CodGuidResgate;" & gCdGuidResgate & ";true"
  CrystalReport1.WindowState = crptMaximized
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 CrystalReport1

  If opt_tipoImpressaoTICKET.Value = True Then
      Call SetPrinterName("TICKET", CrystalReport1)
  Else
      Call SetPrinterName("REL", CrystalReport1)
  End If

  CrystalReport1.Action = 1

  Exit Sub
Erro:
  MsgBox "Erro tentando gerar Carnê do Programa Fidelidade. Cód: " & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_mostraSaldo_Click()
On Error GoTo Erro
  Dim lTotalPontosJaAdquiridos As Long
  Dim lPontos As Long
  Dim dSaldo As Double

  Dim arrAux() As String
  Dim iCont As Integer
  
  
  If Len(LTrim(RTrim(txt_totalpontos.Text))) = 0 Then
      MsgBox "Campo 'Total de pontos acumulados' deve conter valor", vbInformation, "Atenção"
      Exit Sub
  Else
        lTotalPontosJaAdquiridos = CLng(txt_totalpontos.Text)
  End If
  
  If Len(LTrim(RTrim(txt_resgatarPontos.Text))) > 0 Then
      lPontos = CLng(txt_resgatarPontos.Text)
      dSaldo = lPontos * dVl_ProgFidelidadeParaCadaPonto
      
      If lTotalPontosJaAdquiridos < lPontos Then
        MsgBox "RESGATE deve ser igual ou menor ao 'Total de pontos acumulados'", vbInformation, "Atenção"
        txt_saldoReaisParaResgate.Text = ""
        Exit Sub
      End If
  
      txt_saldoReaisParaResgate.Text = Format(dSaldo, FORMAT_VALUE)
  Else
      txt_saldoReaisParaResgate.Text = 0
  End If

  Exit Sub
Erro:
  MsgBox "Entre com valores numéricos", vbInformation, "Atenção"
End Sub

Private Sub cmd_receberResgate_Click()
On Error GoTo Erro:
  Dim sSql As String
  Dim rsPrograma As ADODB.Recordset
  Dim rsProgramaLanc As ADODB.Recordset
  Dim dtValidadeTroca As Date
  Dim iStatusResgate As Integer
  Dim sCpf_cnpjCli As String
  
  If gCdGuidResgate = "" Then
      MsgBox "Sem o código GUID não é possível receber o Resgate!", vbInformation, "Atenção"
      Exit Sub
  End If
  
  ' Verificar se é possível fazer o RESGATE...Se esta dentro do prazo
  sSql = "SELECT * FROM [ProgramaFidelidade_lancamentos] "
''  sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [CPF_CGC_CLIENTE]='" & LTrim(RTrim(txt_cpf.Text)) & "' and "
  sSql = sSql & " WHERE Cd_guid_resgate='" & gCdGuidResgate & "' and Tp_lancamento = 2 "

  Set rsProgramaLanc = New ADODB.Recordset
  rsProgramaLanc.Open sSql, gDB_SQLSERVER
  
  If rsProgramaLanc.BOF And rsProgramaLanc.EOF Then
      ' Não existe um programa de fidelidade da empresa
      MsgBox "Este Carnê " & gCdGuidResgate & " de Resgate NÃO existe em nossos registros! Verifique se as informações estão corretas.", vbInformation, "Atenção"
      rsProgramaLanc.Close
      Set rsProgramaLanc = Nothing
      Exit Sub
  End If
  sCpf_cnpjCli = rsProgramaLanc.Fields("CPF_CGC_CLIENTE").Value
  
  If IsNull(rsProgramaLanc.Fields("Status_guid_resgate").Value) = True Then
    iStatusResgate = 0
  Else
    iStatusResgate = rsProgramaLanc.Fields("Status_guid_resgate").Value
  End If
  rsProgramaLanc.Close
  Set rsProgramaLanc = Nothing
  
  If Not IsNull(iStatusResgate) And iStatusResgate = 1 Then
      MsgBox "Este Carnê Fidelidade " & gCdGuidResgate & " já foi recebido/utilizado anteriormente! ESTA INVÁLIDO!", vbInformation, "Atenção"
      Exit Sub
  End If
  
  ProgramaFidelidadeEmpresaGrupoValida
  
  ' Verificar se é possível fazer o RESGATE...Se esta dentro do prazo
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then
      sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
      sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_programa]=" & lCodPrograma
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
      sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
      sSql = sSql & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "' and [Cd_programa]=" & lCodPrograma
  End If

  Set rsPrograma = New ADODB.Recordset
  rsPrograma.Open sSql, gDB_SQLSERVER
  
  If rsPrograma.BOF And rsPrograma.EOF Then
      ' Não existe um programa de fidelidade da empresa
      MsgBox "Não existe um programa de fidelidade desta empresa.", vbInformation, "Atenção"
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
      lbl_informacao.Caption = "Atenção: Prazo para Resgate deste Carnê Fidelidade " & gCdGuidResgate & " esta expirado! CUPOM INVÁLIDO!"
      Exit Sub
  End If
  
  'Atualizar lançamento no programa de fidelidade - RECEBIDO = 1
  sSql = "Update [ProgramaFidelidade_lancamentos] set [Status_guid_resgate] = 1, [Dt_recebido_guid_resgate]=convert(datetime,'" & Now & "', 103), [Cd_operador_recebido_guid_resgate]=" & gnUserCode
  sSql = sSql & " Where [CNPJ]='" & gCNPJ_CPFControleDeLicencaWebApi & "' and [CPF_CGC_CLIENTE]='" & sCpf_cnpjCli & "' and [Cd_guid_resgate]='" & gCdGuidResgate & "' "
  
  'db_SQLSERVER.Execute sSql
  Dim cmd As New ADODB.command
  cmd.ActiveConnection = gDB_SQLSERVER
  cmd.CommandText = sSql
  cmd.CommandType = adCmdText
  cmd.Execute
  Set cmd = Nothing

  'MsgBox "Carnê Fidelidade recebido com sucesso!", vbInformation, "Sucesso"
  lbl_informacao.Caption = "Parabéns: Carnê Fidelidade " & gCdGuidResgate & " recebido com SUCESSO!"
  
  gClienteEntregouResgatePontos = True
  
  Exit Sub
Erro:
  MsgBox "Erro realizando recebimento de resgate Cód: " & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_resgatar_Click()
On Error GoTo Erro
  Dim sCPF As String
  Dim lTotalPontosJaAdquiridos As Long
  Dim lPontos As Long
  Dim dSaldo As Double
  Dim dSaldoPorPonto As Double
  Dim arrAux() As String
  Dim iCont As Integer
  Dim iRet As Integer
  
  If LTrim(RTrim(txt_cpf.Text)) = "" Then
      MsgBox "CPF/CNPJ do Cliente não identificado para realizar Resgate", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If Len(LTrim(RTrim(txt_totalpontos.Text))) = 0 Or txt_resgatarPontos.Text = "" Or txt_resgatarPontos.Text = "0" Then
      MsgBox "Campo 'Total de pontos acumulados' deve conter valor", vbInformation, "Atenção"
      Exit Sub
  Else
        lTotalPontosJaAdquiridos = CLng(txt_totalpontos.Text)
  End If
  
  If Len(LTrim(RTrim(txt_resgatarPontos.Text))) > 0 Then
      Dim iRetorno As Integer

      iRetorno = validaSaldoPontosNovamente

      If iRetorno = -1 Then
          'MsgBox "Saldo insuficiente! RESGATE deve ser igual ou menor ao 'Total de pontos acumulados'", vbInformation, "Atenção"
          'txt_saldoReaisParaResgate.Text = ""
          Exit Sub
      End If

      lPontos = CLng(txt_resgatarPontos.Text)
      dSaldo = lPontos * dVl_ProgFidelidadeParaCadaPonto
      
''      If lTotalPontosJaAdquiridos < lPontos Then
''        MsgBox "RESGATE deve ser igual ou menor ao 'Total de pontos acumulados'", vbInformation, "Atenção"
''        txt_saldoReaisParaResgate.Text = ""
''        Exit Sub
''      End If
  
      txt_saldoReaisParaResgate.Text = dSaldo
  Else
      txt_saldoReaisParaResgate.Text = 0
      MsgBox "Digite o número de pontos para RESGATE", vbInformation, "Atenção"
      txt_resgatarPontos.SetFocus
      Exit Sub
  End If
  
  gCdGuidResgate = ""
  iRet = ProgramaFidelidadeCriarLancamentoRESGATE(LTrim(RTrim(txt_cpf.Text)), lCodPrograma, lPontos, dSaldo, txt_nomeCli.Text)

  If iRet = 0 Then
      'MsgBox "Resgate realizado com sucesso", vbInformation, "Sucesso"
      lbl_guid.Caption = "Guid: " & gCdGuidResgate & " Resgate de " & lPontos & " pontos foi realizado com sucesso!"
      cmd_resgatar.Enabled = False
      
      gCdClienteCdGuidResgate = CLng(txt_codCli.Text)
      gNmClienteCdGuidResgate = txt_nomeCli.Text
      gSaldoCdGuidResgate = dSaldo
  Else
      lbl_guid.Caption = "ERRO ! Inconsistência no lançamento de RESGATE !!"
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao tentar realizar Resgate de Pontos..Detalhes do Erro: " & Err.Number & " " & Err.Description, vbCritical, "Erro"

End Sub

Private Function validaSaldoPontosNovamente() As Integer
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  Dim arrProg() As String
  Dim sCPF As String
  
  If LTrim(RTrim(txt_cpf.Text)) = "" Then
      MsgBox "CPF/CNPJ do Cliente não identificado", vbInformation, "Atenção"
      validaSaldoPontosNovamente = -1
      Exit Function
  End If
  
  sCPF = LTrim(RTrim(txt_cpf.Text))
  sCPF = Replace(sCPF, "-", "")
  sCPF = Replace(sCPF, ".", "")
  sCPF = Replace(sCPF, ";", "")
  sCPF = Replace(sCPF, "/", "")
  sCPF = Replace(sCPF, "\", "")


  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CNPJ,CPF_CGC_CLIENTE,Cd_programa,Cd_cliente,Dt_criacao,Vl_CompraCliente, "
  strSQL = strSQL & " Nm_PontosAdquiridos,Vl_SaldoEmReais,Tp_lancamento,Cd_operador,Cd_guid_resgate,"
  strSQL = strSQL & " Status_guid_resgate,Dt_recebido_guid_resgate,Cd_operador_recebido_guid_resgate "
  strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos "
  'strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  strSQL = strSQL & " WHERE Cd_programa = " & lCodPrograma
  strSQL = strSQL & " AND CPF_CGC_CLIENTE = '" & sCPF & "' "
  strSQL = strSQL & " ORDER BY Dt_criacao ASC "

  'Set rsPrograma = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma.Open strSQL, gDB_SQLSERVER

  Dim sTp_lancamento As String
  Dim lTotalPtsAcumulados As Long
  
  lTotalPtsAcumulados = 0

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF
  
      If rsPrograma.Fields("Tp_lancamento").Value = 1 Then
          sTp_lancamento = "COMPRA"
          lTotalPtsAcumulados = lTotalPtsAcumulados + rsPrograma.Fields("Nm_PontosAdquiridos").Value
      Else
          sTp_lancamento = "RESGATE"
          lTotalPtsAcumulados = lTotalPtsAcumulados - rsPrograma.Fields("Nm_PontosAdquiridos").Value
      End If
      
      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  Dim lPontos As Long
  
  If Len(LTrim(RTrim(txt_resgatarPontos.Text))) > 0 Then
    lPontos = CLng(txt_resgatarPontos.Text)
    
    If lTotalPtsAcumulados < lPontos Then
      MsgBox "Saldo insuficiente! RESGATE deve ser igual ou menor ao 'Total de pontos acumulados'", vbInformation, "Atenção"
      validaSaldoPontosNovamente = -1
      Exit Function
    End If
  End If

  validaSaldoPontosNovamente = 0
  
  Exit Function
Erro:
  MsgBox "Erro ao validar pontos novamente...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Function

Private Sub Form_Load()
On Error GoTo Erro:
  Dim rsCliFor As Recordset
  Dim strSQL As String
  
  txt_totalpontos.Text = ""
  txt_totalReais.Text = ""
  txt_resgatarPontos.Text = ""
  txt_saldoReaisParaResgate.Text = ""
  
  txt_codCli.Text = lCodigoCliente
  
  strSQL = "SELECT CGC, Nome From Cli_For Where Código = " & lCodigoCliente
  
  Set rsCliFor = db.OpenRecordset(strSQL, dbOpenSnapshot)

  If rsCliFor.EOF And rsCliFor.BOF Then
    MsgBox "Cliente não identificado!", vbInformation, "Atenção"
    rsCliFor.Close
    Set rsCliFor = Nothing
    Unload Me
  End If

  txt_cpf.Text = rsCliFor.Fields(0).Value
  txt_nomeCli.Text = rsCliFor.Fields(1).Value
  rsCliFor.Close
  Set rsCliFor = Nothing
  
  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER
  
  ProgramaFidelidadeEmpresaGrupoValida
  
  If gIndicadorProgramaFidelidadeCNPJPrincipal <> 3 Then
  
    recuperarPontosFidelidade
  
    'Calcula o VENDE MAIS...
    Dim dVendeMais As Double
    Dim sAUX As String
    Dim sAUX2 As String
    Dim l_Nm_PontosAdquiridosVendeMais As Integer
    
    If dVl_ProgFidelidade = 0 Then
      cmd_imprimir.Enabled = False
      cmd_mostraSaldo.Enabled = False
      cmd_receberResgate.Enabled = False
      cmd_resgatar.Enabled = False
      txt_resgatarPontos.Enabled = False
      Exit Sub
    End If
    
    dVendeMais = vTotalDaVendaEmAndamento / dVl_ProgFidelidade
    
    
    l_Nm_PontosAdquiridosVendeMais = Int(vTotalDaVendaEmAndamento / dVl_ProgFidelidade)
    dVendeMais = l_Nm_PontosAdquiridosVendeMais * dVl_ProgFidelidade
    
    dVendeMais = vTotalDaVendaEmAndamento - dVendeMais
    dVendeMais = dVl_ProgFidelidade - dVendeMais
    dVendeMais = FormataValorTexto(dVendeMais, 2)
    
    sAUX = CStr(dVendeMais)
    
    If Len(sAUX) > 2 Then
      sAUX2 = Mid(sAUX, 1, Len(sAUX) - 2) & ","
      sAUX2 = sAUX2 & Mid(sAUX, Len(sAUX) - 1, 2)
      sAUX = sAUX2
    ElseIf Len(sAUX) = 2 Then
      sAUX = "0," & sAUX
    ElseIf Len(sAUX) = 1 Then
      sAUX = "0,0" & sAUX
    End If
    
    lbl_vendeMais.Caption = "Com mais R$ " & sAUX & " nesta compra você adquire 1 ponto a mais em nosso Programa de Fidelidade!!"
    
  '  If gClienteEntregouResgatePontos = True Then
    If gCdGuidResgate <> "" Then
        lblUltimoResgate.Caption = "Cliente: " & gCdClienteCdGuidResgate & " - " & gNmClienteCdGuidResgate & vbCrLf _
          & " Saldo do Resgate de: R$ " & Format(gSaldoCdGuidResgate, FORMAT_VALUE) & vbCrLf _
          & " Guid: " & gCdGuidResgate
    End If
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro na carga da tela Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub recuperarPontosFidelidade()
On Error GoTo Erro
  Dim arrProg() As String
  Dim sCPF As String
  Dim sSql As String
  Dim rsPrograma As ADODB.Recordset
  'Dim dVl_ProgFidelidade As Double
  'Dim dVl_ProgFidelidadeParaCadaPonto As Double
  Dim l_Nm_PontosAdquiridos As Long
  Dim d_Vl_SaldoEmReais As Double
  Dim dtValidadeInicio As Date
  Dim dtValidadeFim As Date
  
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then  '1-CNPJ PRINCIPAL;  2-CNPJ VINCULADO;   3-NADA
    ' Buscar o programa de fidelidade da empresa que esta ATIVO (só pode haver um programa ATIVO)
    sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
    sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_status]= 1 "
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
    ' Buscar o programa de fidelidade da empresa que esta ATIVO (só pode haver um programa ATIVO)
    sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
    sSql = sSql & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "' and [Cd_status]= 1 "
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 3 Then
    MsgBox "Este CNPJ NÃO esta definido em nenhum programa de fidelidade. Caso tenha criado UM programa de fidelidade a poucos instantes, logue novamente no quickStore para atualizar os dados.", vbInformation, "Atenção"
    Exit Sub
  End If

  Set rsPrograma = New ADODB.Recordset
  rsPrograma.Open sSql, gDB_SQLSERVER
  'Set rsPrograma = db_SQLSERVER.OpenRecordset(sSql, dbOpenDynaset, dbSeeChanges)
  
  If rsPrograma.BOF And rsPrograma.EOF Then
      ' Não existe um programa de fidelidade da empresa ATIVO
      rsPrograma.Close
      Set rsPrograma = Nothing
      
      MsgBox "Não existe programa ATIVO para este CNPJ", vbInformation, "Atenção"
      
      ' Fechar conexão com o banco de dados SQL SERVER
      'gnCloseDB_SQLSERVER

      Exit Sub
  End If
  lCodPrograma = rsPrograma.Fields("Cd_programa").Value
  sNomePrograma = rsPrograma.Fields("Nm_programa").Value
  dVl_ProgFidelidade = rsPrograma.Fields("Vl_ProgFidelidade").Value
  dVl_ProgFidelidadeParaCadaPonto = rsPrograma.Fields("Vl_ProgFidelidadeParaCadaPonto").Value
  dtValidadeInicio = rsPrograma.Fields("Dt_IniPrograma").Value
  dtValidadeFim = rsPrograma.Fields("Dt_FimPrograma").Value
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  If Now < dtValidadeInicio Or Now > dtValidadeFim Then
      'Lançamento não será realizado pois mesmo havendo um programa de fidelidade...ele esta fora do prazo

      ' Fechar conexão com o banco de dados SQL SERVER
      'gnCloseDB_SQLSERVER
      MsgBox "Prazo de validade para troca de pontos expirado!", vbInformation, "Atenção"
      cmd_resgatar.Enabled = False
      cmd_imprimir.Enabled = False
      cmd_receberResgate.Enabled = False

      Exit Sub
  End If
    
  txt_resgatarPontos.Text = ""
  txt_saldoReaisParaResgate.Text = ""
  
  Dim rsPrograma2 As New ADODB.Recordset
  Dim strSQL As String
  
  sCPF = txt_cpf.Text
  sCPF = Replace(sCPF, "-", "")
  sCPF = Replace(sCPF, ".", "")
  sCPF = Replace(sCPF, ";", "")
  sCPF = Replace(sCPF, "/", "")
  sCPF = Replace(sCPF, "\", "")
  
  If sCPF = "" Then
      sCPF = "_NADA_"
  End If
  
  strSQL = "SELECT CNPJ,CPF_CGC_CLIENTE,Cd_programa,Cd_cliente,Dt_criacao,Vl_CompraCliente, "
  strSQL = strSQL & " Nm_PontosAdquiridos,Vl_SaldoEmReais,Tp_lancamento,Cd_operador,Cd_guid_resgate,"
  strSQL = strSQL & " Status_guid_resgate,Dt_recebido_guid_resgate,Cd_operador_recebido_guid_resgate "
  strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos "
  'strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  strSQL = strSQL & " WHERE Cd_programa = " & lCodPrograma
  strSQL = strSQL & " AND CPF_CGC_CLIENTE = '" & sCPF & "' "
  strSQL = strSQL & " ORDER BY Dt_criacao ASC "

  'Set rsPrograma2 = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma2.Open strSQL, gDB_SQLSERVER

  Dim sSt_RESGATE As String
  Dim sTp_lancamento As String
  Dim lTotalPtsAcumulados As Long
  Dim dTotalReais As Double
  
  lTotalPtsAcumulados = 0
  dTotalReais = 0

  If Not (rsPrograma2.EOF And rsPrograma2.BOF) Then
    rsPrograma2.MoveFirst
  End If
  While Not rsPrograma2.EOF
  
      If rsPrograma2.Fields("Tp_lancamento").Value = 1 Then
          sTp_lancamento = "COMPRA"
          lTotalPtsAcumulados = lTotalPtsAcumulados + rsPrograma2.Fields("Nm_PontosAdquiridos").Value
          dTotalReais = dTotalReais + rsPrograma2.Fields("Vl_SaldoEmReais").Value
      Else
          sTp_lancamento = "RESGATE"
          lTotalPtsAcumulados = lTotalPtsAcumulados - rsPrograma2.Fields("Nm_PontosAdquiridos").Value
          dTotalReais = dTotalReais - rsPrograma2.Fields("Vl_SaldoEmReais").Value
      End If
      
      rsPrograma2.MoveNext
  Wend
  rsPrograma2.Close
  Set rsPrograma2 = Nothing
  
  txt_totalpontos.Text = lTotalPtsAcumulados
  txt_totalReais.Text = FormataValorTexto(dTotalReais, 2)
  
  txt_resgatarPontos.Text = lTotalPtsAcumulados
  txt_saldoReaisParaResgate.Text = txt_totalReais.Text
  
  Exit Sub
Erro:
  MsgBox "Erro metodo recuperarPontosFidelidade...Detalhes do Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
    
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Sub Form_Unload(Cancel As Integer)
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub

Private Sub txt_resgatarPontos_LostFocus()
    cmd_mostraSaldo_Click
End Sub
