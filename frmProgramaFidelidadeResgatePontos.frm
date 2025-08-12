VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProgramaFidelidadeResgatePontos 
   Caption         =   " Programa Fidelidade x Resgate Pontos"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgramaFidelidadeResgatePontos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   14730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_mostraSaldo 
      BackColor       =   &H00C0FFC0&
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
      Left            =   11130
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6510
      Width           =   1545
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
      Left            =   9570
      TabIndex        =   21
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txt_totalReais 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   19
      Top             =   6930
      Width           =   1455
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   14595
   End
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
      TabIndex        =   8
      Top             =   7380
      Width           =   14595
   End
   Begin VB.TextBox txt_resgatarPontos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   9570
      TabIndex        =   6
      Top             =   6510
      Width           =   1455
   End
   Begin VB.TextBox txt_totalpontos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   16
      Top             =   6510
      Width           =   1455
   End
   Begin VB.CommandButton cmd_pesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
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
      TabIndex        =   4
      Top             =   960
      Width           =   14595
   End
   Begin VB.ComboBox cmb_programaFidelidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   510
      Width           =   5415
   End
   Begin VB.ComboBox cmb_clientes 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6870
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   90
      Width           =   7785
   End
   Begin VB.TextBox txt_cpf 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   75
      Width           =   2355
   End
   Begin MSFlexGridLib.MSFlexGrid grade_programas 
      Height          =   4365
      Left            =   60
      TabIndex        =   5
      Top             =   1770
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   7699
      _Version        =   393216
      Rows            =   1
      Cols            =   14
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   90
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label12 
      Caption         =   "--->"
      Height          =   225
      Left            =   14220
      TabIndex        =   23
      Top             =   6120
      Width           =   465
   End
   Begin VB.Label Label11 
      Caption         =   "--->"
      Height          =   225
      Left            =   14220
      TabIndex        =   22
      Top             =   1500
      Width           =   465
   End
   Begin VB.Label Label10 
      Caption         =   "Saldo em R$  para RESGATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   20
      Top             =   7020
      Width           =   2475
   End
   Begin VB.Label Label9 
      Caption         =   "Saldo em R$  acumulados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1890
      TabIndex        =   18
      Top             =   6990
      Width           =   2085
   End
   Begin VB.Label Label8 
      Caption         =   "Informar a qtde de pontos para RESGATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   17
      Top             =   6570
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Total de pontos acumulados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1890
      TabIndex        =   15
      Top             =   6570
      Width           =   2265
   End
   Begin VB.Label Label6 
      Caption         =   "Área de RESGATE DE PONTOS"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   60
      TabIndex        =   14
      Top             =   6180
      Width           =   3165
   End
   Begin VB.Label Label5 
      Caption         =   "ou"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   13
      Top             =   150
      Width           =   285
   End
   Begin VB.Label Label4 
      Caption         =   "Lançamentos no programa fidelidade do Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   12
      Top             =   1500
      Width           =   3765
   End
   Begin VB.Label Label2 
      Caption         =   "Programa Fidelidade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   570
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "Clientes Cód/Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5310
      TabIndex        =   10
      Top             =   135
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "CPF/CNPJ do Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1605
   End
End
Attribute VB_Name = "frmProgramaFidelidadeResgatePontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrPrograma(1000, 2) As String
Dim arrContador As Integer
Dim sNmCliente As String
Dim flagHAMUD As Boolean

Private Sub cmd_imprimir_Click()
On Error GoTo Erro:
  Dim strNomeArq As String
  Dim arrProg() As String
  Dim sEndereco As String
  Dim sBairro As String
  Dim sCidadeEstado As String
  Dim sFone As String
  Dim sNomeCliente As String
  
  If grade_programas.TextMatrix(grade_programas.RowSel, 8) <> "RESGATE" Then
    MsgBox "Selecione um registro de RESGATE na grade", vbInformation, "Atenção"
    Exit Sub
  End If
  
  If grade_programas.TextMatrix(grade_programas.RowSel, 11) = "Utilizado pelo Cliente" Then
    MsgBox "Este registro de RESGATE já foi usado e resgatado pelo cliente. Selecione outro que não foi usado ainda.", vbInformation, "Atenção"
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
  
  strNomeArq = gsReportPath & "programaFidelidade01.rpt"
  If Dir(strNomeArq) = "" Then
    DisplayMsg "Arquivo """ & strNomeArq & """ não encontrado."
    Exit Sub
  End If
  
  arrProg = Split(cmb_programaFidelidade.Text, " - ")

  CrystalReport1.DataFiles(0) = gsQuickDBFileName
  CrystalReport1.Destination = 0
  CrystalReport1.ReportFileName = strNomeArq
  CrystalReport1.ParameterFields(0) = "NomeEmpresa;" & gNomeEmpresaFilial & ";true"
  CrystalReport1.ParameterFields(1) = "EnderecoEmpresa;" & sEndereco & ";true"
  CrystalReport1.ParameterFields(2) = "NomeCliente;" & sNomeCliente & ";true"
  CrystalReport1.ParameterFields(3) = "CpfCnpjCliente;" & txt_cpf.Text & ";true"
  CrystalReport1.ParameterFields(4) = "NomeProgramaFidelidade;" & arrProg(1) & ";true"
  CrystalReport1.ParameterFields(5) = "DataResgate;" & grade_programas.TextMatrix(grade_programas.RowSel, 1) & ";true"
  CrystalReport1.ParameterFields(6) = "ValorResgate;" & FormataValorTexto(grade_programas.TextMatrix(grade_programas.RowSel, 7), 2) & ";true"
  CrystalReport1.ParameterFields(7) = "EnderecoEmpresa2;" & sBairro & ", " & sCidadeEstado & ", " & sFone & "" & ";true"
  CrystalReport1.ParameterFields(8) = "CodGuidResgate;" & grade_programas.TextMatrix(grade_programas.RowSel, 10) & ";true"
  CrystalReport1.WindowState = crptMaximized

  Call SetPrinterName("REL", CrystalReport1)

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
  Dim dSaldoPorPonto As Double
  Dim arrAux() As String
  Dim iCont As Integer
  
  
  If Len(LTrim(RTrim(txt_totalpontos.Text))) = 0 Then
      MsgBox "Campo 'Total de pontos acumulados' deve conter valor", vbInformation, "Atenção"
      Exit Sub
  Else
        lTotalPontosJaAdquiridos = CLng(txt_totalpontos.Text)
  End If
  
  dSaldoPorPonto = 0
  arrAux = Split(cmb_programaFidelidade.Text, " - ")
  For iCont = 0 To arrContador - 1
    If arrAux(0) = arrPrograma(iCont, 0) Then
          dSaldoPorPonto = arrPrograma(iCont, 1)
    End If
  Next
  
  If Len(LTrim(RTrim(txt_resgatarPontos.Text))) > 0 Then
      lPontos = CLng(txt_resgatarPontos.Text)
      dSaldo = lPontos * dSaldoPorPonto
      
      If lTotalPontosJaAdquiridos < lPontos Then
        MsgBox "RESGATE deve ser igual ou menor ao 'Total de pontos acumulados'", vbInformation, "Atenção"
        txt_saldoReaisParaResgate.Text = ""
        Exit Sub
      End If
  
      txt_saldoReaisParaResgate.Text = dSaldo
  Else
      txt_saldoReaisParaResgate.Text = 0
  End If

  Exit Sub
Erro:
  MsgBox "Entre com valores numéricos", vbInformation, "Atenção"

End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  Dim arrProg() As String
  Dim sCPF As String
  Dim sCNPJAux() As String

  txt_resgatarPontos.Text = ""
  txt_saldoReaisParaResgate.Text = ""

  If LTrim(RTrim(txt_cpf.Text)) = "" And cmb_clientes.Text = "" Then
      MsgBox "Selecione ou Informe o CPF/CNPJ do Cliente", vbInformation, "Atenção"
      txt_cpf.SetFocus
      Exit Sub
  End If

  If cmb_programaFidelidade.Text = "" Then
      MsgBox "Informe qual é o Programa de Fidelidade", vbInformation, "Atenção"
      cmb_programaFidelidade.SetFocus
      Exit Sub
  End If

  arrProg = Split(cmb_programaFidelidade.Text, " - ")

  grade_programas.Rows = 1

  If LTrim(RTrim(txt_cpf.Text)) <> "" Then
      sCPF = LTrim(RTrim(txt_cpf.Text))
  Else
      sCPF = LTrim(RTrim(cmb_clientes.Text))
      
      sCNPJAux = Split(sCPF, " *(")
      sCPF = sCNPJAux(1)
  End If

  
  sCPF = Replace(sCPF, "-", "")
  sCPF = Replace(sCPF, ".", "")
  sCPF = Replace(sCPF, ";", "")
  sCPF = Replace(sCPF, "/", "")
  sCPF = Replace(sCPF, "\", "")
  sCPF = Replace(sCPF, "(", "")
  sCPF = Replace(sCPF, ")", "")
  sCPF = Replace(sCPF, " ", "")
  
  If sCPF = "" Then
      MsgBox "Informe um Cliente válido *Com CPF/CNPJ", vbInformation, "Atenção"
      Exit Sub
  End If

  Dim rsPrograma As New ADODB.Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CNPJ,CPF_CGC_CLIENTE,Cd_programa,Cd_cliente,Dt_criacao,Vl_CompraCliente, "
  strSQL = strSQL & " Nm_PontosAdquiridos,Vl_SaldoEmReais,Tp_lancamento,Cd_operador,Cd_guid_resgate,"
  strSQL = strSQL & " Status_guid_resgate,Dt_recebido_guid_resgate,Cd_operador_recebido_guid_resgate "
  strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos "
  'strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  strSQL = strSQL & " WHERE Cd_programa = " & arrProg(0)
  strSQL = strSQL & " AND CPF_CGC_CLIENTE = '" & sCPF & "' "
  strSQL = strSQL & " ORDER BY Dt_criacao ASC "

  'Set rsPrograma = db_SQLSERVER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  rsPrograma.Open strSQL, gDB_SQLSERVER

  Dim sSt_RESGATE As String
  Dim sTp_lancamento As String
  Dim lTotalPtsAcumulados As Long
  Dim dTotalReais As Double
  
  lTotalPtsAcumulados = 0
  dTotalReais = 0

  If Not (rsPrograma.EOF And rsPrograma.BOF) Then
    rsPrograma.MoveFirst
  End If
  While Not rsPrograma.EOF
  
      If rsPrograma.Fields("Tp_lancamento").Value = 1 Then
          sTp_lancamento = "COMPRA"
          lTotalPtsAcumulados = lTotalPtsAcumulados + rsPrograma.Fields("Nm_PontosAdquiridos").Value
          dTotalReais = dTotalReais + rsPrograma.Fields("Vl_SaldoEmReais").Value
      Else
          sTp_lancamento = "RESGATE"
          lTotalPtsAcumulados = lTotalPtsAcumulados - rsPrograma.Fields("Nm_PontosAdquiridos").Value
          dTotalReais = dTotalReais - rsPrograma.Fields("Vl_SaldoEmReais").Value
          
          If Not IsNull(rsPrograma.Fields("Status_guid_resgate").Value) And rsPrograma.Fields("Status_guid_resgate").Value = 1 Then
              sSt_RESGATE = "Utilizado pelo Cliente"
          Else
              sSt_RESGATE = "Não utilizado pelo Cliente"
          End If
      End If
      
      grade_programas.AddItem 0 & vbTab & rsPrograma.Fields("Dt_criacao").Value & vbTab & _
                              rsPrograma.Fields("CNPJ").Value & vbTab & _
                              rsPrograma.Fields("CPF_CGC_CLIENTE").Value & vbTab & _
                              rsPrograma.Fields("Cd_cliente").Value & vbTab & _
                              Format(rsPrograma.Fields("Vl_CompraCliente").Value, FORMAT_VALUE) & vbTab & _
                              rsPrograma.Fields("Nm_PontosAdquiridos").Value & vbTab & _
                              Format(rsPrograma.Fields("Vl_SaldoEmReais").Value, FORMAT_VALUE) & vbTab & _
                              sTp_lancamento & vbTab & _
                              rsPrograma.Fields("Cd_operador").Value & vbTab & _
                              rsPrograma.Fields("Cd_guid_resgate").Value & vbTab & _
                              sSt_RESGATE & vbTab & _
                              rsPrograma.Fields("Dt_recebido_guid_resgate").Value & vbTab & _
                              rsPrograma.Fields("Cd_operador_recebido_guid_resgate").Value

      rsPrograma.MoveNext
  Wend
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  txt_totalpontos.Text = lTotalPtsAcumulados
  txt_totalReais.Text = FormataValorTexto(dTotalReais, 2)
  
  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da grade...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
    
End Sub

Private Function validaSaldoPontosNovamente() As Integer
On Error GoTo Erro
  Dim iStatus As Integer
  Dim sStatus As String
  Dim arrProg() As String
  Dim sCPF As String
  
  If LTrim(RTrim(txt_cpf.Text)) = "" And cmb_clientes.Text = "" Then
      MsgBox "Informe o CPF/CNPJ do Cliente", vbInformation, "Atenção"
      validaSaldoPontosNovamente = -1
      Exit Function
  End If
  
  If cmb_programaFidelidade.Text = "" Then
      MsgBox "Informe qual é o Programa de Fidelidade", vbInformation, "Atenção"
      validaSaldoPontosNovamente = -1
      Exit Function
  End If
  
  arrProg = Split(cmb_programaFidelidade.Text, " - ")
  
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
  strSQL = strSQL & " WHERE Cd_programa = " & arrProg(0)
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
  
  If LTrim(RTrim(txt_cpf.Text)) = "" And cmb_clientes.Text = "" Then
      MsgBox "Informe o CPF/CNPJ do Cliente", vbInformation, "Atenção"
      txt_cpf.SetFocus
      Exit Sub
  End If
  
  If cmb_programaFidelidade.Text = "" Then
      MsgBox "Informe qual é o Programa de Fidelidade", vbInformation, "Atenção"
      cmb_programaFidelidade.SetFocus
      Exit Sub
  End If
  
  If Len(LTrim(RTrim(txt_totalpontos.Text))) = 0 Or txt_resgatarPontos.Text = "" Or txt_resgatarPontos.Text = "0" Then
      MsgBox "Campo 'Total de pontos acumulados' deve conter valor", vbInformation, "Atenção"
      Exit Sub
  Else
        lTotalPontosJaAdquiridos = CLng(txt_totalpontos.Text)
  End If
  
  dSaldoPorPonto = 0
  arrAux = Split(cmb_programaFidelidade.Text, " - ")
  For iCont = 0 To arrContador - 1
    If arrAux(0) = arrPrograma(iCont, 0) Then
          dSaldoPorPonto = arrPrograma(iCont, 1)
    End If
  Next
  
  If Len(LTrim(RTrim(txt_resgatarPontos.Text))) > 0 Then
      Dim iRetorno As Integer

      iRetorno = validaSaldoPontosNovamente

      If iRetorno = -1 Then
          'MsgBox "Saldo insuficiente! RESGATE deve ser igual ou menor ao 'Total de pontos acumulados'", vbInformation, "Atenção"
          'txt_saldoReaisParaResgate.Text = ""
          Exit Sub
      End If

      lPontos = CLng(txt_resgatarPontos.Text)
      dSaldo = lPontos * dSaldoPorPonto
      
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
  
  iRet = ProgramaFidelidadeCriarLancamentoRESGATE(LTrim(RTrim(txt_cpf.Text)), arrAux(0), lPontos, dSaldo, sNmCliente)

  If iRet = 0 Then
      MsgBox "Resgate realizado com sucesso", vbInformation, "Sucesso"
  End If
  
  cmd_pesquisar_Click

  Exit Sub
Erro:
  MsgBox "Erro ao tentar realizar Resgate de Pontos..Detalhes do Erro: " & Err.Number & " " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
On Error GoTo Erro:
  Dim strSQL As String

  grade_programas.ColWidth(0) = 10
  grade_programas.ColWidth(1) = 1550
  grade_programas.ColWidth(2) = 1500
  grade_programas.ColWidth(3) = 1500
  grade_programas.ColWidth(4) = 1000
  grade_programas.ColWidth(5) = 1100
  grade_programas.ColWidth(6) = 1500
  grade_programas.ColWidth(7) = 1100
  grade_programas.ColWidth(8) = 1280
  grade_programas.ColWidth(9) = 1200
  grade_programas.ColWidth(10) = 2500
  grade_programas.ColWidth(11) = 2200
  grade_programas.ColWidth(12) = 1950
  grade_programas.ColWidth(13) = 1900

  grade_programas.Row = 0
  grade_programas.TextMatrix(0, 1) = "Dt Lançamento"
  grade_programas.TextMatrix(0, 2) = "CNPJ"
  grade_programas.TextMatrix(0, 3) = "CPF/CNPJ Cliente"
  grade_programas.TextMatrix(0, 4) = "Cód Cliente"
  grade_programas.TextMatrix(0, 5) = "Vl Compra"
  grade_programas.TextMatrix(0, 6) = "Pontos Adquiridos"
  grade_programas.TextMatrix(0, 7) = "Vl Ganho"
  grade_programas.TextMatrix(0, 8) = "Tp. lançamento"
  grade_programas.TextMatrix(0, 9) = "Cód.Operador"
  grade_programas.TextMatrix(0, 10) = "Cód RESGATE"
  grade_programas.TextMatrix(0, 11) = "Status RESGATE"
  grade_programas.TextMatrix(0, 12) = "Dt RESGATE Uso Cliente"
  grade_programas.TextMatrix(0, 13) = "Cód RESGATE Operador"
  
'  txt_cnpj.Text = gCNPJ_CPFControleDeLicencaWebApi
'  txt_nomeEmpresa.Text = gNomeEmpresaFilial
  
  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER
  
  Dim rsPrograma As New ADODB.Recordset
  Dim iStatus As Integer
  Dim sStatus As String
  
  arrContador = 0
  
  ProgramaFidelidadeEmpresaGrupoValida
  
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresa "
    strSQL = strSQL & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
    strSQL = "SELECT * FROM ProgramaFidelidade_empresa "
    strSQL = strSQL & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "'"
  Else
    MsgBox "É necessário ser um CNPJ Principal ou Participante de um Programa de Fidelidade para utilizar esta tela.", vbInformation, "Atenção"
  End If

  If gIndicadorProgramaFidelidadeCNPJPrincipal <> 3 Then
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
    
          cmb_programaFidelidade.AddItem rsPrograma.Fields("Cd_programa").Value & " - " & rsPrograma.Fields("Nm_programa").Value & " - " & sStatus
          arrPrograma(arrContador, 0) = rsPrograma.Fields("Cd_programa").Value
          arrPrograma(arrContador, 1) = Format(rsPrograma.Fields("Vl_ProgFidelidadeParaCadaPonto").Value, FORMAT_VALUE)
          
          arrContador = arrContador + 1
          rsPrograma.MoveNext
      Wend
      rsPrograma.Close
      Set rsPrograma = Nothing
  End If

  '***************************************************************
  'Verifica se HAMUD
  If gCNPJ_CPFControleDeLicencaWebApi = "80778855000187" _
      Or gCNPJ_CPFControleDeLicencaWebApi = "06888091000120" _
      Or gCNPJ_CPFControleDeLicencaWebApi = "08518307000190" _
      Or gCNPJ_CPFControleDeLicencaWebApi = "73213944000110" Then

    flagHAMUD = True
  End If
  '***************************************************************


  'Se HAMUD...então fazer um distinct por nomeCliente nos registros de lançamentos (pois são bases de clientes separados)
  If flagHAMUD = True Then
    Dim rsClientes As New ADODB.Recordset

    strSQL = "SELECT DISTINCT(Nm_cliente),CPF_CGC_CLIENTE, Cd_cliente "
    strSQL = strSQL & " FROM ProgramaFidelidade_lancamentos (nolock) "
    strSQL = strSQL & " WHERE CNPJ in('80778855000187','06888091000120','08518307000190','73213944000110')"

    rsClientes.Open strSQL, gDB_SQLSERVER

    cmb_clientes.AddItem ""

    'Carregar a combo de clientes
    While Not rsClientes.EOF
      cmb_clientes.AddItem rsClientes.Fields(2).Value & " - " & rsClientes.Fields(0).Value & " *(" & rsClientes.Fields(1).Value & ")"

      rsClientes.MoveNext
    Wend
    rsClientes.Close
    Set rsClientes = Nothing

  Else
    Dim rsClientes2 As Recordset
    Set rsClientes2 = db.OpenRecordset("Select Código, Nome, CGC from [Cli_For] order by Nome ", dbOpenDynaset)

    cmb_clientes.AddItem ""

    'Carregar a combo de clientes
    While Not rsClientes2.EOF
      cmb_clientes.AddItem rsClientes2.Fields(0).Value & " - " & rsClientes2.Fields(1).Value & " *(" & rsClientes2.Fields(2).Value & ")"

      rsClientes2.MoveNext
    Wend
    rsClientes2.Close
    Set rsClientes2 = Nothing
  End If

  Exit Sub
Erro:
  MsgBox "Erro ao realizar carga da tela...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Sub Form_Unload(Cancel As Integer)
  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER
End Sub

