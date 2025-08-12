VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelPagas1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Contas a Pagar ""Já pagas"" por Centro de Custo"
   ClientHeight    =   2760
   ClientLeft      =   1800
   ClientTop       =   2250
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelPagas1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2760
   ScaleWidth      =   7605
   Begin VB.Frame fraX 
      Height          =   1095
      Left            =   60
      TabIndex        =   12
      Top             =   30
      Width           =   7485
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "RelPagas1.frx":4E95A
         DataSource      =   "datPara"
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   210
         Width           =   1005
         DataFieldList   =   "Filial"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   7752
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1879
         Columns(1).Caption=   "Filial"
         Columns(1).Name =   "Filial"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Filial"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCentro 
         Bindings        =   "RelPagas1.frx":4E970
         DataSource      =   "datCC"
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   630
         Width           =   1005
         DataFieldList   =   "Nome"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   9922
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1746
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblNomeCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   16
         Top             =   630
         Width           =   4860
      End
      Begin VB.Label lblCC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label lblNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   210
         Width           =   4860
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      Height          =   975
      Left            =   60
      TabIndex        =   10
      Top             =   1200
      Width           =   4275
      Begin VB.OptionButton optPeriodoVencimento 
         Appearance      =   0  'Flat
         Caption         =   "Vencimento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3060
         TabIndex        =   4
         Top             =   660
         Width           =   1125
      End
      Begin VB.OptionButton optPeriodoPagamento 
         Appearance      =   0  'Flat
         Caption         =   "Pagamento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3060
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1125
      End
      Begin MSMask.MaskEdBox mskDataIni 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataFim 
         Height          =   315
         Left            =   1740
         TabIndex        =   3
         ToolTipText     =   "Ao teclar [F2] carrega calendário"
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1530
         TabIndex        =   11
         Top             =   300
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar relatório"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2250
      Width           =   7485
   End
   Begin VB.Frame Frame4 
      Caption         =   "Saída"
      Height          =   975
      Left            =   4380
      TabIndex        =   9
      Top             =   1200
      Width           =   3135
      Begin VB.OptionButton optImpressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1650
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optVideo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   330
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Data datCC 
      Caption         =   "datCC"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Data datPara 
      Caption         =   "datPara"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   4800
      Width           =   2055
   End
   Begin Crystal.CrystalReport crpRel 
      Left            =   4680
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRelPagas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'21/10/2008 - mpdea
'Ajustes gerais e inclusão de escolha se o período é por data de vencimento ou de pagamento

'05/05/2005 - Daniel
'
'Projeto: Melhorias para o Centro de Custo
'
'A partir da versão 6.52.0.38 todo campo de Centro de Custo
'estará carregando apenas os Centros que estão ativos no sistema

Dim rsParametros As Recordset
Dim rsCentros As Recordset

Private Sub cmdFechar_Click()
  Unload Me
End Sub

'21/10/2008 - mpdea
'Incluído opção para selecionar por Data de Vencimento ou Data de Pagamento
Private Sub cmdImprimir_Click()
  Dim Erro As Integer
  Dim Str_Data1 As String, Str_Data2 As String
  Dim Str_Rel As String
  '21/10/2008 - mpdea
  Dim str_campo_periodo As String

  Call StatusMsg("")
  
  Rem Verifica empresa
  If IsNull(lblNomeFilial.Caption) Or lblNomeFilial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    cboFilial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  
  Rem Verifica Data
  Erro = False
  If IsNull(mskDataIni.Text) Then Erro = True
  If Not Erro Then If Not IsDate(mskDataIni.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    mskDataIni.SetFocus
    Exit Sub
  End If
  
  Rem Verifica Data Final
  Erro = False
  If IsNull(mskDataFim.Text) Then Erro = True
  If Not Erro Then If Not IsDate(mskDataFim.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    mskDataFim.SetFocus
    Exit Sub
  End If
  
  
  If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    mskDataIni.SetFocus
    Exit Sub
  End If

  '21/10/2008 - mpdea
  If optPeriodoVencimento.Value Then
    str_campo_periodo = "Vencimento"
  Else
    str_campo_periodo = "Pagamento"
  End If

  'Nome do BD
  crpRel.DataFiles(0) = gsQuickDBFileName

  'Saída
  If optVideo.Value Then
    crpRel.Destination = 0
  Else
    crpRel.Destination = 1
  End If
  
  'Nome do arquivo .rpt
  crpRel.ReportFileName = gsReportPath & "PAGAS1.RPT"
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 crpRel

  Rem Seleção
  Str_Data1 = "Date" + Format$(mskDataIni.Text, "(yyyy,mm,dd)")
  Str_Data2 = "Date" + Format$(mskDataFim.Text, "(yyyy,mm,dd)")
  
  Str_Rel = "{Contas a Pagar.Filial} =" + cboFilial.Text
  Str_Rel = Str_Rel + " And {Contas a Pagar." & str_campo_periodo & "} >="
  Str_Rel = Str_Rel + Str_Data1
  Str_Rel = Str_Rel + " And {Contas a Pagar." & str_campo_periodo & "} <=" + Str_Data2
  Str_Rel = Str_Rel + " And {Contas a Pagar.Valor Pago} <> 0"
  
  If lblNomeCC.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Contas a Pagar.Centro de Custo} = " + str(cboCentro.Text)
  End If
  
  crpRel.SelectionFormula = Str_Rel
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  
  crpRel.Formulas(0) = Str_Rel
  
  Str_Rel = "nome_filial = '"
  Str_Rel = Str_Rel + lblNomeFilial.Caption + "'"
  crpRel.Formulas(1) = Str_Rel
  
  Rem data inicial
  Str_Rel = "data_ini = '"
  Str_Rel = Str_Rel + mskDataIni.Text + "'"
  crpRel.Formulas(2) = Str_Rel
  
  Rem data final
  Str_Rel = "data_fim = '"
  Str_Rel = Str_Rel + mskDataFim.Text + "'"
  crpRel.Formulas(3) = Str_Rel
  
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", crpRel)
  
  
  crpRel.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault

End Sub

Private Sub cboCentro_CloseUp()
  cboCentro.Text = cboCentro.Columns(1).Text
  cboCentro_LostFocus
End Sub

Private Sub cboCentro_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboCentro_LostFocus()
  Call StatusMsg("")
 
  lblNomeCC.Caption = ""
  If IsNull(cboCentro.Text) Then Exit Sub
  If cboCentro.Text = "" Then Exit Sub
  If Not IsNumeric(cboCentro.Text) Then Exit Sub
  If Val(cboCentro.Text) < 0 Then Exit Sub
  If Val(cboCentro.Text) > 9999 Then Exit Sub

  rsCentros.Index = "Código"
  rsCentros.Seek "=", Val(cboCentro.Text)
  If rsCentros.NoMatch Then Exit Sub
  lblNomeCC.Caption = rsCentros("Nome")

End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(1).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Call StatusMsg("")
 
  lblNomeFilial.Caption = ""
  If IsNull(cboFilial.Text) Then Exit Sub
  If cboFilial.Text = "" Then Exit Sub
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  If Val(cboFilial.Text) < 0 Then Exit Sub
  If Val(cboFilial.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(cboFilial.Text)
  If rsParametros.NoMatch Then Exit Sub
  lblNomeFilial.Caption = rsParametros("Nome")

End Sub

Private Sub mskDataIni_LostFocus()
  mskDataIni.Text = Ajusta_Data(mskDataIni.Text)
End Sub

Private Sub mskDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      mskDataIni.Text = frmCalendario.gsDateCalender(mskDataIni.Text)
  End Select
End Sub

Private Sub mskDataFim_LostFocus()
  mskDataFim.Text = Ajusta_Data(mskDataFim.Text)
End Sub

Private Sub mskDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      mskDataFim.Text = frmCalendario.gsDateCalender(mskDataFim.Text)
  End Select
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)

  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsCentros = db.OpenRecordset("Centros de Custo", , dbReadOnly)
  
  datPara.DatabaseName = gsQuickDBFileName
  datCC.DatabaseName = gsQuickDBFileName
  
  cboFilial.Text = gnCodFilial

End Sub

Private Sub Form_Unload(Cancel As Integer)

 rsParametros.Close
 rsCentros.Close
 Set rsParametros = Nothing
 Set rsCentros = Nothing


End Sub
