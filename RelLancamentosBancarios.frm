VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelLancamentos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Lançamentos das Contas Correntes"
   ClientHeight    =   2805
   ClientLeft      =   1320
   ClientTop       =   1815
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1450
   Icon            =   "RelLancamentosBancarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2805
   ScaleWidth      =   7185
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   705
      Left            =   120
      TabIndex        =   9
      Top             =   540
      Width           =   6945
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   375
         Left            =   3990
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   375
         Left            =   1350
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3180
         TabIndex        =   11
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   465
         TabIndex        =   10
         Top             =   285
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   705
      Left            =   120
      TabIndex        =   8
      Top             =   1275
      Width           =   6945
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   300
         Width           =   1245
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1380
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2220
      Width           =   6945
   End
   Begin VB.Data Data4 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   3090
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6810
      Top             =   2580
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
      Bindings        =   "RelLancamentosBancarios.frx":4E95A
      DataSource      =   "Data4"
      Height          =   375
      Left            =   660
      TabIndex        =   0
      Top             =   90
      Width           =   1065
      DataFieldList   =   "Descrição"
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   7488
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2275
      Columns(1).Caption=   "Conta"
      Columns(1).Name =   "Conta"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Conta"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1640
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      _ExtentX        =   1879
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Nome_Conta 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1785
      TabIndex        =   7
      Top             =   90
      Width           =   5280
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Conta"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   435
   End
End
Attribute VB_Name = "frmRelLancamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLançamentos As Recordset
Dim rsContas As Recordset

Private Sub Combo_Conta_CloseUp()
Combo_Conta.Text = Combo_Conta.Columns(2).Text
Combo_Conta_LostFocus
End Sub

Private Sub B_Imprime_Click()
  Dim Str_Data1 As String
  Dim Str_Data2 As String
  Dim Str_Rel As String
  Dim Str1 As String
  Dim Erro As Integer
  

  Call StatusMsg("")

  If Nome_Conta.Caption = "" Then
    DisplayMsg "Conta incorreta, verifique."
    Combo_Conta.SetFocus
    Exit Sub
  End If
 
  Rem Verifica Data
  Erro = False
  If IsNull(Data_Ini.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  
  Rem Verifica Data Final
  Erro = False
  If IsNull(Data_Fim.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Data_Fim.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data incorreta, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  'Force atualizacao de saldo
  Call gnAtualizaSaldoBancario(Val(Combo_Conta.Text))

  Rem  Nome do BD
  Str1 = gsQuickDBFileName
  Rel.DataFiles(0) = Str1
  
  Rem Saída
  If B_Vídeo = True Then Rel.Destination = 0
  If B_Impressora = True Then Rel.Destination = 1
  
  Rem Nome do arquivo .rpt
  Str1 = gsReportPath & "LANCA.RPT"
  Rel.ReportFileName = Str1
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  Rel.Formulas(0) = Str_Rel
  
  Str_Rel = "nome_conta = '"
  Str_Rel = Str_Rel + Nome_Conta.Caption + "'"
  Rel.Formulas(1) = Str_Rel
  
  Rem data inicial
  Str_Rel = "data_ini = '"
  Str_Rel = Str_Rel + Data_Ini.Text + "'"
  Rel.Formulas(2) = Str_Rel
  
  Rem data final
  Str_Rel = "data_fim = '"
  Str_Rel = Str_Rel + Data_Fim.Text + "'"
  Rel.Formulas(3) = Str_Rel
  
  Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
  Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
  
  Str_Rel = "{Lançamentos Bancários.Conta} = "
  Str_Rel = Str_Rel + Combo_Conta.Text
  Str_Rel = Str_Rel + " And {Lançamentos Bancários.Data} >=" + Str_Data1
  Str_Rel = Str_Rel + " And {Lançamentos Bancários.Data} <=" + Str_Data2
  
  Rel.SelectionFormula = Str_Rel
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  
  Rel.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
End Sub

Private Sub Combo_Conta_LostFocus()
  Nome_Conta.Caption = ""
  If IsNull(Combo_Conta.Text) Then Exit Sub
  If Not IsNumeric(Combo_Conta.Text) Then Exit Sub
  If Val(Combo_Conta.Text) < 0 Or Val(Combo_Conta.Text) > 999999 Then Exit Sub

  rsContas.Index = "Código"
  rsContas.Seek "=", Val(Combo_Conta.Text)
  If rsContas.NoMatch Then Exit Sub
  Nome_Conta.Caption = rsContas("Descrição")

End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)

  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários", , dbReadOnly)
  
  Data4.DatabaseName = gsQuickDBFileName
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsContas.Close
  rsLançamentos.Close
  Set rsContas = Nothing
  Set rsLançamentos = Nothing
End Sub
