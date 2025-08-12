VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelAcompaProd 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Acompanhamento de Produto"
   ClientHeight    =   3480
   ClientLeft      =   1440
   ClientTop       =   1875
   ClientWidth     =   8700
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
   Icon            =   "RelAcompaProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   8700
   Begin VB.Frame Frame3 
      Caption         =   "Período"
      Height          =   795
      Left            =   105
      TabIndex        =   16
      Top             =   1065
      Width           =   8505
      Begin VB.CommandButton cmd_calendarioDtFim 
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
         Left            =   6840
         Picture         =   "RelAcompaProd.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   270
         Width           =   465
      End
      Begin VB.CommandButton cmd_calendarioDtIni 
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
         Left            =   2730
         Picture         =   "RelAcompaProd.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   270
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   345
         Left            =   5550
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
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
         Height          =   345
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   525
         TabIndex        =   18
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   330
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   4380
      TabIndex        =   15
      Top             =   1950
      Width           =   4215
      Begin VB.OptionButton O_Edição 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Edição"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton O_Grade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Grade"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3150
         TabIndex        =   7
         Top             =   360
         Width           =   840
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   540
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   105
      TabIndex        =   14
      Top             =   1950
      Width           =   4215
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   1245
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   690
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2910
      Width           =   8505
   End
   Begin VB.Data Data1 
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
      Left            =   780
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3255
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Data2"
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
      Left            =   3375
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   3315
      Visible         =   0   'False
      Width           =   2430
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6060
      Top             =   3300
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Prod 
      Bindings        =   "RelAcompaProd.frx":4FB1E
      DataSource      =   "Data2"
      Height          =   345
      Left            =   930
      TabIndex        =   1
      Top             =   615
      Width           =   1815
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8229
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3493
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3201
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelAcompaProd.frx":4FB32
      DataSource      =   "Data1"
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   150
      Width           =   1815
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9340
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1614
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   3201
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   13
      Top             =   195
      Width           =   405
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2805
      TabIndex        =   10
      Top             =   150
      Width           =   5805
   End
   Begin VB.Label Nome_Prod 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2805
      TabIndex        =   11
      Top             =   615
      Width           =   5805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Produto"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   105
      TabIndex        =   12
      Top             =   675
      Width           =   735
   End
End
Attribute VB_Name = "frmRelAcompaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsProdutos As Recordset
Dim rsResumo As Recordset
Dim rsParametros As Recordset
Dim rsEstoque As Recordset

Private Sub Ano_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub B_Imprime_Click()
  Dim Aux As Integer
  Dim Mês_Aux As Integer
  Dim Mês As Integer
  Dim Ano_Aux As Integer
  Dim Venda_Aux As Long
  Dim Nome_Aux As String
  Dim Termina As Integer
  Dim Produto As Double
  Dim Estoque As Double
  Dim Vendas_Aux As Double
  Dim Data_Aux As String
  Dim Str1 As String
  Dim Str_Rel As String
  Dim Str_Aux As String
  Dim sSql As String
  Dim Contador As Long
  Dim Str_Data1, Str_Data2 As String


  Call StatusMsg("")

 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a empresa."
   Combo.SetFocus
   Exit Sub
 End If

 If Filial_Liberada <> 0 Then
   If Val(Combo.Text) <> Filial_Liberada Then
     DisplayMsg "Funcionário não tem acesso a esta filial."
     Exit Sub
   End If
 End If

 If Nome_Prod.Caption = "" Then
    DisplayMsg "Escolha um produto."
    Combo_Prod.SetFocus
    Exit Sub
 End If


  If IsNull(Combo_Prod.Text) Then
    DisplayMsg "Escolha um produto."
    Combo_Prod.SetFocus
    Exit Sub
  End If
    
  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Data inválida."
    Data_Ini.SetFocus
    Exit Sub
  End If
 
  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Data inválida."
    Data_Fim.SetFocus
    Exit Sub
  End If

  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data final deve ser maior que data inicial."
    Data_Ini.SetFocus
    Exit Sub
  End If
  

  Call StatusMsg("Aguarde, imprimindo ...")

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Normal.Value = True Then Str1 = gsReportPath & "Acompa.RPT"
 If O_Grade.Value = True Then Str1 = gsReportPath & "AcompaG.RPT"
 If O_Edição.Value = True Then Str1 = gsReportPath & "AcompaE.RPT"
 
 Rel.ReportFileName = Str1

 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel
 
 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 
 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rel.Formulas(1) = Str_Rel
 

 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")


 Str_Rel = "{Estoque.Filial} =" + Combo.Text
 Str_Rel = Str_Rel + " And {Estoque.Produto} ='"
 Str_Rel = Str_Rel + Combo_Prod.Text + "'"
 Str_Rel = Str_Rel + " And {Estoque.Data} >=" + Str_Data1
 Str_Rel = Str_Rel + " And {Estoque.Data} <=" + Str_Data2
 Rel.SelectionFormula = Str_Rel

 
 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

  Exit Sub

End Sub

Private Sub cmd_calendarioDtFim_Click()
  Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
  Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub Combo_CloseUp()
Combo.Text = Combo.Columns(1).Text
Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Call StatusMsg("")
 
  Nome_Empresa.Caption = ""
  If IsNull(Combo.Text) Then Exit Sub
  If Combo.Text = "" Then Exit Sub
  If Not IsNumeric(Combo.Text) Then Exit Sub
  If Val(Combo.Text) < 0 Then Exit Sub
  If Val(Combo.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

End Sub

Private Sub Combo_Prod_CloseUp()

 Combo_Prod.Text = Combo_Prod.Columns(1).Text
 Combo_Prod_LostFocus

End Sub

Private Sub Combo_Prod_LostFocus()
  Call StatusMsg("")
 
  Nome_Prod.Caption = ""
  If IsNull(Combo_Prod.Text) Then Exit Sub
  If Combo_Prod.Text = "" Then Exit Sub
  If Combo_Prod.Text = "0" Then Exit Sub
  
  
  Combo_Prod.Text = UCase(Combo_Prod.Text)
  
  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Combo_Prod.Text

  If rsProdutos.NoMatch Then Exit Sub
  Nome_Prod.Caption = rsProdutos("Nome")

  If rsProdutos("Tipo") = "N" Then O_Normal.Value = True
  If rsProdutos("Tipo") = "G" Then O_Grade.Value = True
  If rsProdutos("Tipo") = "E" Then O_Edição.Value = True


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
  
 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)

 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName

 Combo.Text = gnCodFilial

 Data_Fim.Text = gsFormatDate(Date)

 O_Grade.Enabled = gbGrade
 O_Edição.Enabled = gbEdicao
 O_Normal.Enabled = True
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsProdutos.Close
  rsParametros.Close
  rsEstoque.Close
  Set rsProdutos = Nothing
  Set rsParametros = Nothing
  Set rsEstoque = Nothing
End Sub
