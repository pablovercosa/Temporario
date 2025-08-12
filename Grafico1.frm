VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmGrafico1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Comparação Mês a Mês - Vendas e Compras"
   ClientHeight    =   2775
   ClientLeft      =   3510
   ClientTop       =   2760
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Grafico1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   7785
   Begin VB.Data Data1 
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
      Height          =   345
      Left            =   5130
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   1170
      Visible         =   0   'False
      Width           =   2295
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
      Height          =   470
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2070
      Width           =   7500
   End
   Begin VB.TextBox Ano 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1290
      Width           =   1620
   End
   Begin VB.ComboBox Combo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Grafico1.frx":4E95A
      Left            =   1185
      List            =   "Grafico1.frx":4E982
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   735
      Width           =   1620
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "Grafico1.frx":4E9EB
      DataSource      =   "Data1"
      Height          =   360
      Left            =   1185
      TabIndex        =   0
      Top             =   240
      Width           =   1620
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
      Columns(0).Width=   9419
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
      _ExtentX        =   2857
      _ExtentY        =   635
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6210
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Nome_Filial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   4770
   End
   Begin VB.Label Label4 
      Caption         =   "Filial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   300
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Ano"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Mês Inicial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   795
      Width           =   975
   End
End
Attribute VB_Name = "frmGrafico1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset

Private Sub Ano_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
 
End Sub


Private Sub B_Imprime_Click()
 Dim Erro As Integer
 Dim Mês As Integer
 Dim Data_Str As String
 Dim Str_Rel As String
 
 Call StatusMsg("")


 If Nome_Filial.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo_Filial.SetFocus
   Exit Sub
 End If
 
 If Combo.Text = "" Then
   DisplayMsg "Escolha o mês."
   Combo.SetFocus
   Exit Sub
 End If
 
 
 Erro = False
 If IsNull(Ano.Text) Then Erro = True
 If Erro = False Then If Ano.Text = "" Then Erro = True
 If Erro = False Then If Not IsNumeric(Ano.Text) Then Erro = True
 If Erro = False Then If Val(Ano.Text) < 1994 Then Erro = True
 If Erro = False Then If Val(Ano.Text) > 2030 Then Erro = True
 
 If Erro = True Then
   DisplayMsg "Digite um ano entre 1994 e 2030, com 4 dígitos."
   Ano.SetFocus
   Exit Sub
 End If
 
 
 Mês = 0
 If Combo.Text = "Janeiro" Then Mês = 1
 If Combo.Text = "Fevereiro" Then Mês = 2
 If Combo.Text = "Março" Then Mês = 3
 If Combo.Text = "Abril" Then Mês = 4
 If Combo.Text = "Maio" Then Mês = 5
 If Combo.Text = "Junho" Then Mês = 6
 If Combo.Text = "Julho" Then Mês = 7
 If Combo.Text = "Agosto" Then Mês = 8
 If Combo.Text = "Setembro" Then Mês = 9
 If Combo.Text = "Outubro" Then Mês = 10
 If Combo.Text = "Novembro" Then Mês = 11
 If Combo.Text = "Dezembro" Then Mês = 12
 
 Data_Str = "01/" + Trim(str(Mês)) + "/" + Ano.Text
 
 
 Rem Nome do arquivo .rpt
 Str_Rel = gsReportPath & "GRAFI1.RPT"
 Rel.ReportFileName = Str_Rel
 
 Rel.DataFiles(0) = gsQuickDBFileName
 
 Rem Saída
 Rel.Destination = 0
 
 
 Str_Rel = "{Resumo Diário.Filial} = " + Combo_Filial.Text
 Str_Rel = Str_Rel + " And {Resumo Diário.Data} >= Date" + Format$(Data_Str, "(yyyy,mm,dd)")
 
 Rel.SelectionFormula = Str_Rel

 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Filial.Caption + "'"

 Rel.Formulas(1) = Str_Rel
 
 Rel.WindowState = crptMaximized
 Call StatusMsg("Aguarde, imprimindo...")
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
 
 Rel.Action = 1

 Call StatusMsg("")
 
End Sub


Private Sub Combo_Filial_CloseUp()

 Combo_Filial.Text = Combo_Filial.Columns(1).Text
 Combo_Filial_LostFocus

End Sub

Private Sub Combo_Filial_LostFocus()
  Nome_Filial.Caption = ""
  If IsNull(Combo_Filial.Text) Then Exit Sub
  If Combo_Filial.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Filial.Text) Then Exit Sub
  If Val(Combo_Filial.Text) < 0 Then Exit Sub
  If Val(Combo_Filial.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Filial.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Filial.Caption = rsParametros("Nome")


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Ano.Text = Year(Date)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Data1.DatabaseName = gsQuickDBFileName
End Sub


