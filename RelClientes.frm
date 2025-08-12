VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelCliFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Clientes e Fornecedores"
   ClientHeight    =   5985
   ClientLeft      =   1170
   ClientTop       =   795
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   8730
   Begin VB.Frame Frame10 
      Caption         =   "Quebra"
      Enabled         =   0   'False
      Height          =   885
      Left            =   5700
      TabIndex        =   38
      Top             =   3300
      Width           =   2925
      Begin VB.CheckBox Quebra_UF_Cidade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "UF / Cidade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
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
      Left            =   1965
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame9 
      Caption         =   "Imprimir somente deste vendedor"
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   4260
      Width           =   8505
      Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
         Bindings        =   "RelClientes.frx":4E95A
         DataSource      =   "Data2"
         Height          =   345
         Left            =   270
         TabIndex        =   22
         ToolTipText     =   "Deixe em branco ou use 0 para todos"
         Top             =   330
         Width           =   1065
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
         Columns(0).Width=   8916
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2064
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1879
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.Label Nome_Vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1380
         TabIndex        =   37
         Top             =   330
         Width           =   6945
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Ordem"
      Height          =   885
      Left            =   2520
      TabIndex        =   31
      Top             =   3300
      Width           =   3105
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   21
         Top             =   540
         Width           =   915
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   20
         Top             =   270
         Value           =   -1  'True
         Width           =   1065
      End
   End
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
      Left            =   195
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cli_For"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
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
      Height          =   525
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5250
      Width           =   8505
   End
   Begin VB.Frame Frame7 
      Caption         =   "Saída"
      Height          =   885
      Left            =   120
      TabIndex        =   30
      Top             =   3300
      Width           =   2325
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   18
         Top             =   270
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Opções"
      Height          =   1155
      Left            =   5700
      TabIndex        =   29
      Top             =   2040
      Width           =   2925
      Begin VB.CheckBox Contatos_Efetuados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Contatos &Efetuados"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   810
         Width           =   1905
      End
      Begin VB.CheckBox Contatos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Contatos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   540
         Width           =   1725
      End
      Begin VB.CheckBox Crédito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Inf. de Crédito"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   15
         Top             =   270
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Detalhes"
      Height          =   1365
      Left            =   5700
      TabIndex        =   26
      Top             =   540
      Width           =   2925
      Begin VB.OptionButton O_Simples 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Simples"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   990
         Width           =   1755
      End
      Begin VB.OptionButton O_P_Detalhado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Pouco detalhado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   1725
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   510
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton O_M_Detalhado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Muito detalhado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Estado"
      Height          =   1155
      Left            =   2520
      TabIndex        =   28
      Top             =   2040
      Width           =   3105
      Begin VB.TextBox Estado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   180
         MaxLength       =   2
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Preencha abaixo para imprimir somente de um estado."
         Height          =   495
         Left            =   180
         TabIndex        =   35
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cidade"
      Height          =   1365
      Left            =   2520
      TabIndex        =   25
      Top             =   540
      Width           =   3105
      Begin VB.TextBox Cidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   150
         MaxLength       =   30
         TabIndex        =   6
         Top             =   900
         Width           =   2835
      End
      Begin VB.Label Label4 
         Caption         =   "Preencha abaixo para imprimir somente os clientes/fornecedores de uma cidade."
         Height          =   675
         Left            =   150
         TabIndex        =   34
         Top             =   240
         Width           =   2865
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atividade"
      Height          =   1155
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Width           =   2325
      Begin VB.OptionButton O_Ativ_Inativ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   810
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton O_Inativos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Inativos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton O_Ativos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Ativos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   1785
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   2325
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1380
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton O_Outros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Outros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1110
         Width           =   1095
      End
      Begin VB.OptionButton O_Rev 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Revendedores"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   825
         Width           =   1455
      End
      Begin VB.OptionButton O_forn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Fornecedores"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   555
         Width           =   1335
      End
      Begin VB.OptionButton O_Cli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Clientes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   270
         Width           =   1035
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   120
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Cli 
      Bindings        =   "RelClientes.frx":4E96E
      DataSource      =   "Data1"
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   195
      Width           =   1350
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
      Columns.Count   =   3
      Columns(0).Width=   8334
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2170
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   1244
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   2381
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5700
      TabIndex        =   33
      Top             =   195
      Width           =   2925
   End
   Begin VB.Label Label2 
      Caption         =   "Imprimir somente este ->"
      Height          =   255
      Left            =   2550
      TabIndex        =   32
      Top             =   225
      Width           =   1785
   End
End
Attribute VB_Name = "frmRelCliFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCliFor As Recordset
Dim rsVendedores As Recordset

Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 
 
 Call StatusMsg("Aguarde...")
  
 Rel.Reset
 
 
  '24/09/2002 - mpdea
  'Configurado relatório para a exibição
  'do botão para setup da impressora
  Rel.WindowShowPrintSetupBtn = True

 
 If IsNull(Cidade.Text) Then Cidade.Text = ""
 If IsNull(Estado.Text) Then Estado.Text = ""

 
 Rem  Seta Valores e Manda Relatório


 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1
 
'05/10/2007 - Celso
'Incluido opção de quebra por UF/Cidade
 
 Rem Nome do arquivo .rpt
 If O_M_Detalhado.Value = True Then Str1 = gsReportPath & "CLI_FOR1"
 If O_Normal.Value = True Then Str1 = gsReportPath & "CLI_FOR2"
 If O_P_Detalhado.Value = True Then Str1 = gsReportPath & "CLI_FOR3"
 If O_Simples.Value = True Then Str1 = gsReportPath & "CLI_FOR4"
 
 If O_P_Detalhado.Value = True Or O_Simples.Value = True Then
    If Quebra_UF_Cidade.Value = 1 Then Str1 = Str1 & "_UF"
 End If
 
 Str1 = Str1 & ".RPT"
 
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Seleção
 If Nome_Cliente.Caption <> "" Then
   Str_Data1 = "{Cli_For.Código} = " + Combo_cli.Text
   GoTo Fim_Selec
 End If
   
 
 Str_Data1 = "{Cli_For.Código} <> -1" 'não serve para nada
 
 If O_Cli.Value = True Then Str_Data1 = Str_Data1 + " And {Cli_For.Tipo} = 'C'"
 If O_forn.Value = True Then Str_Data1 = Str_Data1 + " And {Cli_For.Tipo} = 'F'"
 If O_Rev.Value = True Then Str_Data1 = Str_Data1 + " And {Cli_For.Tipo} = 'R'"
 If O_Outros.Value = True Then Str_Data1 = Str_Data1 + " And {Cli_For.Tipo} = 'O'"

 If O_Ativos.Value = True Then Str_Data1 = Str_Data1 + " And {Cli_For.Inativo} = False"
 If O_Inativos.Value = True Then Str_Data1 = Str_Data1 + " And {Cli_For.Inativo} = True"

 If Cidade.Text <> "" Then
   Str_Data1 = Str_Data1 + " And {Cli_For.Cidade} = '" + Cidade.Text + "'"
 End If
 
 If Estado.Text <> "" Then
   Str_Data1 = Str_Data1 + " And {Cli_For.Estado} = '" + Estado.Text + "'"
 End If
 
 If Nome_Vendedor.Caption <> "" Then
   Str_Data1 = Str_Data1 + " And {Cli_For.Vendedor} = " + Combo_Vendedor.Text
 End If
 
Fim_Selec:
 Rel.SelectionFormula = Str_Data1
 
 If O_Código.Value = True Then Rel.SortFields(0) = "+{Cli_For.Código}"
 If O_Nome.Value = True Then Rel.SortFields(0) = "+{Cli_For.Nome}"
  
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 If Crédito.Value = 1 Then Rel.Formulas(1) = "imprime_crédito = 'SIM'"
 If Crédito.Value = 0 Then Rel.Formulas(1) = "imprime_crédito = 'NÃO'"
 
 If Contatos.Value = 1 Then Rel.Formulas(2) = "imprime_contatos = 'SIM'"
 If Contatos.Value = 0 Then Rel.Formulas(2) = "imprime_contatos = 'NÃO'"
 
 If Contatos_Efetuados.Value = 1 Then Rel.Formulas(3) = "imprime_contatos_efetuados = 'SIM'"
 If Contatos_Efetuados.Value = 0 Then Rel.Formulas(3) = "imprime_contatos_efetuados = 'NÃO'"
 
 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
 If O_M_Detalhado.Value = True Or O_Normal.Value = True Then
    Rel.SubreportToChange = "Crédito"
    Rel.DataFiles(0) = Str1
    Rel.SubreportToChange = "Contatos"
    Rel.DataFiles(0) = Str1
    Rel.SubreportToChange = "Contatos Efetuados"
    Rel.DataFiles(0) = Str1
 End If
 
 Rel.WindowState = crptMaximized
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
 
 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault
 
End Sub

Private Sub Combo_cli_CloseUp()
 Combo_cli.Text = Combo_cli.Columns(1).Text
 Combo_cli_LostFocus

End Sub

Private Sub Combo_cli_LostFocus()
 Nome_Cliente.Caption = ""
 
 If IsNull(Combo_cli.Text) Then Exit Sub
 If Not IsNumeric(Combo_cli.Text) Then Exit Sub
 If Val(Combo_cli.Text) <= 0 Then Exit Sub
 If Val(Combo_cli.Text) > 99999999 Then Exit Sub
 
 rsCliFor.Index = "Código"
 rsCliFor.Seek "=", Val(Combo_cli.Text)
 If rsCliFor.NoMatch Then Exit Sub
 
 Nome_Cliente.Caption = rsCliFor("Nome")
 

End Sub


Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(1).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()

  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Combo_Vendedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) < 1 Then Exit Sub
  If Val(Combo_Vendedor.Text) > 9999 Then Exit Sub
  
  rsVendedores.Index = "Código"
  rsVendedores.Seek "=", Val(Combo_Vendedor.Text)
  If rsVendedores.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsVendedores("Nome")
  
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
 Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
 Set rsVendedores = db.OpenRecordset("Funcionários", , dbReadOnly)
 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName
End Sub

Private Sub O_M_Detalhado_Click()
  Frame6.Enabled = True
  Frame10.Enabled = False
End Sub

Private Sub O_Normal_Click()
  Frame6.Enabled = True
  Frame10.Enabled = False
End Sub


Private Sub O_P_Detalhado_Click()
  Frame6.Enabled = False
  Frame10.Enabled = True
End Sub


Private Sub O_Simples_Click()
  Frame6.Enabled = False
  Frame10.Enabled = True
End Sub


