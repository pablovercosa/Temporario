VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelContaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório da Conta do Cliente"
   ClientHeight    =   2160
   ClientLeft      =   4125
   ClientTop       =   3255
   ClientWidth     =   5580
   Icon            =   "RelContas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2160
   ScaleWidth      =   5580
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Contas"
      Height          =   810
      Left            =   1590
      TabIndex        =   10
      Top             =   1215
      Width           =   2415
      Begin VB.OptionButton O_Pendentes 
         Caption         =   "Pendentes"
         Height          =   255
         Left            =   1110
         TabIndex        =   5
         Top             =   255
         Width           =   1215
      End
      Begin VB.OptionButton O_Pagas 
         Caption         =   "Pagas"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   975
      End
      Begin VB.OptionButton O_Todas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   150
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   4170
      TabIndex        =   6
      Top             =   1620
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelContas.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   120
      Width           =   945
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
      Columns(0).Width=   8943
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2196
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1667
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   4905
      Top             =   495
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
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   3690
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   8
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmRelContaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsCliFor As Recordset

Private Sub B_Imprime_Click()
  Dim Str_Rel As String
  Call StatusMsg("")
  
  'Verifica fornecedor
  If Nome_Cliente.Caption = "" And Val(Combo_Cliente.Text) <> 0 Then
    DisplayMsg "Cliente incorreto, verifique."
    Combo_Cliente.SetFocus
    Exit Sub
  End If
  
  'Nome do BD
  Rel1.DataFiles(0) = gsQuickDBFileName
  Rel1.DataFiles(1) = gsQuickDBFileName
  Rel1.DataFiles(2) = gsQuickDBFileName
  Rel1.DataFiles(3) = gsQuickDBFileName
  
  'Saída
  If B_Vídeo.Value Then
    Rel1.Destination = crptToWindow
  Else
    Rel1.Destination = crptToPrinter
  End If
  
  'Nome do arquivo .rpt
  Rel1.ReportFileName = gsReportPath & "CONTA1.RPT"
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel1
  
  'Seleção
  Str_Rel = ""
  If O_Pagas.Value Then
    Str_Rel = "{Conta Cliente.Valor} = {Conta Cliente.Valor Pago}"
  ElseIf O_Pendentes.Value Then
    Str_Rel = "{Conta Cliente.Valor} > {Conta Cliente.Valor Pago}"
  End If
  
  '  ( ([Conta Cliente].Valor)> Round([Valor Pago],2));
  
  If Nome_Cliente.Caption <> "" Then
    If Str_Rel = "" Then
      Str_Rel = "{Conta Cliente.Cliente} = " & Combo_Cliente.Text
    Else
      Str_Rel = Str_Rel & " And {Conta Cliente.Cliente} = " & Combo_Cliente.Text
    End If
  End If
  
  Rel1.SelectionFormula = Str_Rel
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  
  Rel1.Formulas(0) = Str_Rel
  
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
  
  Rel1.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault


End Sub

Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Cliente_LostFocus()
  Call StatusMsg("")
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub

  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Combo_Cliente.Text
  If Not rsCliFor.NoMatch Then
    Nome_Cliente.Caption = rsCliFor("Nome")
  Else
    Combo_Cliente.Text = 0
  End If

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Data2.DatabaseName = gsQuickDBFileName
End Sub
