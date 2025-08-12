VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEmpSaida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Empréstimos de Saída"
   ClientHeight    =   3630
   ClientLeft      =   3570
   ClientTop       =   1965
   ClientWidth     =   7050
   Icon            =   "RelEmpSaida.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   7050
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   1875
      TabIndex        =   19
      Top             =   2640
      Width           =   2550
      Begin VB.OptionButton O_Edição 
         Caption         =   "Edição"
         Height          =   225
         Left            =   1230
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton O_Grade 
         Caption         =   "Grade"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   540
         Width           =   1065
      End
      Begin VB.OptionButton O_Normal 
         Caption         =   "Normal"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      Height          =   855
      Left            =   150
      TabIndex        =   18
      Top             =   2640
      Width           =   1575
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   555
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4710
      Visible         =   0   'False
      Width           =   1830
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelEmpSaida.frx":058A
      DataSource      =   "Data3"
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   135
      Width           =   1065
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
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
      Columns(0).Width=   9049
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2064
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   6480
      Top             =   1545
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
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5610
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Relatório"
      Height          =   1005
      Left            =   150
      TabIndex        =   15
      Top             =   1530
      Width           =   4275
      Begin VB.OptionButton O_Simples 
         Caption         =   "Simples, somente a última movimentação"
         Height          =   360
         Left            =   150
         TabIndex        =   4
         Top             =   555
         Width           =   3570
      End
      Begin VB.OptionButton O_Completo 
         Caption         =   "Completo (todas as movimentações)"
         Height          =   380
         Left            =   165
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   3690
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   4710
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   4695
      Visible         =   0   'False
      Width           =   1710
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto 
      Bindings        =   "RelEmpSaida.frx":059E
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Top             =   1065
      Width           =   1590
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
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
      Columns(0).Width=   7699
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3836
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2805
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelEmpSaida.frx":05B2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   630
      Width           =   1065
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
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
      Columns(0).Width=   9102
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1879
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2730
      TabIndex        =   17
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Filial :"
      Height          =   225
      Left            =   210
      TabIndex        =   16
      Top             =   210
      Width           =   645
   End
   Begin VB.Label Nome_Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2730
      TabIndex        =   14
      Top             =   1050
      Width           =   4215
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2730
      TabIndex        =   13
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Produto :"
      Height          =   225
      Left            =   210
      TabIndex        =   12
      Top             =   1110
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   225
      Left            =   225
      TabIndex        =   11
      Top             =   690
      Width           =   645
   End
End
Attribute VB_Name = "frmRelEmpSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim rsClientes As Recordset
Dim rsProdutos As Recordset
Dim rsParametros As Recordset

Private Sub B_Imprime_Click()

 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 
 
 Call StatusMsg("")

 Rem Verifica empresa
 If IsNull(Nome_Filial.Caption) Or Nome_Filial.Caption = "" Then
   DisplayMsg "Escolha a empresa."
   Combo_Filial.SetFocus
   Exit Sub
 End If

 If Filial_Liberada <> 0 Then
   If Val(Combo_Filial.Text) <> Filial_Liberada Then
     DisplayMsg "Funcionário não tem acesso a esta filial."
     Exit Sub
   End If
 End If


 Rem Verifica fornecedor
 If Nome_Cliente.Caption = "" And Val(Combo_Cliente.Text) <> 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Cliente incorreto, verifique."
   Combo_Cliente.SetFocus
   Exit Sub
 End If

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel1.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Normal.Value = True Then
   If O_Completo.Value = True Then Str1 = gsReportPath & "EMSAIDA1.RPT"
   If O_Simples.Value = True Then Str1 = gsReportPath & "EMSAIDA2.RPT"
 End If
 If O_Grade.Value = True Then
   If O_Completo.Value = True Then Str1 = gsReportPath & "EMSAIDA3.RPT"
   If O_Simples.Value = True Then Str1 = gsReportPath & "EMSAIDA4.RPT"
 End If
 If O_Edição.Value = True Then
   If O_Completo.Value = True Then Str1 = gsReportPath & "EMSAIDA5.RPT"
   If O_Simples.Value = True Then Str1 = gsReportPath & "EMSAIDA6.RPT"
 End If
 
 
 
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 
 Str_Rel = "{Consignação Saída.Filial} =" + Combo_Filial.Text
 If Nome_Cliente.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Consignação Saída.Cliente} = " + Combo_Cliente.Text
 End If
 If Nome_Produto.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Consignação Saída.Produto} = '" + Combo_Produto.Text + "'"
 End If
 

 Rel1.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel1.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Filial.Caption + "'"
 Rel1.Formulas(1) = Str_Rel


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
  If Combo_Cliente.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 1 Then Exit Sub
  If Val(Combo_Cliente.Text) > 99999999 Then Exit Sub
  
  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  
  Nome_Cliente.Caption = rsClientes("Nome") & ""
  
  

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
 If Val(Combo_Filial.Text) < 1 Or Val(Combo_Filial.Text) > 99 Then Exit Sub
 rsParametros.Index = "Filial"
 rsParametros.Seek "=", Val(Combo_Filial.Text)
 If rsParametros.NoMatch Then Exit Sub
 Nome_Filial.Caption = rsParametros("Nome") & ""
 
End Sub

Private Sub Combo_Produto_CloseUp()

 Combo_Produto.Text = Combo_Produto.Columns(1).Text
 Combo_Produto_LostFocus
 
 
End Sub

Private Sub Combo_Produto_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Produto_LostFocus()
  
  Call StatusMsg("")
  Nome_Produto.Caption = ""
  If IsNull(Combo_Produto.Text) Then Exit Sub
  If Combo_Produto.Text = "" Then Exit Sub
  If Combo_Produto.Text = "0" Then Exit Sub
  
  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Combo_Produto.Text
  If rsProdutos.NoMatch Then Exit Sub
  
  Nome_Produto.Caption = rsProdutos("Nome") & ""
  
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)


 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName
 Data3.DatabaseName = gsQuickDBFileName
 
 Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 
End Sub

