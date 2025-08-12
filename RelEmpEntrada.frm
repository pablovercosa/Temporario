VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEmpEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relat�rio de Empr�stimos de Entrada"
   ClientHeight    =   3795
   ClientLeft      =   3570
   ClientTop       =   1965
   ClientWidth     =   7080
   Icon            =   "RelEmpEntrada.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   7080
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   1845
      TabIndex        =   19
      Top             =   2820
      Width           =   2640
      Begin VB.OptionButton O_Edi��o 
         Caption         =   "Edi��o"
         Height          =   225
         Left            =   1275
         TabIndex        =   9
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton O_Grade 
         Caption         =   "Grade"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   495
         Width           =   1065
      End
      Begin VB.OptionButton O_Normal 
         Caption         =   "Normal"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sa�da"
      Height          =   855
      Left            =   150
      TabIndex        =   18
      Top             =   2820
      Width           =   1575
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton B_V�deo 
         Caption         =   "V�deo"
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
      Left            =   3930
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Par�metro"
      Top             =   4695
      Visible         =   0   'False
      Width           =   1830
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelEmpEntrada.frx":058A
      DataSource      =   "Data3"
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   225
      Width           =   795
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
      _ExtentX        =   1402
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   6465
      Top             =   1560
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
      Top             =   3285
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Relat�rio"
      Height          =   975
      Left            =   150
      TabIndex        =   15
      Top             =   1740
      Width           =   4335
      Begin VB.OptionButton O_Simples 
         Caption         =   "Simples, somente a �ltima movimenta��o"
         Height          =   270
         Left            =   195
         TabIndex        =   4
         Top             =   600
         Width           =   3525
      End
      Begin VB.OptionButton O_Completo 
         Caption         =   "Completo (todas as movimenta��es)"
         Height          =   380
         Left            =   210
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2910
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2010
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   4665
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1740
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto 
      Bindings        =   "RelEmpEntrada.frx":059E
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1365
      TabIndex        =   2
      Top             =   1125
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
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2805
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelEmpEntrada.frx":05B2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1365
      TabIndex        =   1
      Top             =   690
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
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
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
      Left            =   3045
      TabIndex        =   17
      Top             =   210
      Width           =   3930
   End
   Begin VB.Label Label5 
      Caption         =   "Filial :"
      Height          =   225
      Left            =   255
      TabIndex        =   16
      Top             =   315
      Width           =   645
   End
   Begin VB.Label Nome_Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3045
      TabIndex        =   14
      Top             =   1125
      Width           =   3915
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3045
      TabIndex        =   13
      Top             =   690
      Width           =   3915
   End
   Begin VB.Label Label2 
      Caption         =   "Produto :"
      Height          =   225
      Left            =   225
      TabIndex        =   12
      Top             =   1185
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Fornecedor :"
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   735
      Width           =   960
   End
End
Attribute VB_Name = "frmRelEmpEntrada"
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
     DisplayMsg "Funcion�rio n�o tem acesso a esta filial."
     Exit Sub
   End If
 End If


 Rem Verifica fornecedor
 If Nome_Cliente.Caption = "" And Val(Combo_Cliente.Text) <> 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Fornecedor incorreto, verifique."
   Combo_Cliente.SetFocus
   Exit Sub
 End If


 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel1.DataFiles(0) = Str1

 Rem Sa�da
 If B_V�deo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Normal.Value = True Then
   If O_Completo.Value = True Then Str1 = gsReportPath & "EMENTRA1.RPT"
   If O_Simples.Value = True Then Str1 = gsReportPath & "EMENTRA2.RPT"
 End If
 If O_Grade.Value = True Then
   If O_Completo.Value = True Then Str1 = gsReportPath & "EMENTRA3.RPT"
   If O_Simples.Value = True Then Str1 = gsReportPath & "EMENTRA4.RPT"
 End If
 If O_Edi��o.Value = True Then
   If O_Completo.Value = True Then Str1 = gsReportPath & "EMENTRA5.RPT"
   If O_Simples.Value = True Then Str1 = gsReportPath & "EMENTRA6.RPT"
 End If
 
 
 
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 
 Str_Rel = "{Consigna��o Entrada.Filial} =" + Combo_Filial.Text
 If Nome_Cliente.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Consigna��o Entrada.Fornecedor} = " + Combo_Cliente.Text
 End If
 If Nome_Produto.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Consigna��o Entrada.Produto} = '" + Combo_Produto.Text + "'"
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
  'Seta a impressora para relat�rio
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
  
  rsClientes.Index = "C�digo"
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
  
  rsProdutos.Index = "C�digo"
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
 Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
 
End Sub

