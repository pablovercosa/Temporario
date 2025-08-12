VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmGrafico3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Vendas de um Produto Mês a Mês"
   ClientHeight    =   3585
   ClientLeft      =   3255
   ClientTop       =   2760
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Grafico3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   9270
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   90
      TabIndex        =   9
      Top             =   1515
      Width           =   9060
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
         Left            =   3465
         Picture         =   "Grafico3.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   307
         Width           =   465
      End
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
         Left            =   7785
         Picture         =   "Grafico3.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   307
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   6435
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         Height          =   315
         Left            =   2115
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data Final"
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
         Height          =   255
         Left            =   5445
         TabIndex        =   11
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data Inicial"
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
         Height          =   255
         Left            =   990
         TabIndex        =   10
         Top             =   390
         Width           =   1020
      End
   End
   Begin VB.Data Data2 
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
      Height          =   345
      Left            =   5850
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2970
      Width           =   9060
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
      Left            =   4995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   2430
      Visible         =   0   'False
      Width           =   1935
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto 
      Bindings        =   "Grafico3.frx":4FB1E
      DataSource      =   "Data2"
      Height          =   315
      Left            =   945
      TabIndex        =   1
      Top             =   840
      Width           =   2220
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
      Columns(0).Width=   8149
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3704
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3916
      _ExtentY        =   556
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
      Left            =   7740
      Top             =   2565
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "Grafico3.frx":4FB32
      DataSource      =   "Data1"
      Height          =   315
      Left            =   945
      TabIndex        =   0
      Top             =   270
      Width           =   2220
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
      _ExtentX        =   3916
      _ExtentY        =   556
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
   Begin VB.Label Nome_Produto 
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
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Produto"
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
      Left            =   120
      TabIndex        =   7
      Top             =   870
      Width           =   735
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
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   270
      Width           =   5895
   End
   Begin VB.Label Label7 
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
      Left            =   135
      TabIndex        =   5
      Top             =   300
      Width           =   600
   End
End
Attribute VB_Name = "frmGrafico3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEstoque As Recordset
Dim rsParametros As Recordset
Dim rsProdutos As Recordset
Dim rsTempo As Recordset

Public flagChamouDaTela_ProdutosMaisVendidos As Integer '1-chamou da tela Produtos mais vendidos; 0-chamou do menu
Public sCombo_Produto As String
Public sCombo_Filial As String
Public sData_ini As String
Public sData_fim As String

Private Sub B_Imprime_Click()
 Dim Data_Str As String
 Dim Mês As Integer
 Dim Mês1 As Integer
 Dim i As Long
 Dim Aux_Mês As Integer
 Dim Aux_Ano As Long
 Dim Aux_Produto As String
 Dim Aux_Contador As Double
 Dim Classe As Long
 Dim Erro As Integer
 Dim Unidades As Double
 Dim Valores As Double
 Dim sSql As String
 Dim Str_Rel As String
 Dim Str_Aux As String
 Dim Aux_Tamanho As Integer
 Dim Aux_Cor As Integer
 Dim Aux_Edição As Integer
 Dim Aux_Data As Date
 
 Call StatusMsg("")
 

 If Nome_Filial.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo_Filial.SetFocus
   Exit Sub
 End If

 If Nome_Produto.Caption = "" Then
   DisplayMsg "Escolha o produto."
   Combo_Produto.SetFocus
   Exit Sub
 End If
 
 If Not IsDate(Data_Ini.Text) Then
   DisplayMsg "Escolha um período de datas."
   Data_Ini.SetFocus
   Exit Sub
 End If

 If Not IsDate(Data_Fim.Text) Then
   DisplayMsg "Escolha um período de datas."
   Data_Fim.SetFocus
   Exit Sub
 End If
 
 sSql = "Delete * From Gráfico3"
 dbTemp.Execute sSql

 rsEstoque.Index = "Data"
 rsTempo.Index = "Mês"
 Aux_Mês = Mês
 Aux_Contador = 0
 Aux_Produto = Combo_Produto.Text
 Aux_Data = CDate(Data_Ini.Text)
 
Lp1:
 rsEstoque.Seek ">", Val(Combo_Filial.Text), Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Data
 If rsEstoque.NoMatch Then GoTo Imprime
 If rsEstoque("Filial") <> Val(Combo_Filial.Text) Then GoTo Imprime
 If rsEstoque("Produto") <> Aux_Produto Then GoTo Imprime
 If rsEstoque("Data") > CDate(Data_Fim.Text) Then GoTo Imprime
 
 Aux_Data = rsEstoque("Data")
 Aux_Tamanho = rsEstoque("Tamanho")
 Aux_Cor = rsEstoque("Cor")
 Aux_Edição = rsEstoque("Edição")
 
 rsTempo.Seek "=", Year(Aux_Data), Month(Aux_Data)
 If rsTempo.NoMatch Then
   rsTempo.AddNew
     rsTempo("Mês") = Month(Aux_Data)
     rsTempo("Ano") = Year(Aux_Data)
     rsTempo("Unidades Vendidas") = rsEstoque("Vendas")
     rsTempo("Valor Vendas") = rsEstoque("Valor Vendas")
     rsTempo("Nome") = str(rsTempo("Mês")) + " - " + str(rsTempo("Ano"))
   rsTempo.Update
 Else
   rsTempo.Edit
     rsTempo("Unidades Vendidas") = rsTempo("Unidades Vendidas") + rsEstoque("Vendas")
     rsTempo("Valor Vendas") = rsTempo("Valor Vendas") + rsEstoque("Valor Vendas")
   rsTempo.Update
 End If
     
 
 
 
 
 GoTo Lp1
 
 
 
Imprime:

 
  Rem Nome do arquivo .rpt
  Str_Rel = gsReportPath & "GRAFI3.RPT"
  Rel.ReportFileName = Str_Rel
 
  With Rel
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
  End With
 
  Rem Saída
  Rel.Destination = 0
 
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"

  Rel.Formulas(0) = Str_Rel

  Str_Rel = "nome_filial = '"
  Str_Rel = Str_Rel + Nome_Filial.Caption + "'"

  Rel.Formulas(1) = Str_Rel


  Str_Rel = "nome_produto = '"
  Str_Rel = Str_Rel + Combo_Produto.Text + " - " + Nome_Produto.Caption + "'"
  Rel.Formulas(2) = Str_Rel

  Rel.WindowState = crptMaximized
  Call StatusMsg("Aguarde, imprimindo...")
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  
  Rel.Action = 1
  
  Call StatusMsg("")

End Sub


Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
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


Private Sub Combo_Produto_CloseUp()
 Combo_Produto.Text = Combo_Produto.Columns(1).Text
 Combo_Produto_LostFocus

End Sub

Private Sub Combo_Produto_LostFocus()
  Nome_Produto.Caption = ""
  If IsNull(Combo_Produto.Text) Then Exit Sub
  If Combo_Produto.Text = "" Then Exit Sub
 
  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Combo_Produto.Text
  If rsProdutos.NoMatch Then Exit Sub
  Nome_Produto.Caption = rsProdutos("Nome")
  Combo_Produto.Text = UCase(Combo_Produto.Text)

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
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTempo = dbTemp.OpenRecordset("Gráfico3")
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  If flagChamouDaTela_ProdutosMaisVendidos = 1 Then
    Combo_Produto.Text = sCombo_Produto
    Combo_Produto_LostFocus
    Combo_Filial.Text = sCombo_Filial
    Combo_Filial_LostFocus
    Data_Ini.Text = sData_ini
    Data_Fim.Text = sData_fim
  End If
  
  Data_Fim.Text = Format(Now, "dd/mm/yyyy")
  Data_Ini.Text = Format(Now - 180, "dd/mm/yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsEstoque.Close
  rsProdutos.Close
  rsTempo.Close
  Set rsParametros = Nothing
  Set rsEstoque = Nothing
  Set rsProdutos = Nothing
  Set rsTempo = Nothing
End Sub
