VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendasCliente 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendas por Cliente/ Comprar por fornecedor"
   ClientHeight    =   3150
   ClientLeft      =   2115
   ClientTop       =   1980
   ClientWidth     =   7785
   ForeColor       =   &H80000008&
   HelpContextID   =   1520
   Icon            =   "RelVendasCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3150
   ScaleWidth      =   7785
   Begin VB.TextBox txtFilial 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   840
      Width           =   4335
   End
   Begin VB.Data datFilial 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "RelVendasCliente.frx":058A
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   1335
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
      Columns(0).Width=   4022
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Filial"
      Columns(0).FieldLen=   256
      Columns(1).Width=   5450
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Filial"
   End
   Begin VB.Frame Frame5 
      Caption         =   "Período"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5175
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordem"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   5175
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome Produto"
         Height          =   225
         Left            =   3600
         TabIndex        =   18
         Top             =   240
         Width           =   1365
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código do Produto"
         Height          =   225
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1680
      End
      Begin VB.OptionButton O_Data 
         Caption         =   "Data"
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton O_Edição 
         Caption         =   "Edição"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton O_Grade 
         Caption         =   "Grade"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton O_Normal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo do Relatório"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton O_Compras 
         Caption         =   "Compras por Fornecedor"
         Height          =   225
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton O_Vendas 
         Caption         =   "Vendas por Cliente"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   735
      Left            =   5400
      TabIndex        =   12
      Top             =   1560
      Width           =   2295
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   320
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   320
         Value           =   -1  'True
         Width           =   855
      End
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
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cli_For"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   5400
      Top             =   2520
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "RelVendasCliente.frx":05A2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
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
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1984
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      Caption         =   "Filial"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Cliente / Fornecedor :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Nome_Fornecedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      TabIndex        =   21
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "frmRelVendasCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsFornecedores As Recordset

Private Sub B_Cancela_Click()
End Sub

Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Str_Rel2 As String
 Dim Data1 As Variant
 
 
 Call StatusMsg("")


 Rem Verifica fornecedor
 Erro = False
 If IsNull(Combo_Fornecedor.Text) Then Erro = True
 If Not Erro Then If Not IsNumeric(Combo_Fornecedor.Text) Then Erro = True
 If Not Erro Then If Nome_Fornecedor.Caption = "" And Val(Combo_Fornecedor.Text) <> 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Cliente incorreto, verifique."
   Combo_Fornecedor.SetFocus
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

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1
 
 If O_Normal.Value = True Then Str1 = gsReportPath & "VENDA2.RPT"
 If O_Grade.Value = True Then Str1 = gsReportPath & "VENDA2G.RPT"
 If O_Edição.Value = True Then Str1 = gsReportPath & "VENDA2E.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 Rem Seleção
 If O_Vendas.Value = True Then Str_Rel = "{Resumo Clientes.Tipo} = 'C'"
 If O_Compras.Value = True Then Str_Rel = "{Resumo Clientes.Tipo} = 'F'"
 
 
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
 
 Str_Rel = Str_Rel + " And {Resumo Clientes.Dia} >=" + Str_Data1
 Str_Rel = Str_Rel + " And {Resumo Clientes.Dia} <=" + Str_Data2
 
 If (IsNumeric(cboFilial.Text)) Then
   If CInt(cboFilial.Text) <> 0 Then
     Str_Rel = Str_Rel & " And {Resumo Clientes.Filial} = " & cboFilial.Text
   End If
 End If
 
 Rel.SelectionFormula = Str_Rel
 
 Str_Rel2 = "imprime_total = 'S'"
 If Val(Combo_Fornecedor.Text) <> 0 Then
   Str_Rel = "{Resumo Clientes.Cliente} = " + Combo_Fornecedor.Text
   Rel.GroupSelectionFormula = Str_Rel
   Str_Rel2 = "imprime_total = 'N'"
 Else
   Rel.GroupSelectionFormula = ""
 End If
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 Rem data inicial
 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini.Text + "'"
 Rel.Formulas(1) = Str_Rel

 Rem data final
 Str_Rel = "data_fim = '"
 Str_Rel = Str_Rel + Data_Fim.Text + "'"
 Rel.Formulas(2) = Str_Rel

 Rel.Formulas(3) = Str_Rel2
 
 If O_Vendas.Value = True Then Str_Rel = "Titulo = 'Vendas por Cliente'"
 If O_Compras.Value = True Then Str_Rel = "Titulo = 'Compras por Fornecedor'"
 
 Rel.Formulas(4) = Str_Rel

 
 If O_Data.Value = True Then
   Rel.SortFields(0) = "+{Resumo Clientes.Cliente}"
   Rel.SortFields(1) = "+{Resumo Clientes.Dia}"
 End If
 If O_Código.Value = True Then
   Rel.SortFields(0) = "+{Resumo Clientes.Cliente}"
   Rel.SortFields(1) = "+{Produtos.Código Ordenação}"
 End If
 If O_Nome.Value = True Then
   Rel.SortFields(0) = "+{Resumo Clientes.Cliente}"
   Rel.SortFields(1) = "+{Produtos.Nome}"
 End If
 


 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
End Sub

Private Sub cboFilial_LostFocus()
  Dim rsFilial As Recordset
  
  txtFilial.Text = ""
  
  With cboFilial
    If Not IsNumeric(.Text) Then Exit Sub
    
    Set rsFilial = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
    
    If Not (rsFilial.BOF And rsFilial.EOF) Then
      txtFilial.Text = rsFilial.Fields("Nome") & ""
    End If
    
    rsFilial.Close
    Set rsFilial = Nothing
  End With
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Nome_Fornecedor.Caption = ""
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub

  rsFornecedores.Index = "Código"
  rsFornecedores.Seek "=", Combo_Fornecedor.Text
  If Not rsFornecedores.NoMatch Then
    Nome_Fornecedor.Caption = rsFornecedores("Nome")
  Else
    Combo_Fornecedor.Text = 0
  End If

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
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsFornecedores = db.OpenRecordset("Cli_For", , dbReadOnly)

  Data1.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  If gbGrade = False Then O_Grade.Enabled = False
  If gbEdicao = False Then O_Edição.Enabled = False
End Sub
