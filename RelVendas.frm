VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelVendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas"
   ClientHeight    =   4995
   ClientLeft      =   2070
   ClientTop       =   2175
   ClientWidth     =   7125
   Icon            =   "RelVendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4995
   ScaleWidth      =   7125
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox chkShowDescSubTotal 
      Caption         =   "Exibir totalizadores para Desconto no SubTotal da venda (sem rateio)"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Opções"
      Height          =   930
      Left            =   105
      TabIndex        =   27
      Top             =   2130
      Width           =   4095
      Begin VB.CheckBox O_Classe 
         Caption         =   "Separar por classe e sub-classe"
         Height          =   225
         Left            =   270
         TabIndex        =   4
         Top             =   285
         Width           =   3060
      End
      Begin VB.CheckBox O_Vendas_Zero 
         Caption         =   "Não considerar produtos sem vendas"
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   795
      Left            =   120
      TabIndex        =   24
      Top             =   1260
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3690
         TabIndex        =   3
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
         Left            =   1080
         TabIndex        =   2
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   25
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2055
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   5745
      Visible         =   0   'False
      Width           =   1905
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "RelVendas.frx":0CCA
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   750
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
      Columns(0).Width=   8864
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2487
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
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   870
      Left            =   4440
      TabIndex        =   21
      Top             =   2160
      Width           =   2520
      Begin VB.OptionButton O_Edição 
         Caption         =   "Com Edição"
         Height          =   225
         Left            =   1140
         TabIndex        =   13
         Top             =   495
         Width           =   1275
      End
      Begin VB.OptionButton O_Grade 
         Caption         =   "Com Grade"
         Height          =   225
         Left            =   1140
         TabIndex        =   12
         Top             =   195
         Width           =   1275
      End
      Begin VB.OptionButton O_Normal 
         Caption         =   "Normal"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   840
      Left            =   105
      TabIndex        =   20
      Top             =   3135
      Width           =   4110
      Begin VB.OptionButton O_Valores 
         Caption         =   "Ranking por valores"
         Height          =   225
         Left            =   1290
         TabIndex        =   9
         Top             =   510
         Width           =   1905
      End
      Begin VB.OptionButton O_Unidades 
         Caption         =   "Ranking por unidades"
         Height          =   225
         Left            =   1290
         TabIndex        =   8
         Top             =   225
         Width           =   2010
      End
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   510
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   5520
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   510
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   150
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   5730
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5640
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   5670
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelVendas.frx":0CDE
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1275
      TabIndex        =   0
      Top             =   240
      Width           =   720
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
      Columns(0).Width=   3200
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Fornecedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2460
      TabIndex        =   23
      Top             =   735
      Width           =   4485
   End
   Begin VB.Label Label4 
      Caption         =   "Fornecedor :"
      Height          =   225
      Left            =   105
      TabIndex        =   22
      Top             =   825
      Width           =   960
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2460
      TabIndex        =   18
      Top             =   240
      Width           =   4485
   End
   Begin VB.Label Label1 
      Caption         =   "Filial :"
      Height          =   255
      Left            =   135
      TabIndex        =   17
      Top             =   330
      Width           =   495
   End
End
Attribute VB_Name = "frmRelVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsParametros As Recordset
Dim rsFornecedores As Recordset
Dim rsForn_Prod As Recordset

Private Sub B_Imprime_Click()
 Dim Val1 As Integer
 Dim Val2 As Integer
 Dim Erro As Integer
 Dim Str1 As String
 Dim Str2 As String
 Dim Str3 As String
 Dim Str_Data1 As String
 Dim Str_Data2 As String
 Dim Str_Rel As String
 Dim Aux_Data1 As Variant
 Dim rsVendas As Recordset
 Dim rsEstoque As Recordset
 Dim Aux_Data As Date
 Dim Aux_Prod As String
 Dim Aux_Cor As Integer
 Dim Aux_Cor2 As Integer
 Dim Aux_Tamanho As Integer
 Dim Aux_Tamanho2 As Integer
 Dim Aux_Edição As Long
 Dim Aux_Edição2 As Long

  Dim strSQL As String
  Dim rsDescSubTotal As Recordset
  Dim curDescSubTotal As Currency
 
 Call StatusMsg("")
 
 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a Filial."
   Combo_Filial.SetFocus
   Exit Sub
 End If
 
  If Filial_Liberada <> 0 Then
    If Val(Combo_Filial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If

 Erro = False
 If IsNull(Data_Ini.Text) Then Erro = True
 If Erro = False Then If Not IsDate(Data_Ini.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data inválida, verifique."
   Data_Ini.SetFocus
   Exit Sub
 End If
 
 Erro = False
 If IsNull(Data_Fim.Text) Then Erro = True
 If Erro = False Then If Not IsDate(Data_Fim.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data inválida, verifique."
   Data_Fim.SetFocus
   Exit Sub
 End If
 
 Data_Ini.Text = gsFormatDate(Data_Ini.Text)
 Data_Fim.Text = gsFormatDate(Data_Fim.Text)
 
 If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
   DisplayMsg "Data final menor que data inicial, verifique."
   Data_Fim.SetFocus
   Exit Sub
 End If
 
 Rem Abre, Zera e Gera arquivo de vendas
 Call StatusMsg("Aguarde, preparando relatório...")
 Set rsVendas = db.OpenRecordset("ZZZVendas")
 Set rsEstoque = db.OpenRecordset("Estoque")
 db.Execute "Delete * From ZZZVendas"
 
 rsEstoque.Index = "Produto"
 rsVendas.Index = "Produto"
 rsForn_Prod.Index = "Produto"
 Aux_Prod = ""
 Aux_Tamanho = 0
 Aux_Cor = 0
 Aux_Edição = 0
 Aux_Data = CDate(Data_Ini.Text)
Lp1:
 rsEstoque.Seek ">", Val(Combo_Filial.Text), CDate(Aux_Data), Aux_Prod, Aux_Tamanho, Aux_Cor, Aux_Edição
 If rsEstoque.NoMatch Then GoTo Fim_Lp1
 If rsEstoque("Filial") <> Val(Combo_Filial.Text) Then GoTo Fim_Lp1
 If rsEstoque("Data") > CDate(Data_Fim.Text) Then GoTo Fim_Lp1
 
 Call StatusMsg("Verificando dia " & str(Aux_Data))
 
 Aux_Data = rsEstoque("Data")
 Aux_Prod = rsEstoque("Produto")
 Aux_Tamanho = rsEstoque("Tamanho")
 Aux_Cor = rsEstoque("Cor")
 Aux_Edição = rsEstoque("Edição")
 
 If O_Vendas_Zero.Value = 1 Then
   If rsEstoque("Vendas") = 0 Then GoTo Lp1
 End If
 
 
 If Nome_Fornecedor.Caption <> "" Then
   rsForn_Prod.Seek "=", Aux_Prod, Val(Combo_Fornecedor.Text)
   If rsForn_Prod.NoMatch Then GoTo Lp1
 End If
  
 If O_Normal.Value = True Then
   Aux_Tamanho2 = 0
   Aux_Cor2 = 0
   Aux_Edição2 = 0
 End If
 If O_Grade.Value = True Then
   Aux_Tamanho2 = Aux_Tamanho
   Aux_Cor2 = Aux_Cor
   Aux_Edição2 = 0
 End If
 If O_Edição.Value = True Then
   Aux_Tamanho2 = 0
   Aux_Cor2 = 0
   Aux_Edição2 = Aux_Edição
 End If
 
  rsVendas.Seek "=", Aux_Prod, Aux_Tamanho2, Aux_Cor2, Aux_Edição2
  If rsVendas.NoMatch Then
    rsVendas.AddNew
    rsVendas("Produto") = Aux_Prod
    rsVendas("Tamanho") = Aux_Tamanho2
    rsVendas("Cor") = Aux_Cor2
    rsVendas("Edição") = Aux_Edição2
    rsVendas("Classe") = rsEstoque("Classe")
    rsVendas("Sub Classe") = rsEstoque("Sub Classe")
  Else
    rsVendas.Edit
  End If
    
  rsVendas("Vendas") = rsVendas("Vendas") + rsEstoque("Vendas")
  rsVendas("Valor Vendas") = rsVendas("Valor Vendas") + rsEstoque("Valor Vendas") '- rsEstoque("Valor Devolução")
  rsVendas.Update
  GoTo Lp1
 
Fim_Lp1:
 
 Rem  Seta Valores e Manda Relatório

 If O_Classe.Value = 1 Then
   Rel.WindowShowGroupTree = True
 Else
   Rel.WindowShowGroupTree = False
 End If
 
 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If O_Vídeo = True Then Rel.Destination = 0
 If O_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Normal.Value = True Then
   If O_Classe.Value = 0 Then Str1 = gsReportPath & "VENDA1.RPT"
   If O_Classe.Value = 1 Then Str1 = gsReportPath & "VENDA3.RPT"
 End If
 If O_Grade.Value = True Then
   If O_Classe.Value = 0 Then Str1 = gsReportPath & "VENDA1G.RPT"
   If O_Classe.Value = 1 Then Str1 = gsReportPath & "VENDA3G.RPT"
 End If
 If O_Edição.Value = True Then
   If O_Classe.Value = 0 Then Str1 = gsReportPath & "VENDA1E.RPT"
   If O_Classe.Value = 1 Then Str1 = gsReportPath & "VENDA3E.RPT"
 End If
 Rel.ReportFileName = Str1

 Rem Seleção
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Estoque.Filial} =" + Combo_Filial.Text
 Str_Rel = Str_Rel + " And {Estoque.Data} >="
 Str_Rel = Str_Rel + Str_Data1
 Str_Rel = Str_Rel + " And {Estoque.Data} <=" + Str_Data2

 'Rel.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel.Formulas(1) = Str_Rel

 Rem data inicial
 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini.Text + "'"
 Rel.Formulas(2) = Str_Rel

 Rem data final
 Str_Rel = "data_fim = '"
 Str_Rel = Str_Rel + Data_Fim.Text + "'"
 Rel.Formulas(3) = Str_Rel

  
  '-----------------------------------------------------------------------------------
  '20/12/2002 - mpdea
  'Corrigido passagem do desconto quando a opção chkShowDescSubTotal
  'não está ativada
  '
  '30/09/2002 - mpdea
  'Corrigido Run-Time 94 através da função IsDataType
  '
  '19/09/2002 - mpdea
  'Inclusão do totalizador Desconto no SubTotal
  Rel.Formulas(4) = "ShowDescSubTotal = " & chkShowDescSubTotal.Value
  
  If chkShowDescSubTotal.Value = vbChecked Then
    strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
             "Filial = " & Val(Combo_Filial.Text) & " AND " & _
             "Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
             "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
    
    Set rsDescSubTotal = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rsDescSubTotal
      Call IsDataType(dtCurrency, .Fields("Total").Value, curDescSubTotal)
      .Close
    End With
    Set rsDescSubTotal = Nothing
  Else
    curDescSubTotal = 0
  End If
  
  Rel.Formulas(5) = "DescSubTotal = " & Replace(curDescSubTotal, gsCurrencyDecimal, ".")
  '-----------------------------------------------------------------------------------


 Rem ORdem
 Str_Rel = ""
 Rel.SortFields(0) = ""
 Rel.GroupSortFields(0) = ""
 If O_Código.Value = True Then
   Rel.SortFields(0) = "+{Produtos.Código Ordenação}"
 End If
 If O_Nome.Value = True Then
   Rel.SortFields(0) = "+{Produtos.Nome}"
 End If
 If O_Unidades.Value = True Then
   Rel.SortFields(0) = "-{ZZZVendas.Vendas}"
 End If
 If O_Valores.Value = True Then
   Rel.SortFields(0) = "-{ZZZVendas.Valor Vendas}"
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

Private Sub chkShowDescSubTotal_Click()
  '02/10/2002 - mpdea
  'Desativa opção fornecedor para exibição correta do Desconto no SubTotal
  If chkShowDescSubTotal.Value = vbChecked Then
    Combo_Fornecedor.Text = ""
    Nome_Fornecedor.Caption = ""
  End If
End Sub

Private Sub Combo_Filial_CloseUp()
 Combo_Filial.Text = Combo_Filial.Columns(1).Text
 Combo_Filial_LostFocus
End Sub

Private Sub Combo_Filial_LostFocus()

 Nome_Empresa.Caption = ""
 If IsNull(Combo_Filial.Text) Then Exit Sub
 If Not IsNumeric(Combo_Filial.Text) Then Exit Sub
 If Val(Combo_Filial.Text) > 99 Then Exit Sub
 rsParametros.Index = "Filial"
 rsParametros.Seek "=", Val(Combo_Filial.Text)
 If rsParametros.NoMatch Then Exit Sub
 Nome_Empresa.Caption = rsParametros("Nome")
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Call StatusMsg("")
  Nome_Fornecedor.Caption = ""
  rsFornecedores.Index = "Código"
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Combo_Fornecedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  If Val(Combo_Fornecedor.Text) < 1 Then Exit Sub
  rsFornecedores.Seek "=", Val(Combo_Fornecedor.Text)
  If rsFornecedores.NoMatch Then Exit Sub
  Nome_Fornecedor.Caption = rsFornecedores("Nome") & ""
  '02/10/2002 - mpdea
  'Desativa opção Desconto no SubTotal,
  'pois o mesmo é por venda e não por produto (devido ao filtro fornecedor)
  chkShowDescSubTotal.Value = vbUnchecked
End Sub

Private Sub Command1_Click()
  frmRelVendas2.Show
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
  Set rsForn_Prod = db.OpenRecordset("Forn_Prod", , dbReadOnly)
  
  Combo_Filial.Text = gnCodFilial
  Data_Fim.Text = gsFormatDate(Date)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName

  O_Grade.Enabled = gbGrade
  O_Edição.Enabled = gbEdicao

End Sub
