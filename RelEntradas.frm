VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEntradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Entradas"
   ClientHeight    =   3555
   ClientLeft      =   1650
   ClientTop       =   2520
   ClientWidth     =   6840
   HelpContextID   =   1570
   Icon            =   "RelEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   6840
   Begin VB.Data datFornecedor 
      Caption         =   "datFornecedor"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Nome"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datCentroCusto 
      Caption         =   "CC"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   795
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3690
         TabIndex        =   5
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
         TabIndex        =   4
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
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   22
         Top             =   375
         Width           =   885
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2205
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Op_Entrada"
      Top             =   4125
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opção"
      Height          =   855
      Left            =   3195
      TabIndex        =   19
      Top             =   2595
      Width           =   2055
      Begin VB.OptionButton com_produtos 
         Caption         =   "Imprimir produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   495
         Width           =   1695
      End
      Begin VB.OptionButton sem_produtos 
         Caption         =   "Não imprimir produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   1665
      TabIndex        =   18
      Top             =   2580
      Width           =   1335
      Begin VB.OptionButton O_Completo 
         Caption         =   "Completo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   495
         Width           =   1095
      End
      Begin VB.OptionButton O_Resumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   2565
      Width           =   1335
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton O_vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5430
      TabIndex        =   12
      Top             =   3045
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4125
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   5985
      Top             =   2040
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Operação 
      Bindings        =   "RelEntradas.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1215
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
      Columns(1).Width=   2037
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelEntradas.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Top             =   240
      Width           =   1200
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
      _ExtentX        =   2117
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo cboCodigoCC 
      Bindings        =   "RelEntradas.frx":05B2
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      DataFieldList   =   "Nome"
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Codigo"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Código"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Nome"
   End
   Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
      Bindings        =   "RelEntradas.frx":05CF
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
      DataFieldList   =   "Nome"
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Codigo"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Código"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Nome"
   End
   Begin VB.Label Label6 
      Caption         =   "Fornecedor:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label lblNomeFornecedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   25
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Centro de Custo:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   990
      Width           =   1335
   End
   Begin VB.Label lblNomeCC 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   23
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Nome_Operação 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   600
      Width           =   3840
   End
   Begin VB.Label Label2 
      Caption         =   "Operação :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Filial :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   13
      Top             =   225
      Width           =   3840
   End
End
Attribute VB_Name = "frmRelEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'05/05/2005 - Daniel
'
'Projeto: Melhorias para o Centro de Custo
'
'A partir da versão 6.52.0.38 todo campo de Centro de Custo
'estará carregando apenas os Centros que estão ativos no sistema

Private rsParametros As Recordset
Private rsOp_Entrada As Recordset
Private rsEntradas As Recordset

Private Sub B_Imprime_Click()
  Dim Erro As Integer
  Dim Str1 As String, Str_Data1 As String, Str_Data2 As String
  Dim sSql As String
  Dim Aux_Data As Date
  Dim Aux_Seq As Long
  
   
  Call StatusMsg("")
  
  Rem Verifica empresa
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo_Filial.SetFocus
    Exit Sub
  End If

  If Filial_Liberada <> 0 Then
    If Val(Combo_Filial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
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
  
  
  '02/09/2003 - mpdea
  'Status
  Call StatusMsg("Aguarde...")
  
  
  sSql = "DELETE * FROM Entradas WHERE CodUsuarioOwner = " & CStr(gnUserCode)
  Call dbTemp.Execute(sSql)
  
  rsEntradas.Index = "Data"
  Aux_Data = CDate(Data_Ini.Text)
  Aux_Seq = 0
 
 
Lp1:
  rsEntradas.Seek ">", Val(Combo_Filial.Text), Aux_Data, Aux_Seq
  If rsEntradas.NoMatch Then GoTo Imprime
  If rsEntradas("Filial") <> Val(Combo_Filial.Text) Then GoTo Imprime
  If rsEntradas("Data") > CDate(Data_Fim.Text) Then GoTo Imprime
  
  Aux_Data = rsEntradas("Data")
  Aux_Seq = rsEntradas("Sequência")
  
  If Nome_Operação.Caption <> "" Then
    If rsEntradas("Operação") <> Val(Combo_Operação.Text) Then GoTo Lp1
  End If
  
   '02/09/2003 - mpdea
   'Comentado devido a perda de performance
'  Call StatusMsg("Aguarde, verificando movimentação " + str(Aux_Seq))
  
  Call Grava_Temp_Entradas(Val(Combo_Filial.Text), Aux_Seq)
  
  GoTo Lp1


Imprime:
  Rem  Nome do BD
   With Rel1
     .DataFiles(0) = gsTempDBFileName
     .DataFiles(1) = gsQuickDBFileName
   End With
   Rel1.SelectionFormula = "{Entradas.CodUsuarioOwner} = " & CStr(gnUserCode)
  
  Rem Saída
  If O_Vídeo = True Then Rel1.Destination = 0
  If O_Vídeo = False Then Rel1.Destination = 1
  
  Rem Nome do arquivo .rpt
  If O_Resumido.Value = True Then Str1 = gsReportPath & "Entrada2.RPT"
  If O_Completo.Value = True Then Str1 = gsReportPath & "Entrada1.RPT"
  
  Rel1.ReportFileName = Str1
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel1
  
  Rem Seleção
  Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
  Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
  
  
  sSql = "{Entradas.Data} >="
  sSql = sSql + Str_Data1
  sSql = sSql + " And {Entradas.Data} <=" + Str_Data2
  
  
  If Nome_Operação.Caption <> "" Then
    sSql = sSql + " And {Entradas.Cód Operação} = " + Combo_Operação.Text
  End If
  
  If lblNomeCC.Caption <> "" Then
    sSql = sSql + " And {Entradas.CentroCusto} = " + cboCodigoCC.Text
  End If
  
  '06/12/2004 - Daniel
  'Adicionado filtro por Fornecedor
  If lblNomeFornecedor.Caption <> "" Then
    sSql = sSql + " And {Entradas.Cód Fornecedor} = " + cboFornecedor.Text
  End If
  
  Rel1.SelectionFormula = sSql
  
  sSql = "nome_empresa = '"
  sSql = sSql + gsNomeEmpresa + "'"

  Rel1.Formulas(0) = sSql
  
  sSql = "filial = '"
  sSql = sSql + Nome_Empresa.Caption + "'"
  Rel1.Formulas(1) = sSql
  
  Rem data inicial
  sSql = "data_ini = '"
  sSql = sSql + Data_Ini.Text + "'"
  Rel1.Formulas(2) = sSql
  
  Rem data final
  sSql = "data_fim = '"
  sSql = sSql + Data_Fim.Text + "'"
  Rel1.Formulas(3) = sSql
  
  If sem_produtos.Value = True Then sSql = "emite_produtos = 'NAO'"
  If com_produtos.Value = True Then sSql = "emite_produtos = 'SIM'"
  
  Rel1.Formulas(4) = sSql
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
 
 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub cboCodigoCC_CloseUp()
  cboCodigoCC.Text = Trim(cboCodigoCC.Columns(0).Text)
  cboCodigoCC_LostFocus
End Sub

Private Sub cboCodigoCC_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboCodigoCC_LostFocus()
  Dim rs As Recordset
  Set rs = db.OpenRecordset("SELECT * FROM [Centros de Custo]")  'datCentroCusto.Recordset.Clone
    
  lblNomeCC.Caption = ""
  
  If Not IsNumeric(cboCodigoCC.Text) Then Exit Sub
  
  rs.FindFirst "Código = " & Trim(cboCodigoCC.Text)
  If Not rs.NoMatch Then
    lblNomeCC.Caption = rs!Nome & ""
  End If
End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

Private Sub cboFornecedor_LostFocus()
  Dim rs As Recordset
  Set rs = db.OpenRecordset("SELECT Código, Nome FROM Cli_For")
    
  lblNomeFornecedor.Caption = ""
  
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  rs.FindFirst "Código = " & Trim(cboFornecedor.Text)
  If Not rs.NoMatch Then
    lblNomeFornecedor.Caption = rs!Nome & ""
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

Private Sub Combo_Operação_CloseUp()
 Combo_Operação.Text = Combo_Operação.Columns(1).Text
 Combo_Operação_LostFocus
End Sub

Private Sub Combo_Operação_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Operação_LostFocus()
  Call StatusMsg("")
  Nome_Operação.Caption = ""
  If IsNull(Combo_Operação.Text) Then Exit Sub
  If Not IsNumeric(Combo_Operação.Text) Then Exit Sub
  If Val(Combo_Operação.Text) > 999 Then Exit Sub
  rsOp_Entrada.Index = "Código"
  rsOp_Entrada.Seek "=", Val(Combo_Operação.Text)
  If rsOp_Entrada.NoMatch Then Exit Sub
  Nome_Operação.Caption = rsOp_Entrada("Nome")
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

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datCentroCusto.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName
  
  Data_Fim.Text = Format(Date, "dd/mm/yyyy")
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsOp_Entrada = db.OpenRecordset("Operações Entrada", , dbReadOnly)
  Set rsEntradas = db.OpenRecordset("Entradas", , dbReadOnly)

  Combo_Filial.Text = gnCodFilial

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsOp_Entrada.Close
  rsEntradas.Close
  
  Set rsParametros = Nothing
  Set rsOp_Entrada = Nothing
  Set rsEntradas = Nothing
End Sub
