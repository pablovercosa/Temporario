VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelContasReceberPorDtEmissao 
   Caption         =   " Lançamentos de Contas a Receber por Data de Emissão"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelContasReceberPorDtEmissao.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dtaVendedor 
      Caption         =   "dtaVendedor"
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
      Left            =   4140
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   3450
      Visible         =   0   'False
      Width           =   1800
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
      Left            =   60
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3450
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   3450
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   795
      Left            =   3720
      TabIndex        =   17
      Top             =   2100
      Width           =   1575
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   480
         Width           =   1185
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      Height          =   400
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   7920
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opção"
      Height          =   795
      Left            =   5370
      TabIndex        =   12
      Top             =   2100
      Width           =   2640
      Begin VB.OptionButton O_Resumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton O_Completo 
         Caption         =   "Completo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton O_Banco 
         Caption         =   "Para Banco"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1335
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Recebimento"
      Height          =   1080
      Left            =   75
      TabIndex        =   5
      Top             =   960
      Width           =   7920
      Begin VB.OptionButton O_Todos 
         Caption         =   "Todos"
         Height          =   225
         Left            =   105
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Carteira 
         Caption         =   "Carteira"
         Height          =   225
         Left            =   105
         TabIndex        =   9
         Top             =   720
         Width           =   1065
      End
      Begin VB.OptionButton O_Carnet 
         Caption         =   "Carnet"
         Height          =   225
         Left            =   1365
         TabIndex        =   8
         Top             =   735
         Width           =   1065
      End
      Begin VB.OptionButton O_Banco1 
         Caption         =   "Banco"
         Height          =   225
         Left            =   1350
         TabIndex        =   7
         Top             =   315
         Width           =   855
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
         Bindings        =   "frmRelContasReceberPorDtEmissao.frx":4E95A
         DataSource      =   "Data2"
         Height          =   345
         Left            =   2280
         TabIndex        =   6
         Top             =   255
         Width           =   840
         DataFieldList   =   "Descrição"
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
         BackColorOdd    =   8454143
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   6376
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3704
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1720
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   1482
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Enabled         =   0   'False
      End
      Begin VB.Label Nome_Banco 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3150
         TabIndex        =   11
         Top             =   255
         Width           =   4605
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   795
      Left            =   75
      TabIndex        =   0
      Top             =   2100
      Width           =   3585
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
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
         Height          =   285
         Left            =   510
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Até"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "De"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   3
         Top             =   330
         Width           =   285
      End
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   6120
      Top             =   3360
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
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "frmRelContasReceberPorDtEmissao.frx":4E96E
      DataSource      =   "Data1"
      Height          =   345
      Left            =   885
      TabIndex        =   20
      Top             =   75
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
      BackColorOdd    =   8454143
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8520
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1667
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "frmRelContasReceberPorDtEmissao.frx":4E982
      DataSource      =   "dtaVendedor"
      Height          =   345
      Left            =   885
      TabIndex        =   23
      Top             =   510
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
      BevelColorFrame =   0
      BevelColorHighlight=   16777215
      BackColorOdd    =   16777152
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
      _ExtentX        =   1667
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor"
      Height          =   195
      Left            =   105
      TabIndex        =   25
      Top             =   570
      Width           =   690
   End
   Begin VB.Label Nome_Vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1860
      TabIndex        =   24
      Top             =   510
      Width           =   6135
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1860
      TabIndex        =   22
      Top             =   75
      Width           =   6135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      TabIndex        =   21
      Top             =   150
      Width           =   375
   End
End
Attribute VB_Name = "frmRelContasReceberPorDtEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsContas As Recordset
Private rsVendedor As Recordset

Private Sub B_Imprime_Click()
  Dim Val1, Val2, Erro As Integer
  Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
  Dim Str_Rel As String
  Dim Data1 As Variant
   
  Call StatusMsg("")
  
  Rem Verifica empresa
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo.SetFocus
    Exit Sub
  End If

  If Filial_Liberada <> 0 Then
    If Val(Combo.Text) <> Filial_Liberada Then
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
  
  Rem  Nome do BD
  Str1 = gsQuickDBFileName
  Rel1.DataFiles(0) = Str1
  
  Rem Saída
  If B_Vídeo = True Then Rel1.Destination = 0
  If B_Impressora = True Then Rel1.Destination = 1
  
  Rem Nome do arquivo .rpt
  If O_Resumido.Value = True Then Str1 = gsReportPath & "Recebe_Contas_por_dataEmissao.rpt"
  If O_Completo.Value = True Then Str1 = gsReportPath & "RECEBE1C.RPT"
  If O_Banco.Value = True Then Str1 = gsReportPath & "RECEBE1B.RPT"
  
  Rel1.ReportFileName = Str1
  
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

  Rem Seleção
  Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
  Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
  
  Str_Rel = "{Contas a Receber.Filial} =" + Combo.Text
  
  If Nome_Vendedor.Caption <> "" Then
      Str_Rel = Str_Rel + " And {Contas a Receber.Vendedor} =" + Combo_Vendedor.Text
  End If
  
  Str_Rel = Str_Rel + " And {Contas a Receber.Data Emissão} >="
  Str_Rel = Str_Rel + Str_Data1
  Str_Rel = Str_Rel + " And {Contas a Receber.Data Emissão} <=" + Str_Data2
  Str_Rel = Str_Rel + " And {Contas a Receber.Valor Recebido} = 0"
  Str_Rel = Str_Rel + " And {Contas a Receber.Tipo} = 'R'"
  
  If O_Carteira.Value = True Then
    Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'C'"
  End If
  If O_Carnet.Value = True Then
    Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'T'"
  End If
  If O_Banco1.Value = True Then
    Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'B'"
    If Nome_Banco.Caption <> "" Then
      Str_Rel = Str_Rel + " And {Contas a Receber.Conta Boleto} = " + str(Combo_Banco.Text)
    End If
  End If

  Rel1.SelectionFormula = Str_Rel
  
  Str_Rel = "nome_empresa = '"
  '''Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  Str_Rel = Str_Rel + Nome_Vendedor.Caption + "'"
    
  Rel1.Formulas(0) = Str_Rel
  
  Str_Rel = "nome_filial = '"
  Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
  Rel1.Formulas(1) = Str_Rel
  
  Rem data inicial
  Str_Rel = "data_ini = '"
  Str_Rel = Str_Rel + Data_Ini.Text + "'"
  Rel1.Formulas(2) = Str_Rel
  
  Rem data final
  Str_Rel = "data_fim = '"
  Str_Rel = Str_Rel + Data_Fim.Text + "'"
  Rel1.Formulas(3) = Str_Rel

  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
   
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)

   Rel1.Action = 1
  
   Call StatusMsg("")
   MousePointer = vbDefault
End Sub

Private Sub Combo_Banco_CloseUp()
  Combo_Banco.Text = Combo_Banco.Columns(2).Text
  Combo_Banco_LostFocus
End Sub

Private Sub Combo_Banco_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Banco_LostFocus()

  Call StatusMsg("")
  Nome_Banco.Caption = ""
  
  If IsNull(Combo_Banco.Text) Then Exit Sub
  If Combo_Banco.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Banco.Text) Then Exit Sub
  If Val(Combo_Banco.Text) > 9999 Then Exit Sub
  If Val(Combo_Banco.Text) < 1 Then Exit Sub
    
  rsContas.Index = "Código"
  
  rsContas.Seek "=", Val(Combo_Banco.Text)
  If rsContas.NoMatch Then Exit Sub
  
  Nome_Banco.Caption = rsContas("Descrição") & ""
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

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(1).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()
  Call StatusMsg("")
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) > 999 Then Exit Sub
  rsVendedor.Index = "Código"
  rsVendedor.Seek "=", Val(Combo_Vendedor.Text)
  If rsVendedor.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsVendedor("Nome")
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
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsVendedor = db.OpenRecordset("Funcionários", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  dtaVendedor.DatabaseName = gsQuickDBFileName
  
  Combo.Text = gnCodFilial
  Combo_LostFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsVendedor.Close
    rsParametros.Close
    rsContas.Close
    
    Set rsContas = Nothing
    Set rsParametros = Nothing
    Set rsVendedor = Nothing
End Sub

Private Sub O_Banco1_Click()
  Combo_Banco.Enabled = True
  Nome_Banco.Enabled = True
End Sub

Private Sub O_Carnet_Click()
  Combo_Banco.Enabled = False
  Nome_Banco.Enabled = False
End Sub

Private Sub O_Carteira_Click()
  Combo_Banco.Enabled = False
  Nome_Banco.Enabled = False
End Sub

Private Sub O_Todos_Click()
  Combo_Banco.Enabled = False
  Nome_Banco.Enabled = False
End Sub



