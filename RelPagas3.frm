VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelPagas3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contas Pagas por Fornecedor"
   ClientHeight    =   2850
   ClientLeft      =   1635
   ClientTop       =   1665
   ClientWidth     =   6090
   Icon            =   "RelPagas3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   6090
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   795
      Left            =   90
      TabIndex        =   12
      Top             =   1005
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
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   13
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
      Left            =   2025
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   4650
      TabIndex        =   6
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   90
      TabIndex        =   7
      Top             =   1905
      Width           =   1575
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   525
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3615
      Visible         =   0   'False
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "RelPagas3.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   555
      Width           =   975
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
      Columns(0).Width=   9922
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   5535
      Top             =   1080
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
      Bindings        =   "RelPagas3.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   735
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
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Fornecedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   555
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   9
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmRelPagas3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsCliFor As Recordset

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
 Str1 = gsReportPath & "PAGAS3.RPT"
 
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 Rem Seleção
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Contas a Pagar.Filial} =" + Combo.Text
 Str_Rel = Str_Rel + " And {Contas a Pagar.Pagamento} >="
 Str_Rel = Str_Rel + Str_Data1
 Str_Rel = Str_Rel + " And {Contas a Pagar.Pagamento} <=" + Str_Data2
 Str_Rel = Str_Rel + " And {Contas a Pagar.Valor Pago} <> 0"

 If Nome_Fornecedor.Caption <> "" Then
  Str_Rel = Str_Rel + " And {Contas a Pagar.Fornecedor} = " + str(Combo_Fornecedor.Text)
 End If

 Rel1.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

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
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub Combo_CloseUp()
Combo.Text = Combo.Columns(1).Text
Combo_LostFocus

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
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Combo_Fornecedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  If Val(Combo_Fornecedor.Text) < 0 Then Exit Sub
  If Val(Combo_Fornecedor.Text) > 99999999 Then Exit Sub

  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(Combo_Fornecedor.Text)
  If rsCliFor.NoMatch Then Exit Sub
  Nome_Fornecedor.Caption = rsCliFor("Nome")


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
 Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)

 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName

 Combo.Text = gnCodFilial

End Sub

Private Sub Form_Unload(Cancel As Integer)

 rsParametros.Close
 rsCliFor.Close
 
 Set rsParametros = Nothing
 Set rsCliFor = Nothing
 
End Sub
