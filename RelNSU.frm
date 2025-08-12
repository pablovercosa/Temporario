VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelNSU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Correlação"
   ClientHeight    =   2190
   ClientLeft      =   2955
   ClientTop       =   2715
   ClientWidth     =   6120
   ForeColor       =   &H80000008&
   HelpContextID   =   1460
   Icon            =   "RelNSU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2190
   ScaleWidth      =   6120
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3465
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   105
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   4680
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin MSMask.MaskEdBox DataInicio 
      Height          =   315
      Left            =   1245
      TabIndex        =   1
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   660
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   2625
      Top             =   1680
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
      Bindings        =   "RelNSU.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Top             =   240
      Width           =   750
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
      Columns(0).Width=   5583
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1402
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1323
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin MSMask.MaskEdBox DataFim 
      Height          =   315
      Left            =   3645
      TabIndex        =   2
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   660
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Data Final:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   675
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Data Inicial:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   675
      Width           =   975
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2100
      TabIndex        =   7
      Top             =   240
      Width           =   3900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmRelNSU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim rsParametros As Recordset
  Dim rsCaixa      As Recordset
  Dim rsCaixas     As Recordset

Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 Dim Data2 As Variant
 
 Call StatusMsg("")

 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a Filial."
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
 If IsNull(DataInicio.Text) Then Erro = True
 If Not Erro Then If Not IsDate(DataInicio.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data inicial incorreta, verifique."
   DataInicio.SetFocus
   Exit Sub
 End If
 
 Erro = False
 If IsNull(DataFim.Text) Then Erro = True
 If Not Erro Then If Not IsDate(DataFim.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data final incorreta, verifique."
   DataFim.SetFocus
   Exit Sub
 End If
 
 'Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 'Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1
 Rem If B_Arquivo = True Then
 Rem    frmMenu.Relatório.Destination = 2
 Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
 Rem End If

 'Nome do arquivo .rpt
 Str1 = gsReportPath & "NSU.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 'Seleção
 Str_Data1 = "Date" + Format$(DataInicio.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(DateAdd("d", 1, DataInicio.Text), "(yyyy,mm,dd)")


 Str_Rel = "{NSU.Filial} =" + Combo.Text
 Str_Rel = Str_Rel + " And {NSU.Data_Hora} >="
 Str_Rel = Str_Rel + Str_Data1
 Str_Rel = Str_Rel + " And {NSU.Data_Hora} <"
 Str_Rel = Str_Rel + Str_Data2

 
 Rel.SelectionFormula = Str_Rel
 
 Rem Str_Rel = "STR_NOME = 'Empresa " + (DC_Empresas.Text)
 Rem Str_Rel = Str_Rel + " - " + C_Nome_Empresa + " de " + C_Data_Ini.Text + " a " + C_Data_Fim.Text + "'"
 Rem frmMenu.Relatório.Formulas(0) = Str_Rel
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rem Str_Rel = "ttttt"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "DataInicial = '"
 Str_Rel = Str_Rel + DataInicio.Text + "'"
 Rel.Formulas(1) = Str_Rel

 Str_Rel = "DataFinal = '"
 Str_Rel = Str_Rel + DataFim.Text + "'"
 Rel.Formulas(2) = Str_Rel

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
  
 '25/07/2003 - mpdea
 'Seta a impressora para relatório
 Call SetPrinterName("REL", Rel)

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub Combo_CloseUp()
 Combo.Text = Combo.Columns(1).Text
 Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
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

Private Sub DataInicio_LostFocus()
  DataInicio.Text = Ajusta_Data(DataInicio.Text)
End Sub

Private Sub DataInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      DataInicio.Text = frmCalendario.gsDateCalender(DataInicio.Text)
  End Select
End Sub
Private Sub DataFim_LostFocus()
  DataFim.Text = Ajusta_Data(DataFim.Text)
End Sub

Private Sub DataFim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      DataFim.Text = frmCalendario.gsDateCalender(DataFim.Text)
  End Select
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  
  DataInicio.Text = gsFormatDate(Data_Atual)
  DataFim.Text = gsFormatDate(Data_Atual)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  Set rsParametros = Nothing
End Sub

