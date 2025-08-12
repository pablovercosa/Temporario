VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelFinanc1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Financeiro 1 - Operação com movimentação financeira"
   ClientHeight    =   2625
   ClientLeft      =   2640
   ClientTop       =   3000
   ClientWidth     =   6690
   Icon            =   "RelFinanc1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   6690
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   795
      Left            =   120
      TabIndex        =   12
      Top             =   690
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3690
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
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
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
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   13
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   960
      Left            =   1725
      TabIndex        =   11
      Top             =   1545
      Width           =   1590
      Begin VB.OptionButton O_Com 
         Caption         =   "Com serviço"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   630
         Width           =   1275
      End
      Begin VB.OptionButton O_Sem 
         Caption         =   "Sem serviço"
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   975
      Left            =   105
      TabIndex        =   10
      Top             =   1545
      Width           =   1455
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   285
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3180
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5220
      TabIndex        =   7
      Top             =   2100
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6060
      Top             =   735
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelFinanc1.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   165
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
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2010
      TabIndex        =   9
      Top             =   150
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Filial :"
      Height          =   255
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   585
   End
End
Attribute VB_Name = "frmRelFinanc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset

Private Sub B_Cancela_Click()

End Sub


Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Aux_Data1 As Variant

 
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
 
 Data_Ini.Text = Format(CDate(Data_Ini.Text), "dd/mm/yyyy")
 Data_Fim.Text = Format(CDate(Data_Fim.Text), "dd/mm/yyyy")
 
 If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
   DisplayMsg "Data final menor que data inicial, verifique."
   Data_Fim.SetFocus
   Exit Sub
 End If
 

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If O_Vídeo = True Then Rel.Destination = 0
 If O_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Sem.Value = True Then Str1 = gsReportPath & "FINANC1.RPT"
 If O_Com.Value = True Then Str1 = gsReportPath & "FINANC1S.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 Rem Seleção
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Resumo Diário Financeiro.Filial} =" + Combo_Filial.Text
 Str_Rel = Str_Rel + " And {Resumo Diário Financeiro.Data} >="
 Str_Rel = Str_Rel + Str_Data1
 Str_Rel = Str_Rel + " And {Resumo Diário Financeiro.Data} <=" + Str_Data2

 Rel.SelectionFormula = Str_Rel
 
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

 
 
 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault
 
 
 
 
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
  Data_Fim.Text = Format(Date, "dd/mm/yyyy")
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  Combo_Filial.Text = gnCodFilial
  
  If gbServico = False Then O_Com.Enabled = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

 rsParametros.Close
 Set rsParametros = Nothing
End Sub
