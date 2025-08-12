VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Caixa "
   ClientHeight    =   2385
   ClientLeft      =   2955
   ClientTop       =   2715
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1460
   Icon            =   "RelCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2385
   ScaleWidth      =   7875
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
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   600
      Visible         =   0   'False
      Width           =   1905
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
      Bindings        =   "RelCaixa.frx":4E95A
      DataSource      =   "Data2"
      Height          =   330
      Left            =   795
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todos."
      Top             =   600
      Width           =   990
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9419
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1217
      Columns(1).Caption=   "Caixa"
      Columns(1).Name =   "Caixa"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Caixa"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1746
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
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
      Left            =   5460
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   150
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   1050
      Width           =   5445
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   660
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1770
      Width           =   7605
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   330
      Left            =   810
      TabIndex        =   2
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1185
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   582
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   7500
      Top             =   2010
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
      Bindings        =   "RelCaixa.frx":4E96E
      DataSource      =   "Data1"
      Height          =   330
      Left            =   795
      TabIndex        =   0
      Top             =   150
      Width           =   960
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
      BackColorOdd    =   16777152
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
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Nome_Caixa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1830
      TabIndex        =   11
      Top             =   600
      Width           =   5910
   End
   Begin VB.Label Label3 
      Caption         =   "Caixa"
      Height          =   255
      Left            =   135
      TabIndex        =   10
      Top             =   660
      Width           =   540
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   1245
      Width           =   465
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1830
      TabIndex        =   7
      Top             =   150
      Width           =   5910
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   210
      Width           =   435
   End
End
Attribute VB_Name = "frmRelCaixa"
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
 If IsNull(Dia.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Dia.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Dia.SetFocus
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
 Str1 = gsReportPath & "CAIXA.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 'Seleção
 Str_Data1 = "Date" + Format$(Dia.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Caixa.Filial} =" + Combo.Text
 Str_Rel = Str_Rel + " And {Caixa.Data} ="
 Str_Rel = Str_Rel + Str_Data1
 
 Rel.SelectionFormula = Str_Rel
 
 Rem Str_Rel = "STR_NOME = 'Empresa " + (DC_Empresas.Text)
 Rem Str_Rel = Str_Rel + " - " + C_Nome_Empresa + " de " + C_Data_Ini.Text + " a " + C_Data_Fim.Text + "'"
 Rem frmMenu.Relatório.Formulas(0) = Str_Rel
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rem Str_Rel = "ttttt"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "dia = '"
 Str_Rel = Str_Rel + Dia.Text + "'"
 Rel.Formulas(1) = Str_Rel

 If Nome_Caixa.Caption <> "" Then
   Str_Rel = " And {Caixa.Caixa} =" + Combo_Caixa.Text
   Rel.SelectionFormula = Rel.SelectionFormula + Str_Rel
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

Private Sub Combo_Caixa_CloseUp()
 Combo_Caixa.Text = Combo_Caixa.Columns(1).Text
 Combo_Caixa_LostFocus
End Sub

Private Sub Combo_Caixa_LostFocus()
  Nome_Caixa.Caption = ""
  If IsNull(Combo_Caixa.Text) Then Exit Sub
  If Combo_Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Caixa.Text) Then Exit Sub
  If Val(Combo_Caixa.Text) < 0 Then Exit Sub
  If Val(Combo_Caixa.Text) > 99 Then Exit Sub

  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", Val(Combo_Caixa.Text)
  If rsCaixas.NoMatch Then Exit Sub
  Nome_Caixa.Caption = rsCaixas("Descrição")

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

Private Sub Dia_LostFocus()
  Dia.Text = Ajusta_Data(Dia.Text)
End Sub

Private Sub Dia_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Dia.Text = frmCalendario.gsDateCalender(Dia.Text)
  End Select
End Sub

Private Sub Form_Activate()
' Dim Sem_Caixa As Integer
' Dim Retorno As Integer
' Dim Tot_Dinheiro As Double
' Dim Tot_Cheques As Double
' Dim Tot_Pré As Double
' Dim Tot_Cartões As Double
' Dim Tot_Vales As Double
' Dim Saldo_Ant As Double
' Dim Ordem As Long
 


'  'procura para ver se existe caixa no dia
'  Sem_Caixa = False
'  rsCaixa.Index = "Data"
'  rsCaixa.Seek ">", gnCodFilial, 0, Data_Atual, 0
'  If rsCaixa.NoMatch Then Sem_Caixa = True
'  If Sem_Caixa = False Then If gnCodFilial <> rsCaixa("Filial") Then Sem_Caixa = True
'  If Sem_Caixa = False Then If Data_Atual <> rsCaixa("Data") Then Sem_Caixa = True
  
'  If Sem_Caixa = True Then
'     Retorno = MsgBox("O caixa deste dia ainda não foi aberto. É necessário que o caixa seja aberto para a impressão do relatório. Deseja abrir o caixa agora ?", 1, "Atenção")
'     If Retorno = 2 Then
'        Unload frmR_Caixa
'        Exit Sub
'     End If
     
     'Acha o último dia
'     Sem_Caixa = False
'     rsCaixa.Seek "<", gnCodFilial, Data_Atual, 0
'     If rsCaixa.NoMatch Then Sem_Caixa = True
'     If Sem_Caixa = False Then If rsCaixa("Filial") <> gnCodFilial Then Sem_Caixa = True
'
'     If Sem_Caixa = True Then  'Caixa zerado
'       rsCaixa.AddNew
'         rsCaixa("Filial") = gnCodFilial
'         rsCaixa("Data") = Data_Atual
'         rsCaixa("Ordem") = 1
'         rsCaixa("Descrição") = "Início do dia"
'       rsCaixa.Update
'     End If
'     If Sem_Caixa = False Then  'pega último caixa
'       Tot_Dinheiro = rsCaixa("Total Dinheiro")
'       Tot_Cheques = rsCaixa("Total Cheques")
'       Tot_Pré = rsCaixa("Total Cheques Pré")
'       Tot_Cartões = rsCaixa("Total Cartões")
'       Tot_Vales = rsCaixa("Total Vales")
'       Saldo_Ant = rsCaixa("Final")
     
'       rsCaixa.AddNew
'         rsCaixa("Filial") = gnCodFilial
'         rsCaixa("Data") = Data_Atual
'         rsCaixa("Ordem") = 1
'         rsCaixa("Descrição") = "Início do dia"
'         rsCaixa("Dinheiro") = Tot_Dinheiro
'         rsCaixa("Total Dinheiro") = Tot_Dinheiro
'         rsCaixa("Cheques") = Tot_Cheques
'         rsCaixa("Total Cheques") = Tot_Cheques
'         rsCaixa("Cheques Pré") = Tot_Pré
'         rsCaixa("Total Cheques Pré") = Tot_Pré
'         rsCaixa("Cartões") = Tot_Cartões
'         rsCaixa("Total Cartões") = Tot_Cartões
'         rsCaixa("Vales") = Tot_Vales
'         rsCaixa("Total Vales") = Tot_Vales
'         rsCaixa("Saldo Anterior") = Saldo_Ant
'         rsCaixa("Final") = Saldo_Ant
'       rsCaixa.Update
'     End If
'  End If
  

End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  Dia.Text = gsFormatDate(Data_Atual)
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  Set rsParametros = Nothing
  rsCaixa.Close
  Set rsCaixa = Nothing
  rsCaixas.Close
  Set rsCaixas = Nothing
End Sub
