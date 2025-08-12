VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelContatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Contatos - Pendências"
   ClientHeight    =   2730
   ClientLeft      =   1650
   ClientTop       =   2250
   ClientWidth     =   5805
   Icon            =   "RelContatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   5805
   Begin VB.Data Data3 
      Caption         =   "Funcionário"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período de Aviso para Contatos Pendentes"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   135
      TabIndex        =   12
      Top             =   960
      Width           =   3975
      Begin MSMask.MaskEdBox Dia_Fim 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   315
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
      Begin MSMask.MaskEdBox Dia_Ini 
         Height          =   315
         Left            =   645
         TabIndex        =   1
         Top             =   315
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
         AutoSize        =   -1  'True
         Caption         =   "Final :"
         Height          =   195
         Left            =   2040
         TabIndex        =   14
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicial :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   375
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opção"
      Height          =   855
      Left            =   1640
      TabIndex        =   11
      Top             =   1695
      Width           =   2480
      Begin VB.OptionButton O_Pendente 
         Caption         =   "Somente pendentes"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton O_Todos 
         Caption         =   "Todos os contatos"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   1695
      Width           =   1455
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
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
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cli_For"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   4920
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_cli 
      Bindings        =   "RelContatos.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   180
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
      Columns.Count   =   3
      Columns(0).Width=   7805
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1482
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   900
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "RelContatos.frx":059E
      DataSource      =   "Data3"
      Height          =   315
      Left            =   960
      TabIndex        =   16
      Top             =   600
      Width           =   720
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
      Columns.Count   =   3
      Columns(0).Width=   6006
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2990
      Columns(1).Caption=   "Apelido"
      Columns(1).Name =   "Apelido"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Apelido"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1773
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Vendedor 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   17
      Top             =   600
      Width           =   3915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Nome_cli 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   180
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cli / For:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "frmRelContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsClientes As Recordset
Dim rsFuncionarios As Recordset  '20/12/2006 - Anderson - Supri Print - Criação do Filtro Vendedor para o relatório de contatos pendentes


Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 'Dim Erro As Integer
 'Dim Str_Data1, Str_Data2 As String
 
 Call StatusMsg("")

 'Verifica Data
 Erro = False
 If IsNull(Dia_Ini.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Dia_Ini.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Dia_Ini.SetFocus
   Exit Sub
 End If
 
 'Verifica Data Final
 Erro = False
 If IsNull(Dia_Fim.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Dia_Fim.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Dia_Fim.SetFocus
   Exit Sub
 End If

 If CDate(Dia_Ini.Text) > CDate(Dia_Fim.Text) Then
   DisplayMsg "Data inicial deve ser menor ou igual a data final."
   Dia_Ini.SetFocus
   Exit Sub
 End If
 
 'Seta Valores e Manda Relatório

 'Nome do BD
 Str1 = gsQuickDBFileName
 Rel1.DataFiles(0) = Str1

 'Saída
 If O_Vídeo = True Then Rel1.Destination = 0
 If O_Impressora = True Then Rel1.Destination = 1
 Rem If B_Arquivo = True Then
 Rem    frmMenu.Relatório.Destination = 2
 Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
 Rem End If

 'Nome do arquivo .rpt
 Str1 = gsReportPath & "CONTATO.RPT"
 Rel1.ReportFileName = Str1

 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 'Seleção
 '
 '14/06/2005 - Daniel
 'Correção para pegar todos os registros: Cláusula [Else]
 '
 Str_Rel = ""
 If O_Todos.Value = False Then
  Str_Rel = "{Contatos Efetuados.Pendência} = True"
 Else
  '11/08/2005 - Daniel
  'Adicionado o ( ) para não dar erros na cláusula "OR"
  'Correção disponível a partir da beta 6.52.0.72
  Str_Rel = " ({Contatos Efetuados.Pendência} = True OR {Contatos Efetuados.Pendência} = False ) "
 End If
 
 If Nome_cli.Caption <> "" Then
   If Str_Rel <> "" Then Str_Rel = Str_Rel + " AND "
   Str_Rel = Str_Rel + "{cli_for.código} =" + Combo_cli.Text
 End If
 
 '20/12/2006 - Anderson - Supri Print - Criação do Filtro Vendedor para o relatório de contatos pendentes
 If Nome_Vendedor.Caption <> "" Then
   If Str_Rel <> "" Then Str_Rel = Str_Rel + " AND "
   Str_Rel = Str_Rel + "{cli_for.vendedor} =" + Combo_Vendedor.Text
 End If

 Str_Data1 = "Date" + Format$(Dia_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Dia_Fim.Text, "(yyyy,mm,dd)")
  
 If Str_Rel <> "" Then Str_Rel = Str_Rel + " AND "
 
 '15/06/2005 - Daniel
 'Correção para buscar todos os registros
 If O_Todos.Value Then
  Str_Rel = Str_Rel + " {Contatos Efetuados.Data} >= " + Str_Data1
  Str_Rel = Str_Rel + " And {Contatos Efetuados.Data} <= " + Str_Data2
 Else
  Str_Rel = Str_Rel + " {Contatos Efetuados.Data Aviso} >= " + Str_Data1
  Str_Rel = Str_Rel + " And {Contatos Efetuados.Data Aviso} <= " + Str_Data2
 End If
 
 Rel1.SelectionFormula = Str_Rel
 
 Rem Str_Rel = "STR_NOME = 'Empresa " + (DC_Empresas.Text)
 Rem Str_Rel = Str_Rel + " - " + C_Nome_Empresa + " de " + C_Data_Ini.Text + " a " + C_Data_Fim.Text + "'"
 Rem frmMenu.Relatório.Formulas(0) = Str_Rel
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rem Str_Rel = "ttttt"
 Rel1.Formulas(0) = Str_Rel

 Str_Rel = "tipo = '"
 If O_Todos.Value = True Then Str_Rel = Str_Rel + "Todos os Contatos'"
 If O_Todos.Value = False Then Str_Rel = Str_Rel + "Contatos Pendentes'"

 Rel1.Formulas(1) = Str_Rel

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub Combo_cli_CloseUp()
 Combo_cli.Text = Combo_cli.Columns(1).Text
 Combo_cli_LostFocus
End Sub

Private Sub Combo_cli_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_cli_LostFocus()
  Nome_cli.Caption = ""
  If IsNull(Combo_cli.Text) Then Exit Sub
  If Not IsNumeric(Combo_cli.Text) Then Exit Sub
  If Val(Combo_cli.Text) < 0 Or Val(Combo_cli.Text) > 99999999 Then Exit Sub

  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_cli.Text)
  If rsClientes.NoMatch Then Exit Sub
  Nome_cli.Caption = rsClientes("Nome")

  Call StatusMsg("")
End Sub

Private Sub Dia_Fim_LostFocus()
 Dia_Fim.Text = Ajusta_Data(Dia_Fim.Text)
End Sub

Private Sub Dia_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Dia_Fim.Text = frmCalendario.gsDateCalender(Dia_Fim.Text)
  End Select
End Sub

Private Sub Dia_Ini_LostFocus()
 Dia_Ini.Text = Ajusta_Data(Dia_Ini.Text)
End Sub

Private Sub Dia_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Dia_Ini.Text = frmCalendario.gsDateCalender(Dia_Ini.Text)
  End Select
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Data1.DatabaseName = gsQuickDBFileName
  
  '20/12/2006 - Anderson - Supri Print - Criação do Filtro Vendedor para o relatório de contatos pendentes
  Data3.DatabaseName = gsQuickDBFileName
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)

  Dia_Ini.Text = gsFormatDate(Data_Atual)
  Dia_Fim.Text = gsFormatDate(Data_Atual)
End Sub

Private Sub O_Pendente_Click()
  '15/06/2005 - Daniel
  'Correção para filtrar todos ou somente os pendentes
  Frame3.Caption = "Período de Aviso para Contatos Pendentes"
  Frame3.ForeColor = &H80000012
End Sub

Private Sub O_Todos_Click()
  '15/06/2005 - Daniel
  'Correção para filtrar todos ou somente os pendentes
  Frame3.Caption = "Período de Lançamento dos Contatos"
  Frame3.ForeColor = &HFF&
End Sub

'20/12/2006 - Anderson - Supri Print  - Criação do Filtro Vendedor para o relatório de contatos pendentes
Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
  Combo_Vendedor_LostFocus
End Sub

'20/12/2006 - Anderson - Supri Print  - Criação do Filtro Vendedor para o relatório de contatos pendentes
Private Sub Combo_Vendedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

'20/12/2006 - Anderson - Supri Print  - Criação do Filtro Vendedor para o relatório de contatos pendentes
Private Sub Combo_Vendedor_LostFocus()
  Call StatusMsg("")
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) > 9999 Then Exit Sub
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Vendedor.Text)
  If rsFuncionarios.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsFuncionarios("Apelido")
End Sub

