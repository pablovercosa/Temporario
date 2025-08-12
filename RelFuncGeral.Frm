VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelFuncGeral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Funcionários"
   ClientHeight    =   3255
   ClientLeft      =   2205
   ClientTop       =   2250
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelFuncGeral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   7845
   Visible         =   0   'False
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   7575
   End
   Begin VB.Frame frmOpcoes 
      Caption         =   "Opções"
      Height          =   1065
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   7575
      Begin VB.CheckBox chkLiberado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Liberado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4770
         TabIndex        =   7
         ToolTipText     =   "Selecione esta opção para filtrar os funcionários Liberados"
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox chkAtivo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Ativo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4770
         TabIndex        =   6
         ToolTipText     =   "Selecione esta opção para filtrar os funcionários Ativos"
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton optSelecao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Seleção"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3180
         TabIndex        =   5
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton optTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   390
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "{"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   4230
         TabIndex        =   13
         Top             =   165
         Width           =   225
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Im&primir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2070
      Width           =   7575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   675
      Left            =   3930
      TabIndex        =   11
      Top             =   120
      Width           =   3765
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1860
         TabIndex        =   3
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
      Height          =   675
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3765
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1710
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   120
      Top             =   3480
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
End
Attribute VB_Name = "frmRelFuncGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Imprime_Click()
  '08/08/2005 - Daniel
  'Adicionado tratamento de erro
  On Error GoTo TratarErro
  
  Call StatusMsg("")
  
  '08/08/2005 - Daniel
  'Validação para optSelecao
  If optSelecao.Value Then
    If chkAtivo.Value = vbUnchecked And chkLiberado.Value = vbUnchecked Then
      MsgBox "Escolha uma das seguintes opções: [Ativo] ou [Liberado] ou [Ativo e Liberado]", vbExclamation, "Quick Store"
      Exit Sub
    End If
  End If
  
  'Nome do BD
  Rel.DataFiles(0) = gsQuickDBFileName
  
  'Saída
  If O_Vídeo.Value Then
    Rel.Destination = 0
  ElseIf O_Impressora.Value Then
    Rel.Destination = 1
  End If
  
  'Nome do arquivo .rpt
  Rel.ReportFileName = gsReportPath & "USUARIOS.RPT"
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel
  
  '08/08/2005 - Daniel
  'Criado tratamento para atender às verificações para exibição de Todos
  'os Funcionários, Ativos, Liberados ou Ativos e Liberados
  Dim strSelect As String
  
  If optTodos.Value Then
    strSelect = ""
  Else
  
    If chkAtivo.Value = vbChecked Then
      strSelect = "{Funcionários.Ativo}"
    Else
      strSelect = "NOT {Funcionários.Ativo}"
    End If
  
    If chkLiberado.Value = vbChecked Then
      strSelect = strSelect & " AND {Funcionários.Liberado}"
    Else
      strSelect = strSelect & " AND NOT {Funcionários.Liberado}"
    End If
  End If
  
  Rel.SelectionFormula = strSelect
  '---------------------------------------------------------------------
  
  If O_Código.Value Then
    Rel.SortFields(0) = "+{Funcionários.Código}"
  ElseIf O_Nome.Value Then
    Rel.SortFields(0) = "+{Funcionários.Nome}"
  End If
  
  Rel.Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  Rel.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault

  Exit Sub

TratarErro:
  Call StatusMsg("")
  MousePointer = vbDefault
  MsgBox "Ocorrência de erro ao imprimir." & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Err.Clear
  
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
