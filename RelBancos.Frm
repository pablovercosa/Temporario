VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Bancos"
   ClientHeight    =   2145
   ClientLeft      =   2205
   ClientTop       =   2250
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   5565
   Visible         =   0   'False
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
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
      Width           =   5355
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5355
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1020
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5355
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2610
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1020
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   120
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
End
Attribute VB_Name = "frmRelBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'10/11/2005 - mpdea
'Incluído tratamento de erro
Private Sub B_Imprime_Click()
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("Aguarde, imprimindo...")
  Screen.MousePointer = vbHourglass
  
  With Rel
    
    .DataFiles(0) = gsQuickDBFileName
    .ReportFileName = gsReportPath & "BANCOS.RPT"
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
    
    If O_Vídeo.Value Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    If O_Código.Value Then
      .SortFields(0) = "+{Bancos.Código}"
    Else
      .SortFields(0) = "+{Bancos.Nome}"
    End If
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
  
    .Action = 1
  End With
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  MsgBox "Ocorreu o erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
