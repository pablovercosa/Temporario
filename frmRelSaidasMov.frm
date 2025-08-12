VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelSaidasMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Orçamento"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelSaidasMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   6255
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
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2070
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   765
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Width           =   6255
      Begin VB.OptionButton O_vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   510
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   2
         Top             =   330
         Width           =   1275
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   -60
      Top             =   2610
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Validade da proposta"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   1515
   End
End
Attribute VB_Name = "frmRelSaidasMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sequencia As Long
Public Filial As Byte
Public Relatorio As String

Private Sub B_Imprime_Click()
  
  On Error GoTo ErrHandler
  
  'Status
  Call StatusMsg("Aguarde...")
  MousePointer = vbHourglass

  With Rel
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsQuickDBFileName
    .DataFiles(5) = gsQuickDBFileName
    .DataFiles(6) = gsQuickDBFileName
    .DataFiles(7) = gsQuickDBFileName
    
    .Destination = IIf(O_Vídeo.Value, crptToWindow, crptToPrinter)
    .ReportFileName = gsReportPath & Relatorio
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
    
    .SelectionFormula = "{Saídas.Filial} = " & Filial & " AND {Saídas.Sequência} = " & Sequencia

    .Formulas(0) = "mensagem = '" & txtMensagem.Text & "'"

    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
    
    .Action = 1
  End With
  
  Call StatusMsg("")
  MousePointer = vbDefault

  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao imprimir: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
  
  On Error GoTo ErrHandler
  
  txtMensagem.Text = GetSetting("QuickStore", "RelOrcamento", "Mensagem", "")

  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao abrir a tela: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  On Error GoTo ErrHandler
  
  Call SaveSetting("QuickStore", "RelOrcamento", "Mensagem", txtMensagem.Text)

  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao fechar a tela: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub
