VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelTamanhos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Tamanhos"
   ClientHeight    =   1800
   ClientLeft      =   2205
   ClientTop       =   2250
   ClientWidth     =   3810
   Icon            =   "RelTamanhos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1800
   ScaleWidth      =   3810
   Visible         =   0   'False
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   2385
      TabIndex        =   6
      Top             =   1305
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   1920
      Top             =   15
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
Attribute VB_Name = "frmRelTamanhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Imprime_Click()
 Dim Str_Rel, Str1 As String
 
 
 Call StatusMsg("")

 Rem  Seta Valores e Manda Relatório

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If O_Vídeo = True Then Rel.Destination = 0
 If O_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & "TAMANHOS.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 If O_Código.Value = True Then Rel.SortFields(0) = "+{Tamanhos.Código}"
 If O_Nome.Value = True Then Rel.SortFields(0) = "+{Tamanhos.Nome}"

 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rel.Formulas(0) = Str_Rel


 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault



End Sub


Private Sub M_Escolhe_Click()
End Sub


Private Sub Form_Load()
  Call CenterForm(Me)

End Sub
