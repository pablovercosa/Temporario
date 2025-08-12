VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelMoedas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Moedas"
   ClientHeight    =   1770
   ClientLeft      =   2580
   ClientTop       =   2160
   ClientWidth     =   3345
   ForeColor       =   &H80000008&
   Icon            =   "RelMoedas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   3345
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   1950
      TabIndex        =   4
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   600
      Top             =   1320
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
Attribute VB_Name = "frmRelMoedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub B_Cancela_Click()
  Unload Me
End Sub

Private Sub B_Imprime_Click()
 Dim Str_Rel, Str1 As String
 
 Call StatusMsg("")

 Rem  Seta Valores e Manda Relatório

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & "MOEDAS.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 If O_Código.Value = True Then Rel.SortFields(0) = "+{Moedas.Código}"
 If O_Nome.Value = True Then Rel.SortFields(0) = "+{Moedas.Nome}"

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

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
