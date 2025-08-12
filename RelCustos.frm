VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelCustos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Centros de Custo"
   ClientHeight    =   1935
   ClientLeft      =   1740
   ClientTop       =   2025
   ClientWidth     =   4230
   ForeColor       =   &H80000008&
   Icon            =   "RelCustos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1935
   ScaleWidth      =   4230
   Begin VB.Frame fraSituacao 
      Caption         =   "Situação"
      Height          =   1095
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton optTodosCentros 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optInativo 
         Caption         =   "Inativos"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optAtivo 
         Caption         =   "Ativos"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      Height          =   1095
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   495
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   2160
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
Attribute VB_Name = "frmRelCustos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'05/05/2005 - Daniel
'
'Projeto: Melhorias para o Centro de Custo
'
'A partir da versão 6.52.0.38 todo campo de Centro de Custo
'estará carregando apenas os Centros que estão ativos no sistema

Private Sub B_Cancela_Click()
  Unload Me
End Sub

Private Sub B_Imprime_Click()
 Dim Str_Rel    As String
 Dim Str1       As String
 Dim strSelecao As String
 
 Call StatusMsg("")

 Rem  Seta Valores e Manda Relatório

 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & "CUSTOS.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 '05/05/2005 - Daniel
 'Adicionado à Seleção
 If optAtivo.Value Then
  strSelecao = "{Centros de Custo.Ativo} = TRUE"
  Rel.SelectionFormula = strSelecao
 End If
 
 If optInativo.Value Then
  strSelecao = "{Centros de Custo.Ativo} = FALSE"
  Rel.SelectionFormula = strSelecao
 End If
 
 If optTodosCentros.Value Then 'Adicionado filtro para trazer todos ...
  strSelecao = "{Centros de Custo.Ativo} = TRUE OR {Centros de Custo.Ativo} = FALSE"
  Rel.SelectionFormula = strSelecao
 End If
 '------------------------------------------------

 If O_Código.Value = True Then Rel.SortFields(0) = "+{Centros de Custo.Código}"
 If O_Nome.Value = True Then Rel.SortFields(0) = "+{Centros de Custo.Nome}"

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
