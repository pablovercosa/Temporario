VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório dos Saldos das Contas Correntes"
   ClientHeight    =   1890
   ClientLeft      =   2325
   ClientTop       =   1935
   ClientWidth     =   4875
   HelpContextID   =   1440
   Icon            =   "RelSaldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1890
   ScaleWidth      =   4875
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   3480
      TabIndex        =   0
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   1800
      Top             =   1080
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
   Begin VB.Label Label1 
      Caption         =   $"RelSaldos.frx":058A
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmRelSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsContas As Recordset
Dim rsTempo As Recordset
Dim rsLançamentos As Recordset

Private Sub B_Imprime_Click()
  Dim Str_Data1 As String
  Dim Str_Data2 As String
  Dim Str_Rel As String
  Dim Str1 As String
  Dim Erro As Integer
  Dim sSql As String
  Dim Total As Double
  Dim Saldo As Double
  Dim Conta As Integer

  Call StatusMsg("Apagando arquivo temporário")

  sSql = "Delete * From ZZZGeral"
  db.Execute sSql
  
  Call StatusMsg("Verificando saldos...")
  
  Total = 0
  Conta = 0
  rsContas.Index = "Código"
  rsLançamentos.Index = "Conta"
Lp1:
  rsContas.Seek ">", Conta
  If rsContas.NoMatch Then GoTo Fim
  Conta = rsContas("Código")

  rsLançamentos.Seek "<", Conta, CDate("01/01/2050"), 999999999#
  Saldo = 0
  If Not rsLançamentos.NoMatch Then
   If rsLançamentos("Conta") = Conta Then Saldo = rsLançamentos("Saldo Atual")
  End If
 
  rsTempo.AddNew
    rsTempo("Texto") = "Saldo conta " + rsContas("Descrição")
    rsTempo("Valor 1") = Saldo
  rsTempo.Update
  
  Total = Total + Saldo
    
  GoTo Lp1
    
Fim:

 Call StatusMsg("Imprimindo, aguarde...")
 
 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 Rem Saída
 If O_Vídeo = True Then Rel.Destination = 0
 If O_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 Str1 = gsReportPath & "SALDOS.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

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
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários", , dbReadOnly)
  Set rsTempo = db.OpenRecordset("ZZZGeral")
End Sub


