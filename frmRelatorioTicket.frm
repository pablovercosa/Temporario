VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelatorioTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmRelatorioTicket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExibirDadosCliente 
      Caption         =   "Exibir os dados do cliente"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1545
      Begin VB.OptionButton O_Arquivo 
         Caption         =   "Arquivo RTF"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   550
         Width           =   1095
      End
      Begin VB.OptionButton O_vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   5040
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   2520
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
Attribute VB_Name = "frmRelatorioTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sequencia As Long
Public Filial As Byte

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
    
    .Destination = IIf(O_vídeo.Value, crptToWindow, IIf(O_Impressora.Value, crptToPrinter, crptToWindow))
    .ReportFileName = gsReportPath & "Ticket.rpt"
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
    
    .SelectionFormula = "{Saídas.Filial} = " & Filial & " AND {Saídas.Sequência} = " & Sequencia

    .Formulas(0) = "oculta_dados_cliente = " & IIf(chkExibirDadosCliente.Value = vbChecked, 0, 1) & ""
    
    'traz as parcelas para o campo observações
    'pablo 28/07/2023
    Dim sql As String
    sql = "SELECT [Vencimento], [Valor], [Parcela] FROM [Contas a Receber] "
    sql = sql & "WHERE [Filial] = " & Filial & " AND [Sequência] = " & Sequencia
    sql = sql & " ORDER BY [Parcela] ASC;"
    Dim rsParcelas As Recordset
    Set rsParcelas = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim pObservacao As String
    pObservacao = ""
    If Not (rsParcelas.BOF And rsParcelas.EOF) Then
      rsParcelas.MoveLast
      rsParcelas.MoveFirst
      Dim i As Integer
      For i = 0 To rsParcelas.RecordCount - 1
        pObservacao = pObservacao & "(" & CStr(rsParcelas.Fields("Parcela").Value) & ") "
        pObservacao = pObservacao & CStr(rsParcelas.Fields("Vencimento").Value) & " - R$ "
    ' Formate o valor com duas casas decimais usando a função Format
    ' Mauro 13/09/2023
        pObservacao = pObservacao & Format(rsParcelas.Fields("Valor").Value, "#,###0.00") & vbLf
    '   pObservacao = pObservacao & CStr(rsParcelas.Fields("Valor").Value) & vbLf
        rsParcelas.MoveNext
      Next
    End If
   rsParcelas.Close
   Set rsParcelas = Nothing
   If Trim(pObservacao) = "" Then pObservacao = "Nenhuma observação"
   .ParameterFields(0) = "pObservacao;" & Trim(pObservacao) & ";true"
   
   'define a condicional para existência ou não de orçamento
   'pablo 16/01/2025
   sql = ""
   sql = sql & "SELECT [Operações Saída].Tipo "
   sql = sql & "FROM Saídas INNER JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
   sql = sql & "WHERE (((Saídas.Sequência)=" & Sequencia & "));"
   Dim rsTipoOperacao As Recordset
   Set rsTipoOperacao = db.OpenRecordset(sql, dbOpenSnapshot)
   Dim isOrcamento As Boolean
   If Not (rsTipoOperacao.BOF And rsTipoOperacao.EOF) Then
      rsTipoOperacao.MoveLast
      rsTipoOperacao.MoveFirst
      isOrcamento = CStr(rsTipoOperacao.Fields("Tipo").Value) = "O"
   Else
      isOrcamento = False
   End If
   rsTipoOperacao.Close
   Set rsTipoOperacao = Nothing
   .ParameterFields(1) = "pOrcamento;" & isOrcamento & ";true"

    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
    
    .PrintFileName = "Sequencia" + CStr(Sequencia) + "_" + Format(Now, "yyyymmdd_hhmmss") + ".htm"
    .PrintFileType = crptHTML32Ext
    
    
    
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
  
  Dim bln_ret As Boolean
  bln_ret = CBool(GetSetting("QuickStore", "RelTicket", "ExibirDadosCliente", "False"))
  chkExibirDadosCliente.Value = IIf(bln_ret, vbChecked, vbUnchecked)

  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao abrir a tela: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  On Error GoTo ErrHandler
  
  Call SaveSetting("QuickStore", "RelTicket", "ExibirDadosCliente", chkExibirDadosCliente.Value)

  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao fechar a tela: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

