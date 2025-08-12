VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelFluxo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Fluxo de Caixa"
   ClientHeight    =   4035
   ClientLeft      =   2220
   ClientTop       =   1605
   ClientWidth     =   8310
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
   Icon            =   "RelFluxodeCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   8310
   Begin VB.Frame Frame4 
      Caption         =   "Tipo "
      Height          =   1095
      Left            =   5400
      TabIndex        =   22
      Top             =   540
      Width           =   2775
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "&Tudo"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "&Realizado"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   300
         Width           =   1005
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "&A Realizar"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Top             =   660
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saldo Inicial "
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   540
      Width           =   5175
      Begin VB.CommandButton B_Busca 
         BackColor       =   &H00C0FFFF&
         Caption         =   "<< &Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Busca soma dos saldos das contas bancárias"
         Top             =   420
         Width           =   2295
      End
      Begin MSMask.MaskEdBox Saldo_Inicial 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   450
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Valor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   450
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   5175
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3675
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   19
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   18
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frame_Completo 
      Caption         =   "Completo - Pagar"
      Height          =   855
      Left            =   5520
      TabIndex        =   14
      Top             =   2490
      Width           =   2655
      Begin VB.OptionButton O_Descrição 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar descrição"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   510
         Width           =   1695
      End
      Begin VB.OptionButton O_Fornecedor 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar fornecedor"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   2820
      TabIndex        =   13
      Top             =   2490
      Width           =   2655
      Begin VB.OptionButton O_Completo 
         Appearance      =   0  'Flat
         Caption         =   "Completo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton O_resumido 
         Appearance      =   0  'Flat
         Caption         =   "Resumido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   2490
      Width           =   2655
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   6
         Top             =   360
         Width           =   1185
      End
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3420
      Width           =   8055
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
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
      Height          =   315
      Left            =   4890
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   6420
      Top             =   3660
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
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "RelFluxodeCaixa.frx":4E95A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   570
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
      Columns(0).Width=   8811
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1950
      TabIndex        =   16
      Top             =   120
      Width           =   6225
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   150
      Width           =   405
   End
End
Attribute VB_Name = "frmRelFluxo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gdtDataIni As Date
Private gdtDataFim As Date

Private Sub B_Busca_Click()
  Dim Total As Double
  Dim Saldo As Double
  Dim Conta As Integer
  Dim rsContas As Recordset
  Dim rsLançamentos As Recordset
  
  Call StatusMsg("Verificando saldos...")
  
  Total = 0
  Conta = 0
  
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários", , dbReadOnly)
  
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
  
  Total = Total + Saldo
    
  GoTo Lp1
    
Fim:
  rsContas.Close
  Set rsContas = Nothing
  rsLançamentos.Close
  Set rsLançamentos = Nothing
  Saldo_Inicial.Text = Total
  Call StatusMsg("")
  
End Sub

'07/01/2004 - mpdea
'Corrigido RT-3075 na sintaxe SQL
'
Private Sub B_Imprime_Click()
  Dim rsCR As Recordset
  Dim rsCP As Recordset
  Dim rsTemp As Recordset
  Dim rsFluxo As Recordset
  Dim sCriteria As String
  Dim sDesc As String
  Dim sDtAnt As String
  Dim sStr As String
  Dim Saldo_Ant As Double
  Dim sSql As String
  Dim sDataIni As String
  Dim sDataFim As String
  Dim sDataHoje As String
  
  Call StatusMsg("")
  
  If IsNull(lblFilial.Caption) Or lblFilial.Caption = "" Then
    cboFilial.Text = "0"
  End If
  
  If Not IsDate(Data_Ini.Text & "") Then
    DisplayMsg "Data incorreta, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text & "") Then
    DisplayMsg "Data incorreta, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual a data final."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Preparando o arquivo temporário... ")
  sSql = "DELETE * FROM Fluxo"
  Call dbTemp.Execute(sSql, dbFailOnError)
  
  Call StatusMsg("Lendo contas a receber ....")
  
  gdtDataIni = CDate(Data_Ini.Text)
  gdtDataFim = CDate(Data_Fim.Text)
  
  sDataIni = gsGetInvDate(Data_Ini.Text)
  sDataFim = gsGetInvDate(Data_Fim.Text)
  sDataHoje = "#" & Format(Date, "mm/dd/yyyy") & "#"
  
  sSql = "SELECT * FROM [Contas a Receber] "
  sSql = sSql & " LEFT JOIN [Cli_For] ON "
  sSql = sSql & " [Contas a Receber].Cliente = [Cli_For].Código  "
  If gsHandleNull(cboFilial.Text) <> "0" Then
    sSql = sSql & " WHERE [Contas a Receber].Filial = " & CInt(cboFilial.Text) & " AND "
  Else
    sSql = sSql & " WHERE "
  End If
  
  If optType(0).Value Then
    sSql = sSql & " ([Contas a Receber].[Valor Recebido] <> 0) "
    sSql = sSql & " AND ([Contas a Receber].[Data Recebimento] >= " & sDataIni
    sSql = sSql & " AND [Contas a Receber].[Data Recebimento] <= " & sDataFim & ") "
    sSql = sSql & " ORDER BY [Contas a Receber].[Data Recebimento], "
  End If
  If optType(1).Value Then
    sSql = sSql & " ([Contas a Receber].[Valor Recebido] = 0) "
    sSql = sSql & " AND ([Contas a Receber].Vencimento >= " & sDataIni
    sSql = sSql & " AND [Contas a Receber].Vencimento <= " & sDataFim & ") "
    sSql = sSql & " ORDER BY [Contas a Receber].Vencimento, "
  End If
  If optType(2).Value Then
    sSql = sSql & " (([Contas a Receber].Vencimento >= " & sDataIni
    sSql = sSql & " AND [Contas a Receber].Vencimento <= " & sDataFim & ") "
    
    sSql = sSql & " OR ([Contas a Receber].[Data Recebimento] >= " & sDataIni
    sSql = sSql & " AND [Contas a Receber].[Data Recebimento] <= " & sDataFim & ")) "
    
    sSql = sSql & " ORDER BY [Contas a Receber].Vencimento, "
  End If
  
  sSql = sSql & " [Contas a Receber].Filial, "
  sSql = sSql & " [Contas a Receber].Contador "
  Set rsCR = db.OpenRecordset(sSql, dbOpenDynaset)
  
  Do While Not rsCR.EOF
    Call Grava_Fluxo_Entrada( _
        CStr(rsCR("Vencimento").Value & ""), _
        CStr(rsCR("Data Recebimento").Value & ""), _
        rsCR("Código").Value & "", _
        rsCR("Nome").Value & "", _
        rsCR("Valor").Value, _
        rsCR("Valor Recebido").Value)
    rsCR.MoveNext
  Loop

  Call StatusMsg("Lendo contas a pagar ....")
  
  'sSql = "SELECT * INTO CPOK FROM [Contas a Pagar] "
  sSql = "SELECT * FROM [Contas a Pagar] "
  sSql = sSql & " LEFT JOIN [Cli_For] ON "
  sSql = sSql & " ([Contas a Pagar].Fornecedor = [Cli_For].Código) "
  If gsHandleNull(cboFilial.Text) <> "0" Then
    sSql = sSql & " WHERE [Contas a Pagar].Filial = " & CInt(cboFilial.Text) & " AND "
  Else
    sSql = sSql & " WHERE "
  End If
  
  If optType(0).Value Then
    sSql = sSql & " ([Contas a Pagar].[Valor Pago] <> 0) "
    sSql = sSql & " AND ([Contas a Pagar].Pagamento >= " & sDataIni
    sSql = sSql & " AND [Contas a Pagar].Pagamento <= " & sDataFim & ") "
    sSql = sSql & " ORDER BY [Contas a Pagar].Pagamento, "
  End If
  If optType(1).Value Then
    sSql = sSql & " ([Contas a Pagar].[Valor Pago] = 0) "
    sSql = sSql & " AND ([Contas a Pagar].Vencimento >= " & sDataIni
    sSql = sSql & " AND [Contas a Pagar].Vencimento <= " & sDataFim & ") "
    sSql = sSql & " ORDER BY [Contas a Pagar].Vencimento, "
  End If
  If optType(2).Value Then
    sSql = sSql & " (([Contas a Pagar].Vencimento >= " & sDataIni
    sSql = sSql & " AND [Contas a Pagar].Vencimento <= " & sDataFim & ") "
    
    sSql = sSql & " OR ([Contas a Pagar].Pagamento >= " & sDataIni
    sSql = sSql & " AND [Contas a Pagar].Pagamento <= " & sDataFim & ")) "
    
    sSql = sSql & " ORDER BY [Contas a Pagar].Vencimento, "
  End If

  sSql = sSql & " [Contas a Pagar].Filial, "
  sSql = sSql & " [Contas a Pagar].Contador "
  Set rsCP = db.OpenRecordset(sSql, dbOpenDynaset)

  Do While Not rsCP.EOF
    If O_Descrição.Value = True Then
      sDesc = rsCP("Descrição") & ""
    Else
      sDesc = rsCP("Nome").Value & ""
    End If
    Call Grava_Fluxo_Saída( _
          CStr(rsCP("Vencimento").Value & ""), _
          CStr(rsCP("Pagamento").Value & ""), _
          rsCP("Código").Value & "", _
          sDesc, _
          rsCP("Valor").Value, _
          rsCP("Valor Pago").Value)
    rsCP.MoveNext
  Loop

  rsCP.Close
  rsCR.Close
  Set rsCR = Nothing
  Set rsCP = Nothing
  
  Call StatusMsg("Atualizando saldos....")
  
  Saldo_Ant = CCur(gsHandleNull(Saldo_Inicial.Text))

  sSql = "SELECT * FROM Fluxo ORDER BY Data, Ordem "
  Set rsFluxo = dbTemp.OpenRecordset(sSql, dbOpenDynaset)
  
  With rsFluxo
    Do While Not .EOF
      .Edit
      .Fields("Saldo Anterior") = Saldo_Ant
      Saldo_Ant = .Fields("Saldo Anterior") + .Fields("Valor Entrada") - .Fields("Valor Saída")
      .Fields("Saldo Final") = Saldo_Ant
      .Update
      .MoveNext
    Loop
  End With
  
  'Para todos as datas com multiplas entradas, faça o primeiro saldo da ordem
  'ser igual ao último do dia - Rel Analítico.
  sSql = "SELECT * FROM Fluxo ORDER BY Data, Ordem DESC "
  Set rsTemp = dbTemp.OpenRecordset(sSql, dbOpenDynaset)
  
  sDtAnt = ""
  
  Do While Not rsTemp.EOF
    If rsTemp("Ordem").Value > 1 Then
      If CStr(rsTemp("Data").Value) <> sDtAnt Then
        sDtAnt = CStr(rsTemp("Data").Value)
        sCriteria = "Data = " & gsGetInvDate(sDtAnt) & " AND Ordem = 1 "
        With rsFluxo
          .FindFirst sCriteria
          If Not .NoMatch Then
            .Edit
            .Fields("Saldo Final").Value = rsTemp("Saldo Final").Value
            .Update
          End If
        End With
      End If
    End If
    rsTemp.MoveNext
  Loop
  
  
  rsFluxo.Close
  rsTemp.Close
  Set rsFluxo = Nothing
  Set rsTemp = Nothing
  
  Call StatusMsg("")

  With Rel1
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
  End With

  If B_Vídeo = True Then
    Rel1.Destination = 0
  Else
    Rel1.Destination = 1
  End If

  If O_Resumido.Value = True Then
    sStr = "FLUXO2.RPT"
  Else
    sStr = "FLUXO1.RPT"
  End If
  Rel1.ReportFileName = gsReportPath & sStr
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel1
  
  sStr = "nome_empresa = '"
  sStr = sStr & "Empresa: " & gsNomeEmpresa & "'"
  Rel1.Formulas(0) = sStr
  
  sStr = "nome_filial = '"
  sStr = sStr & "Filial: " & lblFilial.Caption & "'"
  Rel1.Formulas(1) = sStr
  
  Rem data inicial
  sStr = "data_ini = '"
  sStr = sStr & Data_Ini.Text & "'"
  Rel1.Formulas(2) = sStr
  
  Rem data final
  sStr = "data_fim = '"
  sStr = sStr & Data_Fim.Text & "'"
  Rel1.Formulas(3) = sStr
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  
  
  Rel1.Action = 1
  
End Sub


Private Sub Grava_Fluxo_Entrada( _
        ByVal sDataVenc As String, _
        ByVal sDataReceb As String, _
        ByVal sCod As String, _
        ByVal sDesc As String, _
        ByVal nValorVenc As Currency, _
        ByVal nValorReceb As Currency)
  Dim rsFluxo As Recordset
  Dim nCtOrd As Long
  Dim sSql As String
  Dim dtDate As Date
  Dim nValor As Currency
  Dim dtDataVenc As Date
  Dim dtDataReceb As Date
  
  dtDataVenc = CDate(sDataVenc)
  If IsDate(sDataReceb) Then
    dtDataReceb = CDate(sDataReceb)
    If dtDataVenc < dtDataReceb Then
      dtDate = dtDataReceb
      If dtDate < gdtDataIni Or dtDate > gdtDataFim Then
        Exit Sub
      End If
      If nValorReceb = 0 Then
        nValor = nValorVenc
      Else
        nValor = nValorReceb
      End If
    Else
      dtDate = dtDataVenc
      If dtDate < gdtDataIni Or dtDate > gdtDataFim Then
        Exit Sub
      End If
      nValor = nValorVenc
    End If
  Else
    dtDate = dtDataVenc
    If dtDate < gdtDataIni Or dtDate > gdtDataFim Then
      Exit Sub
    End If
    nValor = nValorVenc
  End If
  
  sSql = "SELECT * FROM Fluxo "
  sSql = sSql & " WHERE Data = " & gsGetInvDate(dtDate)
  sSql = sSql & "ORDER BY Data, Ordem"
  Set rsFluxo = dbTemp.OpenRecordset(sSql, dbOpenDynaset)
 
  With rsFluxo
    If Not rsFluxo.EOF Then
      rsFluxo.MoveLast
      nCtOrd = rsFluxo("Ordem") + 1
    Else
      nCtOrd = 1
    End If
    .AddNew
    .Fields("Data") = dtDate
    .Fields("Ordem") = nCtOrd
    .Fields("Cód Entrada") = IIf(IsNumeric(sCod), sCod, 0)
    .Fields("Desc Entrada") = Left(sDesc + Space(50), 50)
    .Fields("Valor Entrada") = nValor
    .Update
  End With
  
  rsFluxo.Close
  Set rsFluxo = Nothing
  
End Sub

Private Sub Grava_Fluxo_Saída( _
        ByVal sDataVenc As String, _
        ByVal sDataPagto As String, _
        ByVal sCod As String, _
        ByVal sDesc As String, _
        ByVal nValorVenc As Currency, _
        ByVal nValorPagto As Currency)
  Dim rsFluxo As Recordset
  Dim nCtOrd As Long
  Dim sSql As String
  Dim dtDate As Date
  Dim nValor As Currency
  Dim dtDataVenc As Date
  Dim dtDataPagto As Date

  dtDataVenc = CDate(sDataVenc)
  If IsDate(sDataPagto) Then
    dtDataPagto = CDate(sDataPagto)
    If dtDataVenc < dtDataPagto Then
      dtDate = dtDataPagto
      If dtDate < gdtDataIni Or dtDate > gdtDataFim Then
        Exit Sub
      End If
      nValor = nValorPagto
    Else
      dtDate = dtDataVenc
      If dtDate < gdtDataIni Or dtDate > gdtDataFim Then
        Exit Sub
      End If
      nValor = nValorVenc
    End If
  Else
    dtDate = dtDataVenc
    If dtDate < gdtDataIni Or dtDate > gdtDataFim Then
      Exit Sub
    End If
    nValor = nValorVenc
  End If
  
  sSql = "SELECT * FROM Fluxo "
  sSql = sSql & " WHERE Data = " & gsGetInvDate(dtDate)
  sSql = sSql & " ORDER BY Data, Ordem"
  Set rsFluxo = dbTemp.OpenRecordset(sSql, dbOpenDynaset)
  
  With rsFluxo
    If Not rsFluxo.EOF Then
      rsFluxo.MoveLast
      nCtOrd = rsFluxo("Ordem") + 1
    Else
      nCtOrd = 1
    End If
    .AddNew
    .Fields("Data") = dtDate
    .Fields("Ordem") = nCtOrd
    .Fields("Cód Saída") = IIf(IsNumeric(sCod), sCod, 0)
    .Fields("Desc Saída") = Left(sDesc + Space(50), 50)
    .Fields("Valor Saída") = nValor
    .Update
  End With

  rsFluxo.Close
  Set rsFluxo = Nothing
 
End Sub

'Private Sub cboFilial_Change()
'  If cboFilial.Text = "0" Then
'    lblFilial.Caption = "Todas"
'  End If
'End Sub
'
'Private Sub cboFilial_CloseUp()
'  cboFilial.Text = cboFilial.Columns(1).Text
'  lblFilial.Caption = cboFilial.Columns(0).Text
'  Call StatusMsg("")
'End Sub
'
'Private Sub cboFilial_GotFocus()
'  Call StatusMsg(LoadResString(50))
'End Sub
'
'Private Sub cboFilial_LostFocus()
'  Call StatusMsg("")
'End Sub

Private Sub cboFilial_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboFilial_CloseUp()
  With cboFilial
    lblFilial.Caption = .Columns("Nome").Text
    .Text = .Columns("Filial").Text
  End With
End Sub

Private Sub cboFilial_KeyPress(KeyAscii As Integer)
  If Not cboFilial.DroppedDown Then
    KeyAscii = gnLimitKeyPress(cboFilial, 2, KeyAscii, True)
  End If
End Sub

Private Sub cboFilial_LostFocus()
  Call StatusMsg("")
  If cboFilial.Text = "0" Or cboFilial.Text = "" Then
    cboFilial.Text = "0"
    lblFilial.Caption = "Todas"
  Else
    lblFilial.Caption = gsGetNameFilial(Val(cboFilial.Text))
  End If
End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Data1.DatabaseName = gsQuickDBFileName
  optType_Click (0)
End Sub

Private Sub O_Completo_Click()
 O_Fornecedor.Enabled = True
 O_Descrição.Enabled = True
End Sub

Private Sub O_Resumido_Click()
 O_Fornecedor.Enabled = False
 O_Descrição.Enabled = False
End Sub

Private Sub optType_Click(Index As Integer)
  Select Case Index
    Case 0
      Frame6.Caption = "Período (Data de Vencimento)"
    Case 1
      Frame6.Caption = "Período (Data de Pagamento e Recebimento)"
    Case 2
      Frame6.Caption = "Período (Data de Vencimento)"
  End Select
End Sub
