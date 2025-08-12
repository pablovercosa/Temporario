VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmManeMode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atalhos"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "ManeMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabSecao 
      Height          =   6210
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Primeiros Passos"
      TabPicture(0)   =   "ManeMode.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgSecao(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIcone(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblIcone(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblIcone(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblIcone(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblIcone(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblIcone(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblIcone(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblIcone(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblIcone(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblIcone(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblIcone(32)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Elaborando Preços"
      TabPicture(1)   =   "ManeMode.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblIcone(10)"
      Tab(1).Control(1)=   "lblIcone(11)"
      Tab(1).Control(2)=   "lblIcone(12)"
      Tab(1).Control(3)=   "lblIcone(13)"
      Tab(1).Control(4)=   "lblIcone(14)"
      Tab(1).Control(5)=   "lblIcone(15)"
      Tab(1).Control(6)=   "lblIcone(16)"
      Tab(1).Control(7)=   "lblIcone(17)"
      Tab(1).Control(8)=   "imgSecao(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Entradas e &Saídas"
      TabPicture(2)   =   "ManeMode.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblIcone(27)"
      Tab(2).Control(1)=   "lblIcone(26)"
      Tab(2).Control(2)=   "lblIcone(25)"
      Tab(2).Control(3)=   "lblIcone(24)"
      Tab(2).Control(4)=   "lblIcone(23)"
      Tab(2).Control(5)=   "lblIcone(22)"
      Tab(2).Control(6)=   "lblIcone(21)"
      Tab(2).Control(7)=   "lblIcone(20)"
      Tab(2).Control(8)=   "lblIcone(19)"
      Tab(2).Control(9)=   "lblIcone(18)"
      Tab(2).Control(10)=   "imgSecao(2)"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "&C.Pagar e C.Receber"
      TabPicture(3)   =   "ManeMode.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ActiveBar1"
      Tab(3).Control(1)=   "lblIcone(31)"
      Tab(3).Control(2)=   "lblIcone(30)"
      Tab(3).Control(3)=   "lblIcone(29)"
      Tab(3).Control(4)=   "lblIcone(28)"
      Tab(3).Control(5)=   "imgSecao(3)"
      Tab(3).ControlCount=   6
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1155
         Index           =   32
         Left            =   6150
         TabIndex        =   33
         Top             =   3960
         Width           =   1155
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   -68010
         Top             =   5475
         _ExtentX        =   847
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Bands           =   "ManeMode.frx":0D3A
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1275
         Index           =   31
         Left            =   -71790
         TabIndex        =   32
         Top             =   4110
         Width           =   1620
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1605
         Index           =   30
         Left            =   -69915
         TabIndex        =   31
         Top             =   2970
         Width           =   1830
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1605
         Index           =   29
         Left            =   -74010
         TabIndex        =   30
         Top             =   3015
         Width           =   1830
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Index           =   28
         Left            =   -71745
         TabIndex        =   29
         Top             =   1425
         Width           =   1575
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1245
         Index           =   27
         Left            =   -68820
         TabIndex        =   28
         Top             =   4455
         Width           =   1230
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1185
         Index           =   26
         Left            =   -68400
         TabIndex        =   27
         Top             =   5925
         Width           =   1245
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1140
         Index           =   25
         Left            =   -71655
         TabIndex        =   26
         Top             =   4530
         Width           =   1200
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1125
         Index           =   24
         Left            =   -72900
         TabIndex        =   25
         Top             =   4560
         Width           =   990
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1080
         Index           =   23
         Left            =   -74310
         TabIndex        =   24
         Top             =   4590
         Width           =   1065
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1125
         Index           =   22
         Left            =   -70200
         TabIndex        =   23
         Top             =   4605
         Width           =   1140
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1110
         Index           =   21
         Left            =   -69510
         TabIndex        =   22
         Top             =   2055
         Width           =   1155
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1065
         Index           =   20
         Left            =   -70935
         TabIndex        =   21
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   900
         Index           =   19
         Left            =   -72150
         TabIndex        =   20
         Top             =   2145
         Width           =   945
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   900
         Index           =   18
         Left            =   -73725
         TabIndex        =   19
         Top             =   2115
         Width           =   1320
      End
      Begin VB.Image imgSecao 
         Height          =   5700
         Index           =   3
         Left            =   -74910
         Picture         =   "ManeMode.frx":5C30
         Top             =   420
         Width           =   7800
      End
      Begin VB.Image imgSecao 
         Height          =   5700
         Index           =   2
         Left            =   -74865
         Picture         =   "ManeMode.frx":8F00
         Top             =   405
         Width           =   7800
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1155
         Index           =   9
         Left            =   4815
         TabIndex        =   18
         Top             =   3990
         Width           =   1155
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1125
         Index           =   8
         Left            =   3450
         TabIndex        =   17
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1080
         Index           =   7
         Left            =   2085
         TabIndex        =   16
         Top             =   4005
         Width           =   1125
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1185
         Index           =   6
         Left            =   765
         TabIndex        =   15
         Top             =   3945
         Width           =   1125
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1155
         Index           =   5
         Left            =   6645
         TabIndex        =   14
         Top             =   2325
         Width           =   1065
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1230
         Index           =   4
         Left            =   5400
         TabIndex        =   13
         Top             =   2265
         Width           =   1095
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1155
         Index           =   3
         Left            =   4140
         TabIndex        =   12
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   2
         Left            =   2700
         TabIndex        =   11
         Top             =   2325
         Width           =   1260
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   1
         Left            =   1710
         TabIndex        =   10
         Top             =   2265
         Width           =   930
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1065
         Index           =   0
         Left            =   450
         TabIndex        =   9
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Image imgSecao 
         Height          =   5700
         Index           =   0
         Left            =   150
         Picture         =   "ManeMode.frx":C9DC
         Top             =   450
         Width           =   7800
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1110
         Index           =   10
         Left            =   -74070
         TabIndex        =   8
         Top             =   2445
         Width           =   1260
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1065
         Index           =   11
         Left            =   -72495
         TabIndex        =   7
         Top             =   2490
         Width           =   1380
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   990
         Index           =   12
         Left            =   -70830
         TabIndex        =   6
         Top             =   2505
         Width           =   1290
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   870
         Index           =   13
         Left            =   -69285
         TabIndex        =   5
         Top             =   2655
         Width           =   1500
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   930
         Index           =   14
         Left            =   -74130
         TabIndex        =   4
         Top             =   4110
         Width           =   1365
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   975
         Index           =   15
         Left            =   -72465
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   1020
         Index           =   16
         Left            =   -70905
         TabIndex        =   2
         Top             =   4065
         Width           =   1485
      End
      Begin VB.Label lblIcone 
         BackStyle       =   0  'Transparent
         Height          =   900
         Index           =   17
         Left            =   -69165
         TabIndex        =   1
         Top             =   4110
         Width           =   1320
      End
      Begin VB.Image imgSecao 
         Height          =   5700
         Index           =   1
         Left            =   -74895
         Picture         =   "ManeMode.frx":11D5B
         Top             =   405
         Width           =   7800
      End
   End
End
Attribute VB_Name = "frmManeMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gbError As Boolean
Private gsFileName As String

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Dim F As Form
  
  If frmMain.ActiveBar1.Tools(Tool.Name).Enabled = False Then
    Beep
    Exit Sub
  End If
  
  Select Case Tool.Name
  '
  ' CONTAS A PAGAR
    Case "miCPLancamentos"
      Set F = New frmLancaCPagar
      F.Show
    Case "miCPGeracao"
      Set F = New frmGeraPagar
      F.Show
    Case "miCPManutencao"
      Set F = New frmManContasPagar
      F.Show
    Case "miCPResetCP"
      Set F = New frmApagaPagas
      F.Show
  '
  ' CONTAS A RECEBER
    Case "miCRLancamentos"
      Set F = New frmLancaCReceber
      F.Show
    Case "miCRManutencao"
       Set F = New frmManContasReceber
      F.Show
    Case "miCRResetCR"
      Set F = New frmApagaRecebidas
      F.Show
  '
  ' CHEQUES PRÉ-
    Case "miCRLancPreDatados"
      Set F = New frmChequesPre
      F.Show
    Case "miCRManutPreDatados"
      Set F = New frmManCheques
      F.Show
    Case "miCRResetPreDatados"
      Set F = New frmApagaCheques
      F.Show
  '
  ' CARTÕES DE CRÉDITO
    Case "miCRLancCartoesCredito"
      Set F = New frmLancaCCredito
      F.Show
    Case "miManutCartoesCredito"
      Set F = New frmManCartoes
      F.Show
    Case "miCRResetCartoesCredito"
      Set F = New frmApagaCartoes
      F.Show
  '
  ' CONTAS DE CLIENTES (VULGO CADERNO)
    Case "miCRLancContasClientes"
      Set F = New frmLancaContaCliente
      F.Show
    Case "miCRManutContasClientes"
      Set F = New frmManContas
      F.Show
    Case "miCRResetContasClientes"
      Set F = New frmApagaContaCliente
      F.Show
  
  '
  ' FINANCEIRO
    Case "miRepFinCaixas"
      Set F = New frmRelCaixa
      F.Show
    Case "miRepFinLancBancarios"
      Set F = New frmRelLancamentos
      F.Show
    Case "miRepFinSaldosContas"
      Set F = New frmRelSaldos
      F.Show
    Case "miRepFinDiario1"
      Set F = New frmRelFinanc1
      F.Show
    Case "miRepFinDiario2"
      Set F = New frmRelFinanc2
      F.Show
    Case "miRepFinLucratividade"
      Set F = New frmRelLucratividade
      F.Show
    Case "miRepFinGeral"
      Set F = New frmRelFinGeral
      F.Show
  
  
  ' REPORTS FINANCEIROS - CONTAS A PAGAR
    Case "miRepFinCPPagarVencimento"
      Set F = New frmRelPagar1
      F.Show
    Case "miRepFinCPPagarFornecedor"
      Set F = New frmRelPagar2
      F.Show
    Case "miRepFinCPPagarTodasFiliais"
      Set F = New frmRelPagar3
      F.Show
    Case "miRepFinCPPagarCentroCusto"
      Set F = New frmRelPagar4
      F.Show
    Case "miRepFinCPPagasFornecedor"
      Set F = New frmRelPagas3
      F.Show
    Case "miRepFinCPPagasPagamento"
      Set F = New frmRelPagas2
      F.Show
    Case "miRepFinCPPagasCentroCusto"
      Set F = New frmRelPagas1
      F.Show
  '
  ' REPORTS FINANCEIROS - CONTAS A RECEBER
    Case "miRepFinCRReceberVencimento"
      Set F = New frmRelReceber1
      F.Show
    Case "miRepFinCRReceberCliente"
      Set F = New frmRelReceber2
      F.Show
    Case "miRepFinCRPosicaoCliente"
      '15/06/2004 - Daniel
      'Rel. Solicitado pela TV Shopping mas aproveitado para os demais clientes
      Set F = New frmRelPosicaoFinanceiraCliente
      F.Show
    Case "miRepFinCRReceberVendedor"
      Set F = New frmRelRecebidas1
      F.Show
    Case "miRepFinCRRecebidasRecebimento"
      Set F = New frmRelRecebidas2
      F.Show
    Case "miRepFinCRRecebidasVendedor"
      Set F = New frmRelRecebidas1
      F.Show
    Case "miRepFinCRRecebidasCliente"
      Set F = New frmRelRecebidas3
      F.Show
    Case "miRepFinCRChequesPre"
      Set F = New frmRelCheque
      F.Show
    Case "miRepFinCRCartoesCredito"
      Set F = New frmRelCartoes
      F.Show
    Case "miRepFinCRContasClientes"
      Set F = New frmRelContaCliente
      F.Show
    Case "miRepCRBoletos"
      Set F = New frmImprimeBoletos
      F.Show
    Case "miRepCRCarnes"
      Set F = New frmImprimeCarnes
      F.Show
  '
  ' REPORTS FINANCEIROS - FLUXO DE CAIXA
    Case "miRepFinFluxoCaixa"
      Set F = New frmRelFluxo
      F.Show
  '
  End Select
  '
End Sub

'-----------------------------------------------------------------------------------
'20/01/2003 - mpdea
'Não exibe erro ao não encontrar a imagem correspondente ao convênio
'Permanece a imagem atual (controle)
Private Sub Form_Activate()
'  If gbError = True Then
'    gsTitle = LoadResString(201)
'    gsMsg = "Erro ao carregar arquivo de imagens."
'    gsMsg = gsMsg & vbCrLf & "Arquivo ausente: " & gsFileName
'    gnStyle = vbOKOnly + vbCritical
'    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'    Unload Me
'  End If
End Sub

Private Sub Form_Load()
  Dim sFileName As String
  Dim nI As Integer
  
  Call CenterForm(Me)
'  gbError = False
  
  For nI = 1 To 4
    sFileName = App.Path & "\Imagens\Image" & CStr(nI) & CStr(gnNumConvenio) & ".jpg"
    If Dir(sFileName) <> "" Then
      imgSecao(nI - 1).Picture = LoadPicture(sFileName)
    Else
'      gbError = True
'      gsFileName = sFileName
    End If
  Next nI
End Sub
'-----------------------------------------------------------------------------------

Private Sub lblIcone_Click(Index As Integer)
  Dim Tool As ActiveBarLibraryCtl.Tool
  Dim F As Form
  
  Select Case Index
    Case 0
      Set Tool = frmMain.ActiveBar1.Tools("miParEmpresa")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 5, 18, 23
      Set Tool = frmMain.ActiveBar1.Tools("miCadProdutos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 1, 19, 24
      Set Tool = frmMain.ActiveBar1.Tools("miCadFuncionarios")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 2, 20, 25
      Set Tool = frmMain.ActiveBar1.Tools("miCadCliFor")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 3
      Set Tool = frmMain.ActiveBar1.Tools("miCadClasses")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 4
      Set Tool = frmMain.ActiveBar1.Tools("miCadSubClasses")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 8
      Set Tool = frmMain.ActiveBar1.Tools("miCadClassificacaoFiscal")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 6
      If gbGrade Then
        Set Tool = frmMain.ActiveBar1.Tools("miCadCores")
        If Tool.Enabled = False Then
          Beep
          Exit Sub
        End If
        Call frmMain.ActiveBar1_Click(Tool)
      Else
        Beep
      End If
    Case 7
      If gbGrade Then
        Set Tool = frmMain.ActiveBar1.Tools("miCadTamanhos")
        If Tool.Enabled = False Then
          Beep
          Exit Sub
        End If
        Call frmMain.ActiveBar1_Click(Tool)
      Else
        Beep
      End If
    Case 32
      Set Tool = frmMain.ActiveBar1.Tools("miCadProdutos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      gsCodProduto = ""
      '31/08/2006 - Anderson
      'Implementação de pesquisa avançada na tela de consulta do produto
      'Set F = New frmConsultaProd
      'F.Show vbModal
      frmConsultaProd.Show
      'Set F = Nothing
    Case 9, 10
      Set Tool = frmMain.ActiveBar1.Tools("miPreCriaRecriaTabela")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 11
      Set Tool = frmMain.ActiveBar1.Tools("miPreResetTabPrecos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 12
      Set Tool = frmMain.ActiveBar1.Tools("miPreCopiaIndice")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 13
      Set Tool = frmMain.ActiveBar1.Tools("miPreCopiaValor")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 14
      Set Tool = frmMain.ActiveBar1.Tools("miPreConfiguraTabPrecos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 16
      Set Tool = frmMain.ActiveBar1.Tools("miPreAlteraPrecos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 15
      Set Tool = frmMain.ActiveBar1.Tools("miPreDigitacao")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 17
      Set Tool = frmMain.ActiveBar1.Tools("miPreImprimirPrecos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 21
      Set Tool = frmMain.ActiveBar1.Tools("miMovEntradas")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 22
      Set Tool = frmMain.ActiveBar1.Tools("miMovVendaRapida")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 27
      Set Tool = frmMain.ActiveBar1.Tools("miMovSaidas")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 28
      Set Tool = frmMain.ActiveBar1.Tools("miCadCentroCusto")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 29
      Set Tool = frmMain.ActiveBar1.Tools("miCPLancamentos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
    Case 30
      Set Tool = frmMain.ActiveBar1.Tools("miCRLancamentos")
      If Tool.Enabled = False Then
        Beep
        Exit Sub
      End If
      Call frmMain.ActiveBar1_Click(Tool)
  End Select
End Sub

Private Sub lblIcone_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Index
    Case 31
      ActiveBar1.Bands("mnuRepFinanceiros").TrackPopup -1, -1
    Case 29
      ActiveBar1.Bands("mnuCP").TrackPopup -1, -1
    Case 30
      ActiveBar1.Bands("mnuCR").TrackPopup -1, -1
  End Select
  
End Sub

Private Sub lblIcone_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblIcone(Index).MouseIcon = LoadResPicture(101, vbResIcon)
  lblIcone(Index).MousePointer = 99
End Sub
