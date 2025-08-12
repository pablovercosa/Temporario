VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmOpSaida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"OpSaida.frx":0000
   ClientHeight    =   7320
   ClientLeft      =   1350
   ClientTop       =   1350
   ClientWidth     =   13890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   HelpContextID   =   1200
   Icon            =   "OpSaida.frx":00E8
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7320
   ScaleWidth      =   13890
   Begin VB.CheckBox chkSomaIPITotalBC 
      Appearance      =   0  'Flat
      Caption         =   "Somar IPI ao total da BC para ICMS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   510
      TabIndex        =   72
      Top             =   6540
      Width           =   2865
   End
   Begin VB.CheckBox chkSomaIPITotalNF 
      Appearance      =   0  'Flat
      Caption         =   "Somar IPI ao total da nota"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   510
      TabIndex        =   71
      Top             =   6270
      Width           =   2745
   End
   Begin VB.CommandButton cmd_pesquisarOperSaida 
      Height          =   435
      Left            =   8310
      Picture         =   "OpSaida.frx":4EA42
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   60
      Width           =   555
   End
   Begin VB.TextBox txtCSOSN 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   12945
      MaxLength       =   3
      TabIndex        =   52
      ToolTipText     =   "Obrigatório somente no caso de se encaixar como Simples Nacional Ex: 101 ou 102 ou 400"
      Top             =   90
      Width           =   885
   End
   Begin VB.Data datGrupoFiscal 
      Caption         =   "datGrupoFiscal"
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
      Height          =   375
      Left            =   8130
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM GrupoFiscal ORDER BY Nome"
      Top             =   6870
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Caption         =   "Impostos"
      Height          =   1695
      Left            =   60
      TabIndex        =   44
      Top             =   5250
      Width           =   8775
      Begin VB.ComboBox cboMotivoDesoneracao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "OpSaida.frx":4F684
         Left            =   3480
         List            =   "OpSaida.frx":4F691
         TabIndex        =   56
         Top             =   435
         Width           =   5055
      End
      Begin VB.CheckBox chkSomaIcmsRetidoTotalNota 
         Appearance      =   0  'Flat
         Caption         =   "Somar ICMS ST ao total da nota"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   510
         Width           =   2745
      End
      Begin VB.TextBox txtIcmsFrete 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox ICM 
         Appearance      =   0  'Flat
         Caption         =   "Calcula ICM"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox IPI 
         Appearance      =   0  'Flat
         Caption         =   "Calcula IPI"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   780
         Width           =   1215
      End
      Begin VB.CheckBox O_Calcula_ISS 
         Appearance      =   0  'Flat
         Caption         =   "Calcula ISS"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2580
         TabIndex        =   21
         Top             =   810
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox O_Base_IPI 
         Appearance      =   0  'Flat
         Caption         =   "Base ICM deve somar IPI"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3990
         TabIndex        =   26
         Top             =   810
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkIpiTot 
         Appearance      =   0  'Flat
         Caption         =   "Calcula IPI somente p/Total"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkIcmFrete 
         Appearance      =   0  'Flat
         Caption         =   "Calcula ICM  Frete"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6030
         TabIndex        =   23
         Top             =   1095
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblMotivoDesoneracao 
         Caption         =   "Motivo Desoneração"
         Enabled         =   0   'False
         Height          =   225
         Left            =   3480
         TabIndex        =   55
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Enabled         =   0   'False
         Height          =   195
         Left            =   8250
         TabIndex        =   45
         Top             =   1110
         Visible         =   0   'False
         Width           =   165
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Movimentação"
      Height          =   4665
      Left            =   60
      TabIndex        =   43
      Top             =   540
      Width           =   8775
      Begin VB.Frame fraAcerta 
         Caption         =   "Acerta Empréstimo de Entrada"
         Height          =   675
         Left            =   150
         TabIndex        =   64
         Top             =   3960
         Width           =   2745
         Begin VB.CheckBox chkAcertaEmprestimo 
            Appearance      =   0  'Flat
            Caption         =   "Sim"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   150
            TabIndex        =   65
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.Frame FraNFCe 
         Caption         =   "NFCe"
         Height          =   735
         Left            =   150
         TabIndex        =   62
         Top             =   3045
         Width           =   2745
         Begin VB.CheckBox ChkPermiteDadosCliente 
            Appearance      =   0  'Flat
            Caption         =   "Permite Emitir Dados do Cliente"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   135
            TabIndex        =   63
            Top             =   300
            Width           =   2550
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   2985
         TabIndex        =   57
         Top             =   3045
         Width           =   5685
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   180
            Picture         =   "OpSaida.frx":4F6EA
            ScaleHeight     =   1410
            ScaleWidth      =   1020
            TabIndex        =   60
            Top             =   135
            Width           =   1020
         End
         Begin VB.CheckBox chkInformante 
            Appearance      =   0  'Flat
            Caption         =   "Operação de venda realizada pelo Celular/Tablet (App) para compradores de OUTROS estados"
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1275
            TabIndex        =   59
            Top             =   975
            Width           =   4260
         End
         Begin VB.CheckBox chkEmitirNFManualmente 
            Appearance      =   0  'Flat
            Caption         =   "Operação de venda realizada pelo Celular/Tablet (App) para compradores do MESMO estado"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1275
            TabIndex        =   58
            ToolTipText     =   "Permite alterar o nº da nota fiscal de saída"
            Top             =   315
            Width           =   4260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "ou"
            Height          =   195
            Left            =   3120
            TabIndex        =   61
            Top             =   720
            Width           =   180
         End
      End
      Begin VB.TextBox txtModeloDocumentoFiscal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6345
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2730
         Width           =   735
      End
      Begin VB.CheckBox chkExibirTelaNumeroDocumento 
         Appearance      =   0  'Flat
         Caption         =   "Exibir tela para preenchimento do número de documento (CPF ou CNPJ)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4425
         TabIndex        =   19
         Top             =   2355
         Width           =   3975
      End
      Begin VB.CheckBox chkSomarProdutos 
         Appearance      =   0  'Flat
         Caption         =   "Somar Produtos ao Total da Nota"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   12
         ToolTipText     =   "Ao marcar este campo, esta Op. de Saída não calculará os impostos sobre Serviço: PIS, COFINS, CSLL."
         Top             =   2715
         Width           =   3255
      End
      Begin VB.Frame fraSadigWeb 
         Caption         =   "Sadig Web"
         Height          =   615
         Left            =   4425
         TabIndex        =   49
         Top             =   1680
         Width           =   4245
         Begin VB.ComboBox cboSadigWebTipo 
            Height          =   315
            ItemData        =   "OpSaida.frx":500DC
            Left            =   1320
            List            =   "OpSaida.frx":500EC
            TabIndex        =   18
            Top             =   210
            Width           =   2655
         End
         Begin VB.Label Label58 
            Caption         =   "Tipo de Saída"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
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
         Height          =   345
         Left            =   7410
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Op_Saída"
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkAlteraStatusPedidoWeb 
         Appearance      =   0  'Flat
         Caption         =   "Altera status do pedido web para recebido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4425
         TabIndex        =   16
         Top             =   1170
         Width           =   3855
      End
      Begin VB.CheckBox chkSomaSeguro 
         Appearance      =   0  'Flat
         Caption         =   "Somar seguro oriundo da web no total a receber"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4425
         TabIndex        =   15
         Top             =   900
         Width           =   3885
      End
      Begin VB.CheckBox chkComissaoServicos 
         Appearance      =   0  'Flat
         Caption         =   "Comissão sobre Vendas (Serviços)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   11
         ToolTipText     =   "Ao marcar este campo, esta Op. de Saída não calculará os impostos sobre Serviço: PIS, COFINS, CSLL."
         Top             =   2445
         Width           =   3255
      End
      Begin VB.CheckBox chkValidade 
         Appearance      =   0  'Flat
         Caption         =   "Existe Validade para Desativar as Reservas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4425
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox chkEntregas 
         Appearance      =   0  'Flat
         Caption         =   "Trabalhar com entregas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4425
         TabIndex        =   13
         Top             =   225
         Width           =   2175
      End
      Begin VB.CheckBox Senha 
         Appearance      =   0  'Flat
         Caption         =   "Exige senha do gerente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   1710
         Width           =   2535
      End
      Begin VB.CheckBox chkTelaObsTransp 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar Tela de Observações/ Transportadora na Emissão Nota/ Ticket"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   165
         TabIndex        =   10
         Top             =   1935
         Width           =   3720
      End
      Begin VB.CheckBox chkSomaFrete 
         Appearance      =   0  'Flat
         Caption         =   "Somar frete no total a receber"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   8
         Top             =   1425
         Width           =   3495
      End
      Begin VB.CheckBox Estoque 
         Appearance      =   0  'Flat
         Caption         =   "Diminui estoque"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   225
         Width           =   1575
      End
      Begin VB.CheckBox Dinheiro 
         Appearance      =   0  'Flat
         Caption         =   "Movimenta dinheiro (sai/entra no caixa ou contas a receber)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   165
         TabIndex        =   5
         Top             =   495
         Width           =   3705
      End
      Begin VB.CheckBox Nota 
         Appearance      =   0  'Flat
         Caption         =   "Emite nota"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   1170
         Width           =   1095
      End
      Begin VB.CheckBox Comissão 
         Appearance      =   0  'Flat
         Caption         =   "Gera comissão ao vendedor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   6
         Top             =   900
         Width           =   2775
      End
      Begin SSDataWidgets_B.SSDBCombo cboOpEntrega 
         Bindings        =   "OpSaida.frx":5011F
         DataSource      =   "Data2"
         Height          =   300
         Left            =   4695
         TabIndex        =   14
         Top             =   540
         Width           =   795
         DataFieldList   =   "Nome"
         MaxDropDownItems=   16
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
         BackColorOdd    =   12648384
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   6959
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1879
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1402
         _ExtentY        =   529
         _StockProps     =   93
         ForeColor       =   -2147483630
         BackColor       =   16777215
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Modelo documento fiscal"
         Height          =   195
         Index           =   0
         Left            =   4425
         TabIndex        =   51
         Top             =   2775
         Width           =   1755
      End
      Begin VB.Label lblOpEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5505
         TabIndex        =   46
         Top             =   510
         Width           =   3165
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ticket"
      Height          =   1695
      Left            =   8910
      TabIndex        =   42
      Top             =   5250
      Width           =   4965
      Begin VB.ComboBox Combo_Tickets 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   38
         Top             =   675
         Width           =   2250
      End
      Begin VB.OptionButton O_Específico 
         Appearance      =   0  'Flat
         Caption         =   "&Sempre usar este ticket:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2490
         TabIndex        =   37
         Top             =   405
         Width           =   2115
      End
      Begin VB.OptionButton O_Perguntar 
         Appearance      =   0  'Flat
         Caption         =   "&Perguntar qual ticket usar no momento da impressão"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   36
         Top             =   345
         Value           =   -1  'True
         Width           =   2190
      End
   End
   Begin VB.TextBox Código_Fiscal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   9735
      MaxLength       =   4
      TabIndex        =   2
      ToolTipText     =   "Código de operação fiscal  EX: 5102 para venda estadual"
      Top             =   90
      Width           =   885
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   1
      Top             =   90
      Width           =   6615
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   720
      MaxLength       =   3
      TabIndex        =   0
      Top             =   90
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Classificação"
      Height          =   4665
      Left            =   8895
      TabIndex        =   40
      Top             =   540
      Width           =   4965
      Begin VB.Frame frm_classificacaoEspecifica_dev_rem_gra 
         Caption         =   "Obter tributos"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   600
         TabIndex        =   66
         Top             =   1800
         Width           =   4275
         Begin VB.OptionButton opt_classificacaoEspec_TRIB_ENTRADAS 
            Appearance      =   0  'Flat
            Caption         =   "De entrada do produto"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   150
            TabIndex        =   68
            Top             =   270
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton opt_classificacaoEspec_TRIB_SAIDAS 
            Appearance      =   0  'Flat
            Caption         =   "De saída do produto"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   67
            Top             =   255
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "* Informações de tributos gravados na tela de 'Cadastro de Produtos'"
            Enabled         =   0   'False
            Height          =   375
            Left            =   210
            TabIndex        =   69
            Top             =   540
            Width           =   3585
         End
      End
      Begin VB.CheckBox chkExigeAprovacaoOrcamento 
         Appearance      =   0  'Flat
         Caption         =   "Exige Aprovação"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   3450
         Width           =   1545
      End
      Begin VB.OptionButton O_Orçamento 
         Appearance      =   0  'Flat
         Caption         =   "&Orçamento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3180
         Width           =   1815
      End
      Begin VB.OptionButton O_Empréstimo 
         Appearance      =   0  'Flat
         Caption         =   "&Empréstimo Saída"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2850
         Width           =   1935
      End
      Begin VB.OptionButton O_Grátis_Saída 
         Appearance      =   0  'Flat
         Caption         =   "Devolução  /   Remessa  /   Grátis Saída"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1515
         Width           =   3255
      End
      Begin VB.OptionButton O_Ajuste_Saída 
         Appearance      =   0  'Flat
         Caption         =   "&Ajuste de Saída"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1128
         Width           =   1695
      End
      Begin VB.OptionButton O_Trans_Saída 
         Appearance      =   0  'Flat
         Caption         =   "&Transferência de Saída"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   744
         Width           =   2055
      End
      Begin VB.OptionButton O_Venda 
         Appearance      =   0  'Flat
         Caption         =   "&Venda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboGrupoFiscal 
      Bindings        =   "OpSaida.frx":50133
      DataSource      =   "datGrupoFiscal"
      Height          =   315
      Left            =   6060
      TabIndex        =   3
      Top             =   6930
      Visible         =   0   'False
      Width           =   585
      DataFieldList   =   "Código"
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2752
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Código"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   8255
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   1032
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Nome"
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4980
      TabIndex        =   54
      Top             =   3570
      Width           =   1215
   End
   Begin VB.Label lblSimplesNacional 
      AutoSize        =   -1  'True
      Caption         =   "Simples Nacional - CSO "
      Height          =   195
      Left            =   11190
      TabIndex        =   53
      Top             =   135
      Width           =   1680
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Grupo Fiscal"
      Height          =   195
      Index           =   9
      Left            =   5040
      TabIndex        =   48
      Top             =   6960
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblGrupoFiscal 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6960
      TabIndex        =   47
      Top             =   6930
      Visible         =   0   'False
      Width           =   795
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   4320
      Top             =   6900
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
      Bands           =   "OpSaida.frx":50150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CFOP"
      Height          =   195
      Left            =   9270
      TabIndex        =   41
      Top             =   135
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Left            =   60
      TabIndex        =   39
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmOpSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsOpSaidas    As Recordset
Private Num_Registro  As Variant
Private m_blnPSV      As Boolean

'04/07/2006 - mpdea
'Comentado referências ao case Agrofarm para que esteja acessível
'a todos os clientes, pois é uma configuração
'
'18/02/2005 - Daniel
'Solicitante: Agrofarm - RS
'Private m_blnAgrofarm As Boolean

Private Sub ShowRecord()
  Código.Text = rsOpSaidas("Código")
  Nome.Text = rsOpSaidas("Nome") & ""
  
  '31/01/2006 - mpdea
  'Grupo Fiscal
  cboGrupoFiscal.Text = rsOpSaidas.Fields("GrupoFiscal").Value & ""
  Call cboGrupoFiscal_LostFocus
  
  If rsOpSaidas("Tipo") = "V" Then O_Venda.Value = True
  If rsOpSaidas("Tipo") = "T" Then O_Trans_Saída.Value = True
  If rsOpSaidas("Tipo") = "A" Then O_Ajuste_Saída.Value = True
  If rsOpSaidas("Tipo") = "G" Then O_Grátis_Saída.Value = True
  If rsOpSaidas("Tipo") = "E" Then O_Empréstimo.Value = True
  If rsOpSaidas("Tipo") = "O" Then O_Orçamento.Value = True
  
  '19/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then
    chkValidade.Enabled = O_Ajuste_Saída.Value
    chkValidade.Value = -rsOpSaidas.Fields("Validade").Value
  End If
  '-------------------------------------------------------
  
  Estoque.Value = -rsOpSaidas("Estoque")
  Dinheiro.Value = -rsOpSaidas("Dinheiro")
  Comissão.Value = -rsOpSaidas("Comissão")
  ' Etiquetas.Value = -rsOpSaidas("Etiquetas")
  Nota.Value = -rsOpSaidas("Nota")
  
  '17/09/2009 - mpdea
  'Modelo de documento fiscal
  txtModeloDocumentoFiscal.Text = rsOpSaidas.Fields("ModeloDocumentoFiscal").Value & ""
  
  ICM.Value = -rsOpSaidas("ICM")
  IPI.Value = -rsOpSaidas("IPI")
  chkIpiTot.Value = -rsOpSaidas("IPI TOT")
  
  O_Calcula_ISS.Value = -rsOpSaidas("Calcula ISS")
  O_Base_IPI.Value = -rsOpSaidas("Base ICM com IPI")
  
  '11/11/2008 - mpdea
  chkSomaIcmsRetidoTotalNota.Value = IIf(rsOpSaidas.Fields("SomaIcmsRetidoTotalNota").Value, vbChecked, vbUnchecked)
  
  '09/03/2023 - Pablo
  If IPI.Value = 1 Then chkSomaIPITotalNF.Value = -rsOpSaidas("SomaIpiTotalNota")
  
  '14/03/2023 - Pablo
  If IPI.Value = 1 Then chkSomaIPITotalBC.Value = -rsOpSaidas("SomaIpiTotalBC")
  
  
  chkTelaObsTransp.Value = IIf(rsOpSaidas("InTelaObsTransp"), vbChecked, vbUnchecked)
  
  '13/05/2004 - Daniel
  'Adicionado o campo ComissaoServicos
  chkComissaoServicos.Value = IIf(rsOpSaidas("ComissaoServicos"), vbChecked, vbUnchecked)
  
  '05/11/2007 - Anderson
  'Adicionar o campo somar produtos no total da nota
  chkSomarProdutos.Value = IIf(rsOpSaidas("SomarProdutosTotalNota"), vbChecked, vbUnchecked)
  
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados - Otimização liberada
  '             para todos usuários do Quick Store
  '
  'Tratamento para o campo Emitir NF automaticamente
  chkEmitirNFManualmente.Value = IIf(rsOpSaidas.Fields("EmitirNFManualmente").Value, vbChecked, vbUnchecked)
  
  '04/07/2006 - mpdea
  'Comentado referências ao case Agrofarm para que esteja acessível
  'a todos os clientes, pois é uma configuração
  '
  '18/02/2005 - Daniel
  'Solicitante: Agrofarm - RS
  'Gerenciamento do campo Informante próprio (P)
  'If m_blnAgrofarm Then ...
  chkInformante.Value = IIf(rsOpSaidas("InformanteProprio").Value, vbChecked, vbUnchecked)
  
  Senha.Value = -rsOpSaidas("Senha")
  Código_Fiscal.Text = rsOpSaidas("Código Fiscal") & ""
  
  '30/03/2011 - Andrea
  txtCSOSN.Text = rsOpSaidas("CSO") & ""
 
 
  If Not IsNull(rsOpSaidas.Fields("ObterTributosProduto_EntradaOuSaida").Value) Then
      If rsOpSaidas.Fields("ObterTributosProduto_EntradaOuSaida").Value = 1 Then
          ' Obter TRIBUTOS de produtos (cadastro de produto) Entrada    1
          opt_classificacaoEspec_TRIB_ENTRADAS.Value = True
          opt_classificacaoEspec_TRIB_SAIDAS.Value = False
      Else
          ' Obter TRIBUTOS de produtos (cadastro de produto) Saida      2
          opt_classificacaoEspec_TRIB_ENTRADAS.Value = False
          opt_classificacaoEspec_TRIB_SAIDAS.Value = True
      End If
  Else
      opt_classificacaoEspec_TRIB_ENTRADAS.Value = False
      opt_classificacaoEspec_TRIB_SAIDAS.Value = False
  End If
 
  
  If Trim(rsOpSaidas("Ticket Imprimir") & "") = "" Then
    O_Perguntar.Value = True
    Combo_Tickets.Text = ""
  Else
    O_Específico.Value = True
    Combo_Tickets.Text = rsOpSaidas("Ticket Imprimir") & ""
  End If
  
  chkIcmFrete.Value = -rsOpSaidas("Calcula Icm Frete")
  chkSomaFrete.Value = -rsOpSaidas("Soma Frete")
  
  '12/04/2005 - Daniel
  'Adicionado campo Somar Seguro oriundo da
  'Loja Virtual (Web) ao total a receber
  If rsOpSaidas.Fields("SomarSeguro").Value = True Then
    chkSomaSeguro.Value = vbChecked
  Else
    chkSomaSeguro.Value = vbUnchecked
  End If
  '
  '28/06/2005 - Daniel
  'Solicitante: Osório (SEBO)
  'Adicionado o campo Altera status do pedido web para recebido
  If rsOpSaidas.Fields("AlteraStatusPedidoWeb").Value Then
    chkAlteraStatusPedidoWeb.Value = vbChecked
  Else
    chkAlteraStatusPedidoWeb.Value = vbUnchecked
  End If
  '------------------------------------------------------------
  
  txtIcmsFrete.Text = rsOpSaidas("Perc Icms Frete") & ""
  
  chkEntregas.Value = -(rsOpSaidas.Fields("ControleEntregas"))
'  chkEntregas.Enabled = Not (-Estoque.Value)
  cboOpEntrega.Text = rsOpSaidas.Fields("OpEntrega") & ""
  cboOpEntrega.Enabled = (chkEntregas.Value)
  chkExigeAprovacaoOrcamento.Value = -(rsOpSaidas.Fields("ExigeAprovacaoOrcamento").Value)
  
  '27/08/2004 - Daniel
  'Acerta empréstimo de entrada
  If rsOpSaidas.Fields("AcertaEmprestimoEntrada").Value Then
    chkAcertaEmprestimo.Value = vbChecked
  Else
    chkAcertaEmprestimo.Value = vbUnchecked
  End If
  
  '11/10/2002 - mpdea
  'Exibição da nome da operação
  cboOpEntrega_LostFocus
  
  '19/07/2007 - Anderson
  'Exibe campo SadigWeb
  Select Case "" & rsOpSaidas("SadigWeb_Tipo")
    Case "VE-VENDA", "BO-BONIFICAÇÃO", "OU-OUTROS"
      cboSadigWebTipo.Text = "" & rsOpSaidas("SadigWeb_Tipo").Value
    Case Else
      cboSadigWebTipo.Text = "(Nenhum)"
  End Select
  
  '29/04/2008 - mpdea
  'Exibição de tela para preenchimento do número de documento (CPF ou CNPJ)
  chkExibirTelaNumeroDocumento.Value = IIf(rsOpSaidas.Fields("ExibirTelaNumeroDocumento").Value, vbChecked, vbUnchecked)
  
  ChkPermiteDadosCliente.Value = IIf(rsOpSaidas.Fields("PermiteMostrarCliente").Value, vbChecked, vbUnchecked)
  
  Num_Registro = rsOpSaidas.Bookmark
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rsOpSaidas
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsOpSaidas
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsOpSaidas
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsOpSaidas
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
  End Select
  
  '25/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then VerificaClassificacao
  
End Sub

Private Sub DeleteRecord()
  Dim Aux_Filial As Integer
  Dim Aux_Sequência As Integer
  Dim rsSaidas As Recordset

  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Encontre uma operação antes."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  'Operação bloqueada
  If rsOpSaidas.Fields("Locked").Value Then
    DisplayMsg "Não é possível excluir. Operação bloqueada pelo sistema."
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde, verificando se esta operacão não está em uso.")
  
  Set rsSaidas = db.OpenRecordset("SELECT * FROM Saídas WHERE Operação = " & rsOpSaidas("Código").Value, dbOpenDynaset)
  
  If Not rsSaidas.EOF Then
    DisplayMsg "Este Tipo de Operação não pode ser apagado por estar em uso."
    rsSaidas.Close
    Set rsSaidas = Nothing
    Exit Sub
  End If
  
  rsSaidas.Close
  Set rsSaidas = Nothing
  
  gsTitle = LoadResString(201)
  gsMsg = "Você deseja realmente apagar este Tipo de Operação?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    rsOpSaidas.Delete
    DisplayMsg "Tipo de Operação apagado com sucesso."
    Call ClearScreen
  End If
  
End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  On Error GoTo Processa_Erro
  
  'Operação bloqueada
  If Not IsNull(Num_Registro) Then
    If rsOpSaidas.Fields("Locked").Value Then
      DisplayMsg "Operação bloqueada pelo sistema. Necessário senha do Gerente."
      If Not frmGerente.gbSenhaGerente Then
        Exit Sub
      End If
    End If
  End If
  
  Rem Verifica Conta
  Erro = False
  If IsNull(Código.Text) Then Erro = True
  If Erro = False Then If Código.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Código.Text) Then Erro = True
  If Erro = False Then If Val(Código.Text) < 500 Or Val(Código.Text) > 999 Then Erro = True
  
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Escolha códigos entre 500 e 999"
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Código.SetFocus
    Exit Sub
  End If
  
  
  Erro = False
  If IsNull(Nome.Text) Then Erro = True
  If Erro = False Then If Nome.Text = "" Then Erro = True
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Por favor digite o nome da operação."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Nome.SetFocus
    Exit Sub
  End If
  
  
  '11/10/2002 - mpdea
  'Acrescentado verificação para o controle de entregas
  If chkEntregas.Value = vbChecked Then
    cboOpEntrega_LostFocus
    If lblOpEntrega.Caption = "" Then
      DisplayMsg "Escolha a operação a ser associada com a operação de entrega."
      cboOpEntrega.SetFocus
      Exit Sub
    End If
  End If
  
  
  If O_Específico.Value = True Then
    If Combo_Tickets.Text = "" Then
      gsTitle = LoadResString(201)
      gsMsg = "Escolha o modelo do ticket."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Combo_Tickets.Enabled = True
    End If
  End If
  
  '18/07/2007 - Anderson
  'Validação do tipo de cliente SadigWeb
  If cboSadigWebTipo.Text <> "(Nenhum)" And _
     cboSadigWebTipo.Text <> "VE-VENDA" And _
     cboSadigWebTipo.Text <> "BO-BONIFICAÇÃO" And _
     cboSadigWebTipo.Text <> "OU-OUTROS" Then
     
    MsgBox "Tipo de Saída utilizado para exportar dados no sistema Sadig Web, está inválido!", vbCritical, "Quick Store"
    cboSadigWebTipo.SetFocus
    Exit Sub
          
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
  'Inicia transação
  ws.BeginTrans
  
  With rsOpSaidas
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código") = Val(Código.Text)
    Else
      .Edit
    End If
    
    .Fields("Nome") = Nome.Text
    
    '31/01/2006 - mpdea
    'Grupo Fiscal
    Call cboGrupoFiscal_LostFocus
    If lblGrupoFiscal.Caption = "" Then
      .Fields("GrupoFiscal").Value = Null
    Else
      .Fields("GrupoFiscal").Value = cboGrupoFiscal.Text
    End If
    
    If O_Venda.Value = True Then .Fields("Tipo") = "V"
    If O_Trans_Saída.Value = True Then .Fields("Tipo") = "T"
    If O_Ajuste_Saída.Value = True Then .Fields("Tipo") = "A"
    If O_Grátis_Saída.Value = True Then .Fields("Tipo") = "G"
    If O_Empréstimo.Value = True Then .Fields("Tipo") = "E"
    If O_Orçamento.Value = True Then .Fields("Tipo") = "O"
    
    If Estoque.Value = 1 Then .Fields("Estoque") = True
    If Estoque.Value = 0 Then .Fields("Estoque") = False
    
    If Dinheiro.Value = 1 Then .Fields("Dinheiro") = True
    If Dinheiro.Value = 0 Then .Fields("Dinheiro") = False
    
    If Comissão.Value = 1 Then .Fields("Comissão") = True
    If Comissão.Value = 0 Then .Fields("Comissão") = False
    
    ' If Etiquetas.Value = 1 Then .Fields("Etiquetas") = True
    ' If Etiquetas.Value = 0 Then .Fields("Etiquetas") = False
    
    If Nota.Value = 1 Then .Fields("Nota") = True
    If Nota.Value = 0 Then .Fields("Nota") = False
    
    '17/09/2009 - mpdea
    'Modelo de documento fiscal
    rsOpSaidas.Fields("ModeloDocumentoFiscal").Value = txtModeloDocumentoFiscal.Text
    
    If ICM.Value = 1 Then .Fields("ICM") = True
    If ICM.Value = 0 Then .Fields("ICM") = False
    
    If IPI.Value = 1 Then .Fields("IPI") = True
    If IPI.Value = 0 Then .Fields("IPI") = False
    
    '09/03/2023 - Pablo
    If chkSomaIPITotalNF.Value = 1 Then .Fields("SomaIpiTotalNota") = True
    If chkSomaIPITotalNF.Value = 0 Then .Fields("SomaIpiTotalNota") = False
    
    '14/03/2023 - Pablo
    If chkSomaIPITotalBC.Value = 1 Then .Fields("SomaIpiTotalBC") = True
    If chkSomaIPITotalBC.Value = 0 Then .Fields("SomaIpiTotalBC") = False
    
    If chkIpiTot.Value = 1 Then .Fields("IPI TOT") = True
    If chkIpiTot.Value = 0 Then .Fields("IPI TOT") = False
    
    '19/02/2004 - Daniel
    'Case: PSV
    .Fields("Validade") = (chkValidade.Enabled And chkValidade.Value = 1 And chkValidade.Visible)
    '--------------------------------------------------------------------------------------------
    
    If O_Calcula_ISS.Value = 1 Then .Fields("Calcula ISS") = True
    If O_Calcula_ISS.Value = 0 Then .Fields("Calcula ISS") = False
    
    If O_Base_IPI.Value = 1 Then .Fields("Base ICM com IPI") = True
    If O_Base_IPI.Value = 0 Then .Fields("Base ICM com IPI") = False
    
    '11/11/2008 - mpdea
    .Fields("SomaIcmsRetidoTotalNota").Value = chkSomaIcmsRetidoTotalNota.Value = vbChecked
    
    If Senha.Value = 1 Then .Fields("Senha") = True
    If Senha.Value = 0 Then .Fields("Senha") = False
    
    If opt_classificacaoEspec_TRIB_ENTRADAS.Value = True Then
        ' Obter TRIBUTOS de produtos (cadastro de produto) Entrada    1
        .Fields("ObterTributosProduto_EntradaOuSaida").Value = 1
    Else
        ' Obter TRIBUTOS de produtos (cadastro de produto) Saida      2
        .Fields("ObterTributosProduto_EntradaOuSaida").Value = 2
    End If
    
    .Fields("InTelaObsTransp") = (chkTelaObsTransp.Value = 1)
    
    '13/05/2004 - Daniel
    'Adicionado o campo ComissaoServicos
    If chkComissaoServicos.Value = vbChecked Then
      .Fields("ComissaoServicos").Value = True
    Else
      .Fields("ComissaoServicos").Value = False
    End If
    
    '05/11/2007 - Anderson
    'Adicionar o campo somar produtos no total da nota
    If chkSomarProdutos.Value = vbChecked Then
      .Fields("SomarProdutosTotalNota").Value = True
    Else
      .Fields("SomarProdutosTotalNota").Value = False
    End If
    
    '19/05/2005 - Daniel
    '
    'Solicitante: Pedágio Calçados - Otimização liberada
    '             para todos usuários do Quick Store
    '
    'Tratamento para o campo Emitir NF automaticamente
    If chkEmitirNFManualmente.Value = vbChecked Then
      .Fields("EmitirNFManualmente").Value = True
    Else
      .Fields("EmitirNFManualmente").Value = False
    End If
        
    .Fields("Código Fiscal") = Código_Fiscal.Text
    
    '30/03/2011 - Andrea
    .Fields("CSO") = txtCSOSN.Text
    
    .Fields("Ticket Imprimir") = ""
    If O_Específico.Value = True Then
      .Fields("Ticket Imprimir") = Combo_Tickets.Text
    End If
    .Fields("Calcula Icm Frete") = chkIcmFrete.Value
    .Fields("Soma Frete") = chkSomaFrete.Value
    
    '12/04/2005 - Daniel
    'Adicionado campo Somar Seguro oriundo da
    'Loja Virtual (Web) ao total a receber
    If chkSomaSeguro.Value = vbChecked Then
      .Fields("SomarSeguro").Value = True
    Else
      .Fields("SomarSeguro").Value = False
    End If
    '
    '28/06/2005 - Daniel
    'Solicitante: Osório (SEBO)
    'Adicionado o campo Altera status do pedido web para recebido
    If chkAlteraStatusPedidoWeb.Value = vbChecked Then
      .Fields("AlteraStatusPedidoWeb").Value = True
    Else
      .Fields("AlteraStatusPedidoWeb").Value = False
    End If
    '------------------------------------------------------------
    
    If IsNull(txtIcmsFrete.Text) Or txtIcmsFrete.Text = "" Then txtIcmsFrete.Text = 0
    .Fields("Perc Icms Frete") = txtIcmsFrete.Text
    
    .Fields("ControleEntregas") = -chkEntregas.Value
    
    
    '11/10/2002 - mpdea
    'Alterado para que salve a opção de entrega desmarcada
'    If IsNumeric(cboOpEntrega.Text) Then
      .Fields("OpEntrega") = CInt("0" & cboOpEntrega.Text)
'    End If
    .Fields("ExigeAprovacaoOrcamento").Value = -chkExigeAprovacaoOrcamento.Value
    
    '27/08/2004 - Daniel
    'Acerta empréstimo de entrada
    If chkAcertaEmprestimo.Value = vbChecked Then
      .Fields("AcertaEmprestimoEntrada").Value = True
    Else
      .Fields("AcertaEmprestimoEntrada").Value = False
    End If
    
    '04/07/2006 - mpdea
    'Comentado referências ao case Agrofarm para que esteja acessível
    'a todos os clientes, pois é uma configuração
    '
    '18/02/2005 - Daniel
    'Solicitante: Agrofarm - RS
    'Gerenciamento do campo Informante próprio (P)
    'If m_blnAgrofarm Then
      If chkInformante.Value = vbChecked Then
        .Fields("InformanteProprio").Value = True
      Else
        .Fields("InformanteProprio").Value = False
      End If
    'End If
    
    '19/07/2007 - Anderson
    'Grava informação do tipo de saída SadigWeb
    .Fields("SadigWeb_Tipo").Value = cboSadigWebTipo.Text
    
    '29/04/2008 - mpdea
    'Exibição de tela para preenchimento do número de documento (CPF ou CNPJ)
    rsOpSaidas.Fields("ExibirTelaNumeroDocumento").Value = IIf(chkExibirTelaNumeroDocumento.Value = vbChecked, True, False)
    
    If ChkPermiteDadosCliente.Value = vbChecked Then
      .Fields("PermiteMostrarCliente").Value = True
    Else
      .Fields("PermiteMostrarCliente").Value = False
    End If
    
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
    
  End With
  
  'Efetua registro do Log
  If rsOpSaidas.Fields("Locked").Value Then
    db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
      Now & "#, 'Operação de Saída bloqueada foi alterada pelo usuário " & _
      gnUserCode & "-" & gsUserName & " através de senha do gerente', 'CADASTRO')", dbFailOnError
  End If
  
  'Finaliza transação
  ws.CommitTrans
  
  Call StatusMsg("")
  
  Exit Sub
  
Processa_Erro:
  ws.Rollback
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar registro."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & " - " & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub


Public Sub ClearScreen()
  Call StatusMsg("")
  Código.Text = ""
  Nome.Text = ""
  
  '31/01/2006 - mpdea
  'Grupo Fiscal
  cboGrupoFiscal.Text = ""
  lblGrupoFiscal.Caption = ""
  
  O_Venda.Value = True
  
  Estoque.Value = 0
  Dinheiro.Value = 0
  Comissão.Value = 0
  ' Etiquetas.Value = 0
  Nota.Value = 0
  
  '17/09/2009 - mpdea
  'Modelo de documento fiscal
  txtModeloDocumentoFiscal.Text = ""
  
  ICM.Value = 0
  IPI.Value = 0
  O_Calcula_ISS.Value = 0
  O_Base_IPI.Value = 0
  Senha.Value = 0
  Código_Fiscal.Text = ""
  
  '30/03/2011 - Andrea
  txtCSOSN.Text = ""
  
  '11/11/2008 - mpdea
  chkSomaIcmsRetidoTotalNota.Value = vbUnchecked
  
  O_Perguntar.Value = True
  Combo_Tickets.Text = ""
  
  chkIpiTot.Value = 0
  txtIcmsFrete.Text = 0
  
  chkIcmFrete.Value = False
  chkSomaFrete.Value = False
  
  '12/04/2005 - Daniel
  'Adicionado campo Somar Seguro oriundo da Loja Virtual (Web)
  'ao total a receber
  chkSomaSeguro.Value = False
  '
  '28/06/2005 - Daniel
  'Solicitante: Osório (SEBO)
  'Adicionado o campo Altera status do pedido web para recebido
  chkAlteraStatusPedidoWeb.Value = False
  '------------------------------------------------------------
  
  '11/08/2003 - mpdea
  'Campo não estava sendo desmarcado
  chkTelaObsTransp.Value = vbUnchecked
  
  '13/05/2004 - Daniel
  'Adicionado novo campo ComissaoServicos
  chkComissaoServicos.Value = vbUnchecked
  
  '05/11/2007 - Anderson
  'Adicionar o campo somar produtos no total da nota
  chkSomarProdutos.Value = vbUnchecked
  
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados - Otimização liberada
  '             para todos usuários do Quick Store
  '
  'Tratamento para o campo Emitir NF automaticamente
  chkEmitirNFManualmente.Value = vbUnchecked
  
  '19/02/2004 - Daniel
  If m_blnPSV Then
    chkValidade.Value = vbUnchecked
    chkValidade.Enabled = False
  End If
  '------------------------------
  
  '04/07/2006 - mpdea
  'Comentado referências ao case Agrofarm para que esteja acessível
  'a todos os clientes, pois é uma configuração
  '
  '18/02/2005 - Daniel
  'Solicitante: Agrofarm - RS
  'Gerenciamento do campo Informante próprio (P)
  'If m_blnAgrofarm Then ...
  chkInformante.Value = vbUnchecked
  '------------------------------------------------------
  
  chkEntregas.Value = 0
  cboOpEntrega.Text = ""
  lblOpEntrega.Caption = ""
  chkEntregas.Enabled = True
  cboOpEntrega.Enabled = chkEntregas.Value
  chkExigeAprovacaoOrcamento.Value = vbUnchecked
  
  '27/08/2004 - Daniel
  chkAcertaEmprestimo.Value = vbUnchecked
  ChkPermiteDadosCliente.Value = vbUnchecked
  
  If Not rsOpSaidas.EOF Then
    On Error Resume Next
    rsOpSaidas.MoveFirst
    rsOpSaidas.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  
  opt_classificacaoEspec_TRIB_ENTRADAS.Value = True
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = False
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = False
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = False
  
  Código.SetFocus
  
  Call IPI_Click
  'Call chkSomaIPITotalNF_Click

End Sub


'31/01/2006 - mpdea
'Exibe o código do Grupo Fiscal selecionado
Private Sub cboGrupoFiscal_CloseUp()
  cboGrupoFiscal.Text = cboGrupoFiscal.Columns("Código").Text
  Call cboGrupoFiscal_LostFocus
End Sub

'31/01/2006 - mpdea
'Valida o código e exibe o nome do Grupo Fiscal selecionado
Private Sub cboGrupoFiscal_LostFocus()
  Dim intItem As Integer
  
  
  On Error GoTo ErrHandler
  
  
  lblGrupoFiscal.Caption = ""
  
  If cboGrupoFiscal.Text <> "" Then
    If Not IsDataType(dtInteger, cboGrupoFiscal.Text, intItem) Then
      DisplayMsg "Código de Grupo Fiscal inválido."
      cboGrupoFiscal.Text = ""
      cboGrupoFiscal.SetFocus
      Exit Sub
    End If
        
    If intItem < 1 Or intItem > 9999 Then
      DisplayMsg "Código de Grupo Fiscal inválido."
      cboGrupoFiscal.Text = ""
      cboGrupoFiscal.SetFocus
      Exit Sub
    End If
    
    With datGrupoFiscal.Recordset
      .FindFirst "Código = " & intItem
      If Not .NoMatch Then
        lblGrupoFiscal.Caption = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

'-----------------------------------------------------------------------------------
'11/10/2002 - mpdea
'Eventos incluídos
Private Sub cboOpEntrega_Click()
  cboOpEntrega.Text = cboOpEntrega.Columns(1).Text
End Sub

Private Sub cboOpEntrega_CloseUp()
  cboOpEntrega.Text = cboOpEntrega.Columns(1).Text
  cboOpEntrega_LostFocus
End Sub

Private Sub cboOpEntrega_LostFocus()
 Dim intRet As Integer
 
 Call IsDataType(dtInteger, cboOpEntrega.Text, intRet)
 lblOpEntrega.Caption = gstrGetNameOper(tmSaidas, intRet)
End Sub
'-----------------------------------------------------------------------------------

Private Sub chkEntregas_Click()
  cboOpEntrega.Enabled = chkEntregas.Value = vbChecked
  lblOpEntrega.Enabled = chkEntregas.Value = vbChecked
End Sub

Private Sub chkIcmFrete_Click()
If chkIcmFrete.Value = False Then
   txtIcmsFrete.Enabled = False
Else
   txtIcmsFrete.Enabled = True
End If
End Sub

Private Sub cmd_pesquisarOperSaida_Click()
  MsgBox "Navegue pelas setas acima para visualizar as operações cadastradas", vbInformation, "Atenção"
'  Dim objFrmPesquisaOperSaida As frmOpEntSaiPesquisa
'  Set objFrmPesquisaOperSaida = New frmOpEntSaiPesquisa
'  objFrmPesquisaOperSaida.iOrigemOperacao = 1  '
'  objFrmPesquisaOperSaida.Show
End Sub

Private Sub Código_LostFocus()
  
  If Not IsNumeric(Código.Text) Then Exit Sub
  If Val(Código) <= 0 Then Exit Sub
  
  With rsOpSaidas
    .FindFirst "Código = " & CInt(Código.Text)
    If Not .NoMatch Then
      Call ShowRecord
    Else
      Beep
    End If
  End With
  
End Sub

Private Sub Estoque_Click()
'  chkEntregas.Enabled = Not CBool(-Estoque.Value)
'  cboOpEntrega.Enabled = Not CBool(-Estoque.Value)
  
  '11/10/2002 - mpdea
  'Desmarca a opção de entregas e limpa o código da operação
  'se a opção diminuir estoque estiver marcada
  If Estoque.Value = vbChecked Then
    chkEntregas.Value = vbUnchecked
    cboOpEntrega.Text = ""
    lblOpEntrega.Caption = ""
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10003)
      Set objHelp = Nothing
  Else
      Call HandleKeyDown(KeyCode, Shift)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

'06/01/2006 - mpdea
'Incluído tratamento de erros
Private Sub Form_Load()
  Dim Texto As String
  Dim Aux_Str As String
  Dim Fim As Integer
 
  
  On Error GoTo ErrHandler
 
 
  Call CenterForm(Me)
  Num_Registro = Null
    
  '06/02/2006 - mpdea
  'Seta Banco de dados para controle Data
  datGrupoFiscal.DatabaseName = gsQuickDBFileName
  
  '11/10/2002 - mpdea
  'Operações de saída
  Data2.DatabaseName = gsQuickDBFileName
  
  Set rsOpSaidas = db.OpenRecordset("Operações Saída", dbOpenDynaset)
  
  Rem Enche Combo_Tickets
  Texto = Dir(gsConfigPath & "*.CTI")
  Combo_Tickets.Clear
  
  Do While Len(Texto) > 0
    Combo_Tickets.AddItem Texto
    Texto = Dir
  Loop
 
  Call ActiveBarLoadToolTips(Me)
  
  '19/02/2004 - Daniel
  'Case.......: PSV Informática
  'Finalidade.: Compôr o field Validade em Operações Saída
  If CheckSerialCaseMod("QS35552-811", "QS37705-639", "QS37825-830", "QS38933-772", "QS39369-521") Then
  
     m_blnPSV = True
     
     chkValidade.Visible = True
     chkValidade.Enabled = False
     
  End If
  '-----------------------------------------
  
  '27/08/2004 - Daniel
  'Case: Resultado
  'Acerta Empréstimo de Entrada
  If CheckSerialCaseMod("QS40590-987") Then
    fraAcerta.Enabled = True
    chkAcertaEmprestimo.Enabled = True
  Else
    fraAcerta.Enabled = False
    chkAcertaEmprestimo.Enabled = False
  End If
  '-----------------------------------------
  
  '04/07/2006 - mpdea
  'Comentado referências ao case Agrofarm para que esteja acessível
  'a todos os clientes, pois é uma configuração
  '
  '18/02/2005 - Daniel
  'Solicitante: Agrofarm - RS
  'Gerenciamento do campo Informante próprio (P). Para toda operação
  'que possuir este campo habilitado, no momento da geração do
  'arquivo para o Sintegra no registro 50 o campo emitente será
  'igual a P.
  'No caso da Agrofarm as vezes eles emitem notas contra eles mesmos
  'no momento da entrada quando alguma venda ao consumidor retorna
  'como devolução
'  If CheckSerialCaseMod("QS35815-716", "QS37243-804") Then
'    m_blnAgrofarm = True
'    fraReceita.Enabled = True
'    chkInformante.Enabled = True
'  Else
'    fraReceita.Enabled = False
'    chkInformante.Enabled = False
'  End If
  '-----------------------------------------------------------------
  
  '25/04/2005 - Daniel
  'Verificação se a empresa utiliza o Quick Web em caso negativo
  'desabilitaremos os checks "chkSomaSeguro" e "chkAlteraStatusPedidoWeb"
  If VerificaUsoWeb Then
    chkSomaSeguro.Enabled = True
    chkAlteraStatusPedidoWeb.Enabled = True
  Else
    chkSomaSeguro.Enabled = False
    chkAlteraStatusPedidoWeb.Enabled = False
  End If
  '-----------------------------------------------------------------
  
  '19/07/2007 - Anderson
  fraSadigWeb.Visible = g_blnSadigWeb
  cboSadigWebTipo.Text = "(Nenhum)"
  
  Me.Show
  DoEvents
   
  Call ClearScreen
  
  Exit Sub
    
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsOpSaidas.Close
  Set rsOpSaidas = Nothing
End Sub



Private Sub IPI_Click()
  chkSomaIPITotalNF.Enabled = IPI.Value
  If Not IPI.Value Then chkSomaIPITotalNF.Value = 0
  chkSomaIPITotalBC.Enabled = IPI.Value
  If Not IPI.Value Then chkSomaIPITotalBC.Value = 0
End Sub

Private Sub IPI_KeyPress(KeyAscii As Integer)
  Call IPI_Click
End Sub

Private Sub O_Ajuste_Saída_Click()
  chkExigeAprovacaoOrcamento.Enabled = False
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = False
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = False
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = False
  
  '25/02/2004 - Daniel
  'Case: PSV
  VerificaClassificacao

End Sub

Private Sub O_Empréstimo_Click()
  chkExigeAprovacaoOrcamento.Enabled = False
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = False
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = False
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = False

End Sub

Private Sub O_Específico_Click()
 If O_Específico.Value = True Then
   Combo_Tickets.Enabled = True
 End If
End Sub

Private Sub O_Grátis_Saída_Click()
  chkExigeAprovacaoOrcamento.Enabled = False
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = True
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = True
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = True
End Sub

Private Sub O_Orçamento_Click()
  chkExigeAprovacaoOrcamento.Enabled = True
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = False
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = False
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = False
End Sub

Private Sub O_Perguntar_Click()
 If O_Específico.Value = False Then
   Combo_Tickets.Enabled = False
 End If
End Sub

Private Sub O_Trans_Saída_Click()
  chkExigeAprovacaoOrcamento.Enabled = False
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = False
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = False
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = False
End Sub

Private Sub O_Venda_Click()
  chkExigeAprovacaoOrcamento.Enabled = False
  frm_classificacaoEspecifica_dev_rem_gra.Enabled = False
  opt_classificacaoEspec_TRIB_ENTRADAS.Enabled = False
  opt_classificacaoEspec_TRIB_SAIDAS.Enabled = False
End Sub

Private Sub txtIcmsFrete_KeyPress(KeyAscii As Integer)
KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub VerificaClassificacao()
  '25/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then 'Caso a rotina seja chamada a partir do O_Ajuste_Saída_Click() precisamos verificar se é PSV
    chkValidade.Visible = O_Ajuste_Saída.Value
    chkValidade.Enabled = O_Ajuste_Saída.Value
  End If
  
End Sub

Private Function VerificaUsoWeb() As Boolean
  '25/04/2005 - Daniel
  'Verificação se a empresa utiliza o Quick Web
  Dim rstParametro As Recordset
  
  Set rstParametro = db.OpenRecordset("SELECT WorkWeb FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  
  With rstParametro
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      VerificaUsoWeb = .Fields("WorkWeb").Value
    End If
    .Close
  End With
  
  Set rstParametro = Nothing
  
End Function
