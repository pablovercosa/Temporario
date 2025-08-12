VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNFe 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Acelerador de Nota Fiscal Eletrônica"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15480
   FillColor       =   &H00E5E5E5&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E5E5E5&
   Icon            =   "frmNFe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   15480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_atalho 
      BackColor       =   &H000000FF&
      Caption         =   "ATALHO XML"
      Height          =   435
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   8010
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picture_statusProcessamento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   9600
      Picture         =   "frmNFe.frx":4E95A
      ScaleHeight     =   615
      ScaleWidth      =   855
      TabIndex        =   60
      Top             =   8950
      Width           =   855
   End
   Begin VB.CommandButton cmd_calendarioDtFim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7470
      Picture         =   "frmNFe.frx":4F1C1
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   45
      Width           =   465
   End
   Begin VB.CommandButton cmd_calendarioDtIni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5265
      Picture         =   "frmNFe.frx":4FAA3
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   45
      Width           =   465
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   7335
      Left            =   45
      TabIndex        =   3
      Top             =   585
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "NFe"
      TabPicture(0)   =   "frmNFe.frx":50385
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMotivoCancelamento(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "msk_dataRetroativa"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grdMovimento1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdMarcarDesmarcar(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdMarcarDesmarcar(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_visualizarXML_AbaNFe"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_imprimirDanfe"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "grdMovimento"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdPesquisar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_enviarNFe"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd_posicaoConsulta"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMotivoCancelamento"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd_cancelarNFe"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chk_naoNFCe"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmd_ajusteStatusNoQuick"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chk_ordemNumNFe"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbo_visualizarNFeGradeVerde"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chk_enviaComDataRetroativa"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmd_dataRetroativa"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cbo_operacao"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "NFe Erros e Críticas"
      TabPicture(1)   =   "frmNFe.frx":503A1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_posicaoNFeErro"
      Tab(1).Control(1)=   "cmd_helpRejeicao"
      Tab(1).Control(2)=   "cmdPesquisarRetornos"
      Tab(1).Control(3)=   "cmd_visualizaXML"
      Tab(1).Control(4)=   "grdRetorno"
      Tab(1).Control(5)=   "gridRetorno"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "NFe Carta Correção"
      TabPicture(2)   =   "frmNFe.frx":503BD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblMotivoInutilizacao(1)"
      Tab(2).Control(1)=   "lblSerie(0)"
      Tab(2).Control(2)=   "lblNumeroNFeInicial(1)"
      Tab(2).Control(3)=   "lbl_tipoDoc(2)"
      Tab(2).Control(4)=   "grdCC"
      Tab(2).Control(5)=   "cmd_enviarCC"
      Tab(2).Control(6)=   "cmd_pesquisarCC"
      Tab(2).Control(7)=   "txt_descCC"
      Tab(2).Control(8)=   "txt_serieCC"
      Tab(2).Control(9)=   "txt_numNFeCC"
      Tab(2).Control(10)=   "cmd_imprimirDanfeCC"
      Tab(2).Control(11)=   "txt_tipoDocCC"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Inutilizadas"
      TabPicture(3)   =   "frmNFe.frx":503D9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblAno(0)"
      Tab(3).Control(1)=   "lblNumeroNFeFinal(1)"
      Tab(3).Control(2)=   "lblNumeroNFeInicial(0)"
      Tab(3).Control(3)=   "lblSerie(1)"
      Tab(3).Control(4)=   "lblMotivoInutilizacao(0)"
      Tab(3).Control(5)=   "grdInutilizadas"
      Tab(3).Control(6)=   "txtNumeroNFeInicial"
      Tab(3).Control(7)=   "txtNumeroNFeFinal"
      Tab(3).Control(8)=   "txt_anoInutilizacao"
      Tab(3).Control(9)=   "txtSerie"
      Tab(3).Control(10)=   "txtMotivoInutilizacao"
      Tab(3).Control(11)=   "cmd_pesqInutilizadas"
      Tab(3).Control(12)=   "cmd_inutilizarNFe"
      Tab(3).Control(13)=   "opt_inutilizadasNFe"
      Tab(3).Control(14)=   "opt_inutilizadasNFCe"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Contingência NFCe/NFe"
      TabPicture(4)   =   "frmNFe.frx":503F5
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ssTab_contingencia"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "NFe Parametrização"
      TabPicture(5)   =   "frmNFe.frx":50411
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).Control(1)=   "Frame2"
      Tab(5).Control(2)=   "Frame3"
      Tab(5).Control(3)=   "Frame6"
      Tab(5).Control(4)=   "Frame4"
      Tab(5).Control(5)=   "Frame5"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "NFCe"
      TabPicture(6)   =   "frmNFe.frx":5042D
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblTitle(1)"
      Tab(6).Control(1)=   "grd_nfceNormal"
      Tab(6).Control(2)=   "msk_dataDiaNFCe"
      Tab(6).Control(3)=   "cmd_visualizarXML_nfceNormal"
      Tab(6).Control(4)=   "cmd_pesquisarNFCeNormal"
      Tab(6).Control(5)=   "cmd_calendarioNFCeAbaCinza"
      Tab(6).Control(6)=   "cmd_NFCe_xml_manutencao"
      Tab(6).ControlCount=   7
      Begin VB.TextBox txt_posicaoNFeErro 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   1845
         Left            =   -74850
         MultiLine       =   -1  'True
         TabIndex        =   112
         Top             =   4800
         Width           =   14955
      End
      Begin VB.CommandButton cmd_helpRejeicao 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Detalhamento do erro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -67200
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   6675
         Width           =   7335
      End
      Begin VB.ComboBox cbo_operacao 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   7710
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   900
         Width           =   5685
      End
      Begin VB.OptionButton opt_inutilizadasNFCe 
         Appearance      =   0  'Flat
         Caption         =   "NFC-e/Cupom Fiscal"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67740
         TabIndex        =   107
         Top             =   450
         Width           =   2175
      End
      Begin VB.OptionButton opt_inutilizadasNFe 
         Appearance      =   0  'Flat
         Caption         =   "NF-e"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -69630
         TabIndex        =   106
         Top             =   450
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.CommandButton cmd_dataRetroativa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         Picture         =   "frmNFe.frx":50449
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   5520
         Width           =   465
      End
      Begin VB.CheckBox chk_enviaComDataRetroativa 
         Caption         =   "Enviar com data retroativa"
         Height          =   255
         Left            =   210
         TabIndex        =   103
         Top             =   5610
         Width           =   2355
      End
      Begin VB.ComboBox cbo_visualizarNFeGradeVerde 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmNFe.frx":50D2B
         Left            =   1020
         List            =   "frmNFe.frx":50D41
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   900
         Width           =   3225
      End
      Begin VB.CheckBox chk_ordemNumNFe 
         Appearance      =   0  'Flat
         Caption         =   "Ordem por Número NFe"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4350
         TabIndex        =   100
         Top             =   945
         Width           =   2205
      End
      Begin VB.CommandButton cmd_NFCe_xml_manutencao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Visualizar XML (Manutenção)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -61200
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   6750
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmd_ajusteStatusNoQuick 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Ajustar status p/ autorizada (apenas no quick)"
         Height          =   465
         Left            =   10260
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   6750
         Width           =   4890
      End
      Begin VB.CommandButton cmd_calendarioNFCeAbaCinza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72660
         Picture         =   "frmNFe.frx":50D87
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   390
         Width           =   465
      End
      Begin VB.CommandButton cmd_pesquisarNFCeNormal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   840
         Width           =   14805
      End
      Begin VB.CommandButton cmd_visualizarXML_nfceNormal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Visualizar XML de Retorno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   5820
         Width           =   14805
      End
      Begin TabDlg.SSTab ssTab_contingencia 
         Height          =   6525
         Left            =   -74820
         TabIndex        =   75
         Top             =   510
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   11509
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "NFCe"
         TabPicture(0)   =   "frmNFe.frx":51669
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblTitle(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "grid_nfce_cont"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "msk_dataDia"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmd_calendarioNFCe"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmd_posicaoNFCe_cont"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmd_pesquisarNFCe"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmd_visualizaNFCe_cont"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "NFe"
         TabPicture(1)   =   "frmNFe.frx":51685
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblHoraEntradaContingencia(0)"
         Tab(1).Control(1)=   "lblMotivoContingencia(1)"
         Tab(1).Control(2)=   "lblDataEntradaContingencia(9)"
         Tab(1).Control(3)=   "dteEntradaContingencia"
         Tab(1).Control(4)=   "txtHoraEntradaContingencia"
         Tab(1).Control(5)=   "txtMotivoContingencia"
         Tab(1).ControlCount=   6
         Begin VB.CommandButton cmd_visualizaNFCe_cont 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "Visualizar XML (Manutenção)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   470
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   5880
            Width           =   14805
         End
         Begin VB.CommandButton cmd_pesquisarNFCe 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "Pesquisar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   470
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   900
            Width           =   14805
         End
         Begin VB.CommandButton cmd_posicaoNFCe_cont 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "Posição Consulta NFCe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   470
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   5310
            Width           =   14805
         End
         Begin VB.CommandButton cmd_calendarioNFCe 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2280
            Picture         =   "frmNFe.frx":516A1
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   450
            Width           =   465
         End
         Begin VB.TextBox txtMotivoContingencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74610
            TabIndex        =   77
            Top             =   1110
            Width           =   10020
         End
         Begin VB.TextBox txtHoraEntradaContingencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74610
            TabIndex        =   76
            Top             =   2895
            Width           =   1710
         End
         Begin MSMask.MaskEdBox dteEntradaContingencia 
            Height          =   375
            Left            =   -74595
            TabIndex        =   78
            ToolTipText     =   "Pressione F2 para Calendário"
            Top             =   1995
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   14737632
            MaxLength       =   22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
         Begin MSMask.MaskEdBox msk_dataDia 
            Height          =   315
            Left            =   930
            TabIndex        =   83
            ToolTipText     =   "Pressione F2 para Calendário"
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12648447
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
         Begin SSDataWidgets_B.SSDBGrid grid_nfce_cont 
            Height          =   3765
            Left            =   120
            TabIndex        =   87
            Top             =   1455
            Width           =   14790
            _Version        =   196617
            DataMode        =   2
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Col.Count       =   9
            CheckBox3D      =   0   'False
            AllowColumnSizing=   0   'False
            AllowGroupMoving=   0   'False
            AllowGroupSwapping=   0   'False
            AllowGroupShrinking=   0   'False
            AllowDragDrop   =   0   'False
            SelectTypeRow   =   1
            MaxSelectedRows =   5
            ForeColorEven   =   0
            BackColorEven   =   15724527
            BackColorOdd    =   12632319
            RowHeight       =   450
            ExtraHeight     =   265
            Columns.Count   =   9
            Columns(0).Width=   1111
            Columns(0).Caption=   "Enviar"
            Columns(0).Name =   "Enviar"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Style=   2
            Columns(1).Width=   1931
            Columns(1).Caption=   "Data"
            Columns(1).Name =   "Data"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   1852
            Columns(2).Caption=   "Sequência"
            Columns(2).Name =   "Sequencia"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   1164
            Columns(3).Caption=   "Serie"
            Columns(3).Name =   "Serie"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   2090
            Columns(4).Caption=   "Nota Fiscal"
            Columns(4).Name =   "NotaFiscal"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   2037
            Columns(5).Caption=   "Total"
            Columns(5).Name =   "Total"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   2990
            Columns(6).Caption=   "Status"
            Columns(6).Name =   "Status"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   8255
            Columns(7).Caption=   "ChaveAcesso"
            Columns(7).Name =   "ChaveAcesso"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   6985
            Columns(8).Caption=   "retNFCe"
            Columns(8).Name =   "retNFCe"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            _ExtentX        =   26088
            _ExtentY        =   6641
            _StockProps     =   79
            ForeColor       =   0
            BackColor       =   -2147483648
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Dia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   84
            Top             =   540
            Width           =   240
         End
         Begin VB.Label lblDataEntradaContingencia 
            AutoSize        =   -1  'True
            Caption         =   "Data Entrada em Contingência"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   -74610
            TabIndex        =   81
            Top             =   1755
            Width           =   2490
         End
         Begin VB.Label lblMotivoContingencia 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Contingência"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74610
            TabIndex        =   80
            Top             =   870
            Width           =   1650
         End
         Begin VB.Label lblHoraEntradaContingencia 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74610
            TabIndex        =   79
            Top             =   2655
            Width           =   375
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
         Height          =   585
         Left            =   -65490
         TabIndex        =   74
         Top             =   330
         Width           =   5745
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   -61560
         TabIndex        =   73
         Top             =   930
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   -65520
         TabIndex        =   72
         Top             =   6420
         Width           =   3885
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -74910
         TabIndex        =   71
         Top             =   6420
         Width           =   9345
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5475
         Left            =   -74910
         TabIndex        =   64
         Top             =   930
         Width           =   13275
         Begin VB.CheckBox chk_nfeInfAdProd 
            Appearance      =   0  'Flat
            Caption         =   "NFe > Considerar incluir a informação na TAG <InfAdProd>"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   108
            Top             =   210
            Width           =   4695
         End
         Begin VB.CommandButton cmd_salvarProdutos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Salvar informações dos produtos acima"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9585
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   4470
            Width           =   3555
         End
         Begin VB.CheckBox chk_xPed_nItemPed 
            Appearance      =   0  'Flat
            Caption         =   "NFe > Considerar incluir informação nas TAGs <xPed> e <nItemPed>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   495
            Width           =   6885
         End
         Begin VB.TextBox txt_sequencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   9960
            TabIndex        =   66
            Top             =   270
            Width           =   1365
         End
         Begin VB.CommandButton cmd_mostrarProdutos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mostrar Produtos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11625
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   255
            Width           =   1515
         End
         Begin SSDataWidgets_B.SSDBGrid grid_produtos 
            Height          =   3555
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   13020
            _Version        =   196617
            DataMode        =   2
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "WeblySleek UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Col.Count       =   4
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BackColorOdd    =   12632256
            RowHeight       =   423
            ExtraHeight     =   185
            Columns.Count   =   4
            Columns(0).Width=   3200
            Columns(0).Caption=   "Código do Produto"
            Columns(0).Name =   "CodigoProduto"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5239
            Columns(1).Caption=   "Nome"
            Columns(1).Name =   "Nome"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Caption=   "xPed"
            Columns(2).Name =   "xPed"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Caption=   "nItemPed"
            Columns(3).Name =   "nItemPed"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            _ExtentX        =   22966
            _ExtentY        =   6271
            _StockProps     =   79
            Caption         =   "Lista de Produtos da Sequência acima"
            ForeColor       =   -2147483630
            BackColor       =   -2147483648
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl_seq 
            Caption         =   "Sequência"
            Height          =   225
            Left            =   9120
            TabIndex        =   68
            Top             =   330
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   -74910
         TabIndex        =   62
         Top             =   330
         Width           =   9345
         Begin VB.CheckBox chk_nfeDevolucao_impostoDevol 
            Appearance      =   0  'Flat
            Caption         =   "NFe devolução > Considerar o grupo TAG <impostoDevol> ao invés de <IPI><IPITrib>...</IPITrib></IPI>"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   210
            Width           =   9075
         End
      End
      Begin VB.CheckBox chk_naoNFCe 
         Caption         =   "Não NFCe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14040
         TabIndex        =   56
         Top             =   900
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.TextBox txt_tipoDocCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74820
         TabIndex        =   53
         Text            =   "S"
         Top             =   6195
         Width           =   795
      End
      Begin VB.CommandButton cmd_imprimirDanfeCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Imprimir Danfe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -67260
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   6675
         Width           =   7425
      End
      Begin VB.TextBox txt_numNFeCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73935
         TabIndex        =   46
         Top             =   5625
         Width           =   1605
      End
      Begin VB.TextBox txt_serieCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74820
         TabIndex        =   45
         Top             =   5625
         Width           =   795
      End
      Begin VB.TextBox txt_descCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   -72030
         TabIndex        =   44
         Top             =   5625
         Width           =   12180
      End
      Begin VB.CommandButton cmd_pesquisarCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pesquisar NFe Cartas Correção"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   420
         Width           =   14985
      End
      Begin VB.CommandButton cmd_enviarCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enviar Carta Correção"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   6675
         Width           =   7425
      End
      Begin VB.CommandButton cmd_cancelarNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancelar NFe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10260
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   6210
         Width           =   4890
      End
      Begin VB.TextBox txtMotivoCancelamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   10260
         TabIndex        =   39
         Top             =   5730
         Width           =   4890
      End
      Begin VB.CommandButton cmd_posicaoConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Posição Consulta NFe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6750
         Width           =   4890
      End
      Begin VB.CommandButton cmd_enviarNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enviar NFe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6210
         Width           =   4890
      End
      Begin VB.CommandButton cmdPesquisar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   390
         Width           =   14970
      End
      Begin VB.CommandButton cmdPesquisarRetornos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pesquisar Erros/Críticas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   420
         Width           =   14985
      End
      Begin VB.CommandButton cmd_inutilizarNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Inutilizar NFe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6720
         Width           =   14985
      End
      Begin VB.CommandButton cmd_pesqInutilizadas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pesquisar NFe Inutilizadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   795
         Width           =   14985
      End
      Begin VB.TextBox txtMotivoInutilizacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68790
         TabIndex        =   26
         Top             =   6210
         Width           =   8895
      End
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73455
         TabIndex        =   25
         Top             =   6210
         Width           =   795
      End
      Begin VB.TextBox txt_anoInutilizacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74865
         TabIndex        =   24
         Top             =   6210
         Width           =   795
      End
      Begin VB.TextBox txtNumeroNFeFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70680
         TabIndex        =   23
         Top             =   6210
         Width           =   1605
      End
      Begin VB.TextBox txtNumeroNFeInicial 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72420
         TabIndex        =   22
         Top             =   6210
         Width           =   1605
      End
      Begin VB.CommandButton cmd_visualizaXML 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Visualizar XML"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6675
         Width           =   7335
      End
      Begin SSDataWidgets_B.SSDBGrid grdMovimento 
         Height          =   4215
         Left            =   210
         TabIndex        =   19
         Top             =   1290
         Width           =   14970
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   13
         CheckBox3D      =   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         MaxSelectedRows =   5
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   12648384
         RowHeight       =   450
         ExtraHeight     =   238
         Columns.Count   =   13
         Columns(0).Width=   1111
         Columns(0).Caption=   "Enviar"
         Columns(0).Name =   "Enviar"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   2143
         Columns(1).Caption=   "Data"
         Columns(1).Name =   "Data"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1852
         Columns(2).Caption=   "Sequência"
         Columns(2).Name =   "Sequencia"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2117
         Columns(3).Caption=   "Código"
         Columns(3).Name =   "CodigoClienteFornecedor"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   4895
         Columns(4).Caption=   "Nome Cliente/Fornecedor"
         Columns(4).Name =   "Nome Cliente/Fornecedor"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1164
         Columns(5).Caption=   "Serie"
         Columns(5).Name =   "Serie"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2090
         Columns(6).Caption=   "Nota Fiscal"
         Columns(6).Name =   "NotaFiscal"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   2037
         Columns(7).Caption=   "Total"
         Columns(7).Name =   "Total"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   2990
         Columns(8).Caption=   "Status"
         Columns(8).Name =   "Status"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   4895
         Columns(9).Caption=   "ChaveAcesso"
         Columns(9).Name =   "ChaveAcesso"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   3200
         Columns(10).Caption=   "ProtocoloAutorização"
         Columns(10).Name=   "ProtocoloAutorização"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   3545
         Columns(11).Caption=   "ProtocoloCancelamento"
         Columns(11).Name=   "ProtocoloCancelamento"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   6985
         Columns(12).Caption=   "arquivoDanfe"
         Columns(12).Name=   "arquivoDanfe"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         _ExtentX        =   26405
         _ExtentY        =   7435
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmd_imprimirDanfe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimir Danfe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6210
         Width           =   4890
      End
      Begin VB.CommandButton cmd_visualizarXML_AbaNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Visualizar XML"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6750
         Width           =   4890
      End
      Begin VB.CommandButton cmdMarcarDesmarcar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -150
         Picture         =   "frmNFe.frx":51F83
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   930
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CommandButton cmdMarcarDesmarcar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         DisabledPicture =   "frmNFe.frx":522B8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -195
         Picture         =   "frmNFe.frx":525F3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   810
         Visible         =   0   'False
         Width           =   315
      End
      Begin SSDataWidgets_B.SSDBGrid grdMovimento1 
         Height          =   750
         Left            =   15120
         TabIndex        =   6
         Top             =   7305
         Visible         =   0   'False
         Width           =   945
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GroupHeaders    =   0   'False
         Col.Count       =   13
         UseGroups       =   -1  'True
         DividerStyle    =   0
         BeveColorScheme =   0
         BevelColorFrame =   10066329
         BevelColorShadow=   15066597
         CheckBox3D      =   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   2
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorEven   =   15066597
         BackColorOdd    =   12648384
         RowHeight       =   423
         ExtraHeight     =   26
         Groups(0).Width =   29554
         Groups(0).HasHeadForeColor=   -1  'True
         Groups(0).HasHeadBackColor=   -1  'True
         Groups(0).HeadForeColor=   14671839
         Groups(0).HeadBackColor=   -2147483639
         Groups(0).Columns.Count=   13
         Groups(0).Columns(0).Width=   1005
         Groups(0).Columns(0).Caption=   "Enviar"
         Groups(0).Columns(0).Name=   "Enviar"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   8
         Groups(0).Columns(0).FieldLen=   256
         Groups(0).Columns(0).Style=   2
         Groups(0).Columns(0).HeadBackColor=   14671839
         Groups(0).Columns(1).Width=   1640
         Groups(0).Columns(1).Caption=   "Data"
         Groups(0).Columns(1).Name=   "Data"
         Groups(0).Columns(1).Alignment=   1
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   8
         Groups(0).Columns(1).FieldLen=   256
         Groups(0).Columns(1).Locked=   -1  'True
         Groups(0).Columns(2).Width=   1588
         Groups(0).Columns(2).Caption=   "Sequência"
         Groups(0).Columns(2).Name=   "Sequencia"
         Groups(0).Columns(2).Alignment=   1
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   8
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).Locked=   -1  'True
         Groups(0).Columns(3).Width=   1455
         Groups(0).Columns(3).Caption=   "Código"
         Groups(0).Columns(3).Name=   "CodigoClienteFornecedor"
         Groups(0).Columns(3).Alignment=   1
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   8
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).Locked=   -1  'True
         Groups(0).Columns(4).Width=   4577
         Groups(0).Columns(4).Caption=   "Nome Cliente/Fornecedor"
         Groups(0).Columns(4).Name=   "NomeClienteFornecedor"
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(4).Locked=   -1  'True
         Groups(0).Columns(5).Width=   926
         Groups(0).Columns(5).Caption=   "Serie"
         Groups(0).Columns(5).Name=   "Serie"
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(6).Width=   1958
         Groups(0).Columns(6).Caption=   "Nota Fiscal"
         Groups(0).Columns(6).Name=   "NotaFiscal"
         Groups(0).Columns(6).Alignment=   1
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(6).Locked=   -1  'True
         Groups(0).Columns(7).Width=   1905
         Groups(0).Columns(7).Caption=   "Total"
         Groups(0).Columns(7).Name=   "Total"
         Groups(0).Columns(7).Alignment=   1
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).DataType=   8
         Groups(0).Columns(7).FieldLen=   256
         Groups(0).Columns(7).Locked=   -1  'True
         Groups(0).Columns(8).Width=   2990
         Groups(0).Columns(8).Caption=   "Status"
         Groups(0).Columns(8).Name=   "Status"
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   8
         Groups(0).Columns(8).FieldLen=   256
         Groups(0).Columns(8).Locked=   -1  'True
         Groups(0).Columns(9).Width=   4419
         Groups(0).Columns(9).Caption=   "ChaveAcesso"
         Groups(0).Columns(9).Name=   "ChaveAcesso"
         Groups(0).Columns(9).DataField=   "Column 9"
         Groups(0).Columns(9).DataType=   8
         Groups(0).Columns(9).FieldLen=   256
         Groups(0).Columns(10).Width=   2963
         Groups(0).Columns(10).Caption=   "ProtocoloAutorização"
         Groups(0).Columns(10).Name=   "ProtocoloAutorização"
         Groups(0).Columns(10).DataField=   "Column 10"
         Groups(0).Columns(10).DataType=   8
         Groups(0).Columns(10).FieldLen=   256
         Groups(0).Columns(11).Width=   2249
         Groups(0).Columns(11).Caption=   "ProtocoloCancelamento"
         Groups(0).Columns(11).Name=   "ProtocoloCancelamento"
         Groups(0).Columns(11).DataField=   "Column 11"
         Groups(0).Columns(11).DataType=   8
         Groups(0).Columns(11).FieldLen=   256
         Groups(0).Columns(12).Width=   1879
         Groups(0).Columns(12).Caption=   "arquivoDanfe"
         Groups(0).Columns(12).Name=   "arquivoDanfe"
         Groups(0).Columns(12).DataField=   "Column 12"
         Groups(0).Columns(12).DataType=   8
         Groups(0).Columns(12).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   1667
         _ExtentY        =   1323
         _StockProps     =   79
         ForeColor       =   -2147483630
         BackColor       =   15724527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid grdRetorno 
         Height          =   615
         Left            =   -62535
         TabIndex        =   21
         Top             =   7575
         Visible         =   0   'False
         Width           =   14970
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   13
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   2
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   12648447
         RowHeight       =   503
         ExtraHeight     =   26
         Columns.Count   =   13
         Columns(0).Width=   1931
         Columns(0).Caption=   "Nota Fiscal"
         Columns(0).Name =   "NotaFiscal"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1693
         Columns(1).Caption=   "Sequência"
         Columns(1).Name =   "Sequencia"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1535
         Columns(2).Caption=   "Tipo"
         Columns(2).Name =   "Tipo"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "Nome Arquivo"
         Columns(3).Name =   "NomeArquivo"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "Data Hora Arquivo"
         Columns(4).Name =   "DataHoraArquivo"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Data Hora Processamento"
         Columns(5).Name =   "DataHoraProcessamento"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2514
         Columns(6).Caption=   "Código Resposta"
         Columns(6).Name =   "CodigoResposta"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   11668
         Columns(7).Caption=   "Descrição Resposta"
         Columns(7).Name =   "DescricaoResposta"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Caption=   "Protocolo Autorização"
         Columns(8).Name =   "ProtocoloAutorizacao"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3413
         Columns(9).Caption=   "Protocolo Cancelamento"
         Columns(9).Name =   "ProtocoloCancelamento"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   3200
         Columns(10).Caption=   "Digest Value"
         Columns(10).Name=   "DigestValue"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   1535
         Columns(11).Caption=   "Cliente"
         Columns(11).Name=   "Cliente"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   3200
         Columns(12).Caption=   "Nome Cliente"
         Columns(12).Name=   "NomeCliente"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   26405
         _ExtentY        =   1085
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "WeblySleek UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid grdInutilizadas 
         Height          =   4620
         Left            =   -74865
         TabIndex        =   27
         Top             =   1305
         Width           =   14970
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   7
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   2
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   14671839
         RowHeight       =   450
         ExtraHeight     =   79
         Columns.Count   =   7
         Columns(0).Width=   3200
         Columns(0).Caption=   "CNPJ"
         Columns(0).Name =   "CNPJ"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Ano"
         Columns(1).Name =   "Ano"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Serie"
         Columns(2).Name =   "Serie"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Caption=   "NumeroInicial"
         Columns(3).Name =   "NumeroInicial"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "NumeroFinal"
         Columns(4).Name =   "NumeroFinal"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   6324
         Columns(5).Caption=   "Justificativa"
         Columns(5).Name =   "Justificativa"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Caption=   "DataHora"
         Columns(6).Name =   "DataHora"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   26405
         _ExtentY        =   8149
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid grdCC 
         Height          =   4215
         Left            =   -74820
         TabIndex        =   47
         Top             =   960
         Width           =   14970
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   6
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   2
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   16777152
         RowHeight       =   450
         ExtraHeight     =   79
         Columns.Count   =   6
         Columns(0).Width=   3200
         Columns(0).Caption=   "DataHora"
         Columns(0).Name =   "DataHora"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3387
         Columns(1).Caption=   "CNPJ"
         Columns(1).Name =   "CNPJ"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1138
         Columns(2).Caption=   "Serie"
         Columns(2).Name =   "Serie"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1746
         Columns(3).Caption=   "Número"
         Columns(3).Name =   "Número"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   11324
         Columns(4).Caption=   "Descrição"
         Columns(4).Name =   "Descrição"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   5027
         Columns(5).Caption=   "arquivoDanfeCC"
         Columns(5).Name =   "arquivoDanfeCC"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   26405
         _ExtentY        =   7435
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridRetorno 
         Height          =   3780
         Left            =   -74865
         TabIndex        =   55
         Top             =   960
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   6668
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedCols       =   0
         BackColor       =   -2147483648
         BackColorFixed  =   12632256
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483648
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox msk_dataDiaNFCe 
         Height          =   315
         Left            =   -73980
         TabIndex        =   95
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
      Begin SSDataWidgets_B.SSDBGrid grd_nfceNormal 
         Height          =   3765
         Left            =   -74790
         TabIndex        =   96
         Top             =   1395
         Width           =   14790
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   9
         CheckBox3D      =   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         MaxSelectedRows =   5
         ForeColorEven   =   0
         BackColorEven   =   15724527
         BackColorOdd    =   12632256
         RowHeight       =   450
         ExtraHeight     =   159
         Columns.Count   =   9
         Columns(0).Width=   1111
         Columns(0).Caption=   "Enviar"
         Columns(0).Name =   "Enviar"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   2117
         Columns(1).Caption=   "Data"
         Columns(1).Name =   "Data"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1852
         Columns(2).Caption=   "Sequência"
         Columns(2).Name =   "Sequencia"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1164
         Columns(3).Caption=   "Serie"
         Columns(3).Name =   "Serie"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   2090
         Columns(4).Caption=   "Nota Fiscal"
         Columns(4).Name =   "NotaFiscal"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "Status"
         Columns(5).Name =   "Status"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   2037
         Columns(6).Caption=   "Total"
         Columns(6).Name =   "Total"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   8493
         Columns(7).Caption=   "ChaveAcesso"
         Columns(7).Name =   "ChaveAcesso"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   14579
         Columns(8).Caption=   "retNFCe"
         Columns(8).Name =   "retNFCe"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         _ExtentX        =   26088
         _ExtentY        =   6641
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox msk_dataRetroativa 
         Height          =   315
         Left            =   2610
         TabIndex        =   105
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   5580
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin VB.Label Label10 
         Caption         =   "Operação"
         Height          =   255
         Left            =   6960
         TabIndex        =   110
         Top             =   930
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Visualizar"
         Height          =   255
         Left            =   180
         TabIndex        =   102
         Top             =   930
         Width           =   795
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Dia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -74310
         TabIndex        =   97
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lbl_tipoDoc 
         AutoSize        =   -1  'True
         Caption         =   "Tp Doc: 'S'Saída  ou 'E'Entrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -74820
         TabIndex        =   54
         Top             =   5955
         Width           =   2550
      End
      Begin VB.Label lblNumeroNFeInicial 
         AutoSize        =   -1  'True
         Caption         =   "NFe Número"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -73950
         TabIndex        =   50
         Top             =   5385
         Width           =   1020
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -74820
         TabIndex        =   49
         Top             =   5385
         Width           =   405
      End
      Begin VB.Label lblMotivoInutilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição Carta Correção"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -72030
         TabIndex        =   48
         Top             =   5385
         Width           =   2025
      End
      Begin VB.Label lblMotivoCancelamento 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Cancelamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   10260
         TabIndex        =   41
         Top             =   5490
         Width           =   1755
      End
      Begin VB.Label lblMotivoInutilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Inutilização"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -68790
         TabIndex        =   32
         Top             =   5970
         Width           =   1485
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -73455
         TabIndex        =   31
         Top             =   5970
         Width           =   405
      End
      Begin VB.Label lblNumeroNFeInicial 
         AutoSize        =   -1  'True
         Caption         =   "NFe Número Inicial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -72420
         TabIndex        =   30
         Top             =   5970
         Width           =   1515
      End
      Begin VB.Label lblNumeroNFeFinal 
         AutoSize        =   -1  'True
         Caption         =   "NFe Número Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   -70680
         TabIndex        =   29
         Top             =   5970
         Width           =   1425
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         Caption         =   "Ano  (Ex: 2018)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -74865
         TabIndex        =   28
         Top             =   5970
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkConcatenarDescricaoAdicional 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Concatenar Descrição Adicional ao nome do Produto"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   45
      TabIndex        =   9
      Top             =   8025
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ComboBox cboTipoMovimento 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmNFe.frx":5292E
      Left            =   1530
      List            =   "frmNFe.frx":52938
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   1830
   End
   Begin VB.ComboBox cboTipoEmissao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmNFe.frx":5294E
      Left            =   14535
      List            =   "frmNFe.frx":5295E
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   75
      Visible         =   0   'False
      Width           =   885
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   495
      Left            =   15210
      TabIndex        =   11
      Top             =   6660
      Visible         =   0   'False
      Width           =   495
      ExtentX         =   873
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSMask.MaskEdBox Data_Fim 
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   105
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Left            =   3975
      TabIndex        =   1
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   105
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00999999&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   13260
      TabIndex        =   91
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8880
      TabIndex        =   89
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   11040
      TabIndex        =   61
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label lbl_statusProcessamento 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Em processo de transmissão da NFe para a SEFAZ ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10530
      TabIndex        =   59
      Top             =   8115
      Visible         =   0   'False
      Width           =   4875
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4500
      TabIndex        =   51
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6690
      TabIndex        =   18
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2310
      TabIndex        =   17
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   135
      TabIndex        =   16
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Movimento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   135
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   3705
      TabIndex        =   14
      Top             =   135
      Width           =   225
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "até"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   5925
      TabIndex        =   13
      Top             =   135
      Width           =   270
   End
   Begin VB.Label lblTipoEmissao 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Emissão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   13920
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "frmNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'14/09/2009 - mpdea
'Tela para envio e retorno de NFe

Option Explicit


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long
 
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2


Private m_intTipoMovimentoPesquisa As Integer
'24/11/2010 - Andrea
Private m_intTipoEmissao As Integer
Dim sPastaEnvioNfe As String
Dim sPadraoArquivoIntegracao As String

Dim sCNPJ As String



''##############################################################
'' PABLO - 14/10/2022
''##############################################################
' Abrir movimento direto da tela de saídas
Dim param_sequencia As Long

Public Sub SetParametros(ByVal p_sequencia As Long)
    param_sequencia = p_sequencia
End Sub
''##############################################################




'Private Sub ActiveBar_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
'  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
'
'  '24/11/2010 - Andrea
'  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
'
'  Select Case Tool.Name
'    Case "miEnviarNotaFiscal"
'      Call EnviarNFe
'
'    Case "miCancelarNotaFiscal"
'      Call CancelarNFe
'
'    Case "miInutilizarNotasFiscais"
'      Call InutilizarNFe
'
'    Case "miProcessarRetorno"
'      Call ProcessarRetorno
'
'      '08/09/2011 - Andrea
'      'Depois de processar os retornos joga os resultados na tela
'      Call PesquisarRetornos
'
'  End Select
'
'End Sub

Private Sub ProcessarRetornoNFCe()
  On Error GoTo ErrHandler
  
  Dim sChaveAcessoCont As String
  Dim bTemAlgumRegistroSelecionadoNaGrid As Boolean
  Dim strSQL As String
  Dim sRetorno As String
  Dim bm As Variant
  Dim nRow As Long
  Dim rsParametros As Recordset
  Dim rsSaidaSEFAZ As Recordset
  Dim iIndice1 As Long
  Dim iIndice2 As Long
  Dim sStatus As String
  Dim sSequencia As String
  Dim sDetalheAutorizacao As String
  Dim sExMessage As String
    
  bTemAlgumRegistroSelecionadoNaGrid = False
  
'  'Parâmetros da Filial
'  strSQL = "SELECT CGC FROM [Parâmetros Filial] "
'  strSQL = strSQL & "WHERE Filial = " & gnCodFilial
'  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
'  With rsParametros
'    If (.BOF And .EOF) Then
'      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
'      .Close
'      Set rsParametros = Nothing
'      Exit Sub
'    Else
'      'Validações
'      sCNPJ = Trim(.Fields("CGC").Value)
'      sCNPJ = Replace(sCNPJ, "/", "")
'      sCNPJ = Replace(sCNPJ, "-", "")
'      sCNPJ = Replace(sCNPJ, " ", "")
'      sCNPJ = Replace(sCNPJ, ".", "")
'      sCNPJ = Replace(sCNPJ, ",", "")
'    End If
'  End With

  ' Via SOAP
  If bSoapClient_MSSoapInit_NFCe = False Then
    Set soapclient_NFCe = New SoapClient30
    soapclient_NFCe.MSSoapInit sSoapClient_MSSoapInit_NFCe
    soapclient_NFCe.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL_NFCe
    bSoapClient_MSSoapInit_NFCe = True
  End If
  
  With grid_nfce_cont
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)
          
          If CBool(.Columns("Enviar").CellValue(bm)) Then
          
              sChaveAcessoCont = gsHandleNull(.Columns("ChaveAcesso").CellValue(bm))
              sSequencia = .Columns("Sequencia").CellValue(bm)
              
              bTemAlgumRegistroSelecionadoNaGrid = True
            
              'Chama WS
              sRetorno = soapclient_NFCe.GetStatusPorChave_Xml(sCNPJ, sChaveAcessoCont)
              
              sRetorno = Replace(sRetorno, vbCrLf, "")

              iIndice1 = InStr(1, sRetorno, "<statusAutorizacao>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetorno, "</statusAutorizacao>")
                sStatus = Mid(sRetorno, iIndice1 + 19, iIndice2 - (iIndice1 + 19))
              Else
                sStatus = ""
              End If
              
              iIndice1 = InStr(1, sRetorno, "<detalheAutorizacao>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetorno, "</detalheAutorizacao>")
                sDetalheAutorizacao = Mid(sRetorno, iIndice1 + 20, iIndice2 - (iIndice1 + 20))
              Else
                sDetalheAutorizacao = ""
              End If
              
              iIndice1 = InStr(1, sRetorno, "<exMessage>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetorno, "</exMessage>")
                sExMessage = Mid(sRetorno, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
              Else
                sExMessage = ""
              End If
              
              If InStr(1, sRetorno, "<detalheAutorizacao>100") > -1 Then
                  'Autorizado após processamento em contingencia
                  sStatus = "OK"
              Else
                  MsgBox "NFCe com ERRO:" & vbCrLf & sDetalheAutorizacao & vbCrLf & sExMessage, vbInformation, "NFCe Posição de retorno"
                  sStatus = "Rejeitado/Pendente"
              End If
              
              'Atualizar tabela Saídas e a Grid
              Set rsSaidaSEFAZ = db.OpenRecordset("Select * from [Saídas] where Filial = " & gnCodFilial & " And Sequência = " & sSequencia & "")
      
              rsSaidaSEFAZ.Edit
              rsSaidaSEFAZ!retNFCe_contingencia = sRetorno
              rsSaidaSEFAZ!NFCe_contingencia_status = sStatus
              'rsSaidaSEFAZ!NFCe_contingencia_num = sNumNFCeRetCont
              'rsSaidaSEFAZ!NFCe_contingencia_serie = sSerieNFCeRetCont
              'rsSaidaSEFAZ!NFCe_contingencia_chave = sChaveNFCeRetCont
              rsSaidaSEFAZ.Update
              rsSaidaSEFAZ.Close
              Set rsSaidaSEFAZ = Nothing
          End If
      Next nRow
  End With

  If bTemAlgumRegistroSelecionadoNaGrid = False Then
      MsgBox "Selecione uma NFCe na grade.", vbInformation, "Informação"
  End If

  Exit Sub

ErrHandler:
  MsgBox "Erro ao processar Retorno : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"

End Sub

Private Sub ProcessarRetorno()
  Dim strSQL As String
  Dim rsParametros As Recordset
  Dim rsSaidas As Recordset
  Dim rsCliFor As Recordset
  Dim rsCadastroTransportadora As Recordset
  Dim rsNFe As Recordset
  Dim rsNFeRetorno As Recordset
  Dim rsNFeXML As Recordset
  Dim strVersaoLayoutEnvio As String
  Dim strPadraoArquivoIntegracao As String
  Dim strPastaRetorno As String
  Dim bTemAlgumRegistroSelecionadoNaGrid As Boolean
  
  Dim iSerieNFe As Integer

  '05/04/2010 - mpdea
  Dim str_split_linha() As String
  Dim str_descricao_status2_resposta As String
  Dim str_Retorno As String
  
  On Error GoTo ErrHandler
  
  bTemAlgumRegistroSelecionadoNaGrid = False
  
   
  'Parâmetros da Filial
  strSQL = "SELECT CGC, [Razão Social], Nome, Inscrição, Endereço, EnderecoNumero, EnderecoComplemento, "
  strSQL = strSQL & "Bairro, Cidade, Estado, CEP, Pais, Fone, "
  strSQL = strSQL & "AmbienteNfe, FormatoImpressaoDanfeNfe, ModDetBaseCalculoIcms, ModDetBaseCalculoIcmsSt, "
  strSQL = strSQL & "PastaEnvioNfe, PastaRetornoNfe, VersaoLayoutEnvio, PadraoArquivoIntegracao "
  strSQL = strSQL & "FROM [Parâmetros Filial] "
  strSQL = strSQL & "WHERE Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsParametros
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
      .Close
      Set rsParametros = Nothing
      Exit Sub
    Else
      'Validações
      sCNPJ = Trim(.Fields("CGC").Value)
      sCNPJ = Replace(sCNPJ, "/", "")
      sCNPJ = Replace(sCNPJ, "-", "")
      sCNPJ = Replace(sCNPJ, " ", "")
      sCNPJ = Replace(sCNPJ, ".", "")
      sCNPJ = Replace(sCNPJ, ",", "")

      strVersaoLayoutEnvio = Trim(.Fields("VersaoLayoutEnvio").Value)
      If strVersaoLayoutEnvio = "" Then
        strVersaoLayoutEnvio = "4.00"
      End If
      strPadraoArquivoIntegracao = Trim(.Fields("PadraoArquivoIntegracao").Value)
      
      If strPadraoArquivoIntegracao = "TXT" Then
          strPastaRetorno = Trim(.Fields("PastaRetornoNfe").Value)
          If strPastaRetorno = "" Then
            MsgBox "Pasta de retorno não configurada para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
            .Close
            Set rsParametros = Nothing
            Exit Sub
          Else
            If Right(strPastaRetorno, 1) <> "\" Then
              strPastaRetorno = strPastaRetorno & "\"
            End If
          End If
      End If
    End If
  End With
  
  Dim lngNumNFe As Long
  Dim lngOID_xmlLoteBenefix As Long
  Dim lngRet As Long

  If strPadraoArquivoIntegracao = "XML" Then
    
    'CRIA O SOAPCLIENTE E FAZ AS DEVIDAS CONFIGURAÇÕES
    If bSoapClient_MSSoapInit = False Then
      Set soapclient = New SoapClient30
      soapclient.MSSoapInit sSoapClient_MSSoapInit
      soapclient.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL
      bSoapClient_MSSoapInit = True
    End If
    
'    Dim soapclient As SoapClient30
'    Set soapclient = New SoapClient30
    
'    soapclient.MSSoapInit sSoapClient_MSSoapInit
'    soapclient.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL
    
    'CRIA VARIAVEIS COM INFORMAÇÕES NECESSÁRIAS PARA CHAMADA DO MÉTODO
    Dim rsPesq As Recordset
    Dim usuario As String
    Dim Senha As String
    Dim indicaContingencia As Boolean
    Dim nomeArquivoDanfe As String
    Dim retNFeSec As String
    Dim retProtNFe As String
    Dim dhRecbto As String
    Dim cStat As String
    Dim nProt As String
    Dim iIndice As Integer
    Dim iIndice1 As Integer
    Dim iIndice2 As Integer
    Dim bm As Variant
    Dim nRow As Long
    Dim lngSequencia As Long
    Dim sStatusDescricao As String
    Dim sStatusDescricao2 As String
    Dim bChamouGetOid As Boolean

    'ESTAMOS USANDO O CNPJ DA BENEFIX PARA REALIZAR O TESTE
    'cnpj = "06039615000108"
    usuario = "TaskService"
    Senha = "4sjvpuKl7+Hns//ijN7o"

    With grdMovimento
        For nRow = 0 To .Rows - 1
            bm = .AddItemBookmark(nRow)
          
            If CBool(.Columns("Enviar").CellValue(bm)) Then
                bTemAlgumRegistroSelecionadoNaGrid = True
                lngSequencia = CLng(gsHandleNull(.Columns("Sequencia").CellValue(bm)))
                
                strSQL = "Select oid_xmlLoteBenefix, Numero, Serie FROM [NFe] where Sequencia=" & lngSequencia & " And Filial=" & gnCodFilial
                Set rsNFeXML = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
                If Not (rsNFeXML.BOF And rsNFeXML.EOF) Then
                    'If rsNFeXML.Fields("oid_xmlLoteBenefix").Value <> 0 And Not IsNull(rsNFeXML.Fields("oid_xmlLoteBenefix").Value) Then
                    If Not IsNull(rsNFeXML.Fields("oid_xmlLoteBenefix").Value) And rsNFeXML.Fields("oid_xmlLoteBenefix").Value <> "" Then
                        lngOID_xmlLoteBenefix = rsNFeXML.Fields("oid_xmlLoteBenefix").Value
                        lngNumNFe = rsNFeXML.Fields("Numero").Value
                        iSerieNFe = rsNFeXML.Fields("Serie").Value
                    End If
                End If
                rsNFeXML.Close
                Set rsNFeXML = Nothing
                
                bChamouGetOid = False

                If lngOID_xmlLoteBenefix <= 0 Then
                    ' Retorno:
                    '       Maior que 0 - Numero do OID
                    '       Menor ou igual a zero - Erro (código)
                    '       -1001;-1004,-1016 - erros de cadastro
                    '       -1007 - Documento não encontrado
                    
                    bChamouGetOid = True
                    lngOID_xmlLoteBenefix = soapclient.GetOid(sCNPJ, 0, usuario, Senha, iSerieNFe, lngNumNFe)
                    
                    If lngOID_xmlLoteBenefix > 0 Then
                        lngRet = soapclient.SolicitarImpressaoNFeProt(sCNPJ, 0, usuario, Senha, lngOID_xmlLoteBenefix, lngNumNFe, indicaContingencia, nomeArquivoDanfe, retNFeSec, retProtNFe)
                    Else
                        lngRet = lngOID_xmlLoteBenefix
                    End If
                Else
                    'Chama WS
                    lngRet = soapclient.SolicitarImpressaoNFeProt(sCNPJ, 0, usuario, Senha, lngOID_xmlLoteBenefix, lngNumNFe, indicaContingencia, nomeArquivoDanfe, retNFeSec, retProtNFe)
                End If
        
                If lngRet = 1 Then    'Processou com sucesso, neste caso, cStat 100 AUTORIZADA
                  
                  'O Código/cStat, se não tiver ‘retProtNFe’, pesquise em ‘retNFeSec’
                  If Not IsNull(retProtNFe) And retProtNFe <> "" Then
                  
                    '<cStat>100</cStat>
                    iIndice1 = InStr(1, retProtNFe, "<cStat>")
                    If iIndice1 > 0 Then
                      iIndice2 = InStr(1, retProtNFe, "</cStat>")
                      cStat = Mid(retProtNFe, iIndice1 + 7, iIndice2 - (iIndice1 + 7))
                    End If
                    
                    '<nProt>327100000066004</nProt>
                    iIndice1 = InStr(1, retProtNFe, "<nProt>")
                    If iIndice1 > 0 Then
                      iIndice2 = InStr(1, retProtNFe, "</nProt>")
                      nProt = Mid(retProtNFe, iIndice1 + 7, iIndice2 - (iIndice1 + 7))
                    End If
                    
                    '<dhRecbto>2010-05-10T13:53:54</dhRecbto>
                    iIndice1 = InStr(1, retProtNFe, "<dhRecbto>")
                    If iIndice1 > 0 Then
                      iIndice2 = InStr(1, retProtNFe, "</dhRecbto>")
                      dhRecbto = Mid(retProtNFe, iIndice1 + 10, iIndice2 - (iIndice1 + 10))
                      dhRecbto = Replace(dhRecbto, "T", " ")
                      dhRecbto = Replace(dhRecbto, "-01:00", "")
                      dhRecbto = Replace(dhRecbto, "-02:00", "")
                      dhRecbto = Replace(dhRecbto, "-03:00", "")
                      dhRecbto = Replace(dhRecbto, "-04:00", "")
                    End If
                    
                  Else
                    iIndice1 = InStr(1, retNFeSec, "<Codigo>")
                    If iIndice1 > 0 Then
                      iIndice2 = InStr(1, retNFeSec, "</Codigo>")
                      cStat = Mid(retNFeSec, iIndice1 + 8, iIndice2 - (iIndice1 + 8)) ' <Codigo>1408</Codigo> // Remove o sinal '-'
                      cStat = Replace(cStat, "-", "")
                    End If
                  End If
         
                  If bChamouGetOid = True Then
                      'Atualizar tabela NFe e a Grid
                      strSQL = "UPDATE NFe Set status=" & cStat & ", ProtocoloAutorizacao='" & nProt & "',"
                      strSQL = strSQL & " DataHoraAutorizacao='" & dhRecbto & "', "
                      strSQL = strSQL & " nomeDanfe='" & nomeArquivoDanfe & "',oid_xmlLoteBenefix=" & lngOID_xmlLoteBenefix
                      strSQL = strSQL & " Where Sequencia=" & lngSequencia
                      strSQL = strSQL & " And Numero=" & lngNumNFe & " And Filial=" & gnCodFilial
                      strSQL = strSQL & " And Modelo ='55' "
                  Else
                      'Atualizar tabela NFe e a Grid
                      strSQL = "UPDATE NFe Set status=" & cStat & ", ProtocoloAutorizacao='" & nProt & "',"
                      strSQL = strSQL & " DataHoraAutorizacao='" & dhRecbto & "', "
                      strSQL = strSQL & " nomeDanfe='" & nomeArquivoDanfe & "' "
                      strSQL = strSQL & " Where Sequencia=" & lngSequencia
                      strSQL = strSQL & " And Numero=" & lngNumNFe & " And Filial=" & gnCodFilial
                      strSQL = strSQL & " And Modelo ='55' And oid_xmlLoteBenefix=" & lngOID_xmlLoteBenefix
                  End If
                  db.Execute strSQL
                  
                  bChamouGetOid = False
              Else
                  '*************************
                  'Erro/Inconsistencia
              
                  strSQL = "UPDATE NFe Set oid_xmlLoteBenefix=" & lngRet & ", Status=" & lngRet & " Where Sequencia=" & lngSequencia & " And Numero=" & lngNumNFe & " And Filial=" & gnCodFilial
                  db.Execute strSQL
        
                  iIndice = InStr(1, retNFeSec, "<ErroXML>")
                  If iIndice > 0 Then
                    iIndice2 = InStr(iIndice, retNFeSec, "</ErroXML>")
                    sStatusDescricao = Mid$(retNFeSec, iIndice + 9, iIndice2 - (iIndice + 9))
                    
                    If Len(sStatusDescricao) > 255 Then
                      sStatusDescricao = Mid$(sStatusDescricao, 1, 255)
                    End If
                  End If
                  sStatusDescricao = TrataCaracteresEspeciaisASCII_Traducao(sStatusDescricao)
                  
                  iIndice = InStr(1, retNFeSec, "<Mensagem>")
                  If iIndice > 0 Then
                    iIndice2 = InStr(iIndice, retNFeSec, "</Mensagem>")
                    sStatusDescricao2 = Mid$(retNFeSec, iIndice + 10, iIndice2 - (iIndice + 10))
                    
                    If Len(sStatusDescricao2) > 255 Then
                      sStatusDescricao2 = Mid$(sStatusDescricao2, 1, 255)
                    End If
                  End If
                  
                  strSQL = "SELECT * FROM [NFeRetorno] "
                  strSQL = strSQL & "WHERE Filial = " & gnCodFilial & " And Sequencia=" & lngSequencia & " And TipoMovimento=0"
                  Set rsPesq = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
                  
                  If sStatusDescricao = "" Then
                      sStatusDescricao = sStatusDescricao2
                  End If
                  
                  If sStatusDescricao = "" Then
                    sStatusDescricao = "Erro codigo: " & CStr(lngRet)
                    sStatusDescricao2 = "Erro codigo: " & CStr(lngRet)
                  End If
                    
                  If (rsPesq.BOF And rsPesq.EOF) Then
                    strSQL = "Insert into NFeRetorno (Filial, Sequencia, TipoMovimento, DataHora, Protocolo, "
                    strSQL = strSQL & " StatusDescricao, StatusDescricao2, status)"
                    strSQL = strSQL & " values(" & gnCodFilial & "," & lngSequencia & ",0,'" & Now & "','',"
                    strSQL = strSQL & "'" & sStatusDescricao & "','" & sStatusDescricao2 & "', " & lngRet & ")"
                    
                    db.Execute strSQL
                  Else
                    strSQL = "Update NFeRetorno set DataHora= '" & Now & "', StatusDescricao='" & sStatusDescricao & "',"
                    strSQL = strSQL & " StatusDescricao2='" & sStatusDescricao2 & "', status=" & lngRet & " "
                    strSQL = strSQL & " Where Filial=" & gnCodFilial & " And Sequencia=" & lngSequencia & " And TipoMovimento=0 "
                    
                    db.Execute strSQL
                  End If
                  rsPesq.Close
                  Set rsPesq = Nothing
        
                  'MsgBox "Inconsistência na EMISSÃO da NFe: " & pNumero & "  CódigoErro: " & retOID & " -> Mais detalhes, Botão 'Pesquise Retornos' e Visualize na Aba 'Retornos'", vbInformation
              End If
          End If
        Next nRow
    End With
  ElseIf strPadraoArquivoIntegracao = "TXT" Then
      Dim FSys As New FileSystemObject
      
      Dim pasta As Folder
      Dim TodosArquivos As Files
      Dim ARQUIVO As File
      Dim Instream As TextStream
      Dim str_linhaTexto As String
      Dim Str_Aux As String
          
      Set pasta = FSys.GetFolder(strPastaRetorno)
      Set TodosArquivos = pasta.Files
        
      For Each ARQUIVO In TodosArquivos
        Set Instream = FSys.OpenTextFile(ARQUIVO, ForReading, False, TristateFalse)
        
        'If InStr(Arquivo.Name, "ENVIO") > 0 Then 'Para ler somente os arquivos de retorno de envio.
        
          While Instream.AtEndOfStream = False
          
            str_linhaTexto = Instream.ReadLine
            
            Dim str_identificador_TAG  As String
            Dim str_chave_acesso As String
            Dim str_data_hora_processamento As String
            Dim str_numero_protocolo_autorizacao As String
            Dim str_digest_value As String
            Dim str_codigo_status_resposta As String
            Dim int_tamanho_que_falta As Integer
            Dim str_descricao_status_resposta As String
            Dim str_dia As String
            Dim str_mes As String
            Dim str_ano As String
            Dim Str_Data As String
            
            str_identificador_TAG = Mid(str_linhaTexto, 1, 4)
            
            'Para  as linhas de TAG 1100, gravar registro na tabela de retorno
            'Ifnot str_identificador_TAG = "1000" Then
            
            If str_linhaTexto = "" Then
              GoTo DeletarArquivo
            End If
            
            str_split_linha = Split(str_linhaTexto, ";")
            Select Case str_split_linha(1)
            Case "retImprNFe"
              GoTo ProximoArquivo
            Case "envEvento"
              GoTo ProximoArquivo
            Case "inutNFe"
              GoTo ProximoArquivo
            Case "retInutNFe"
              GoTo ProximoArquivo
            Case "conSitNFe"
              GoTo ProximoArquivo
            End Select
            
            If str_identificador_TAG = "1000" Then
              str_chave_acesso = Mid(str_linhaTexto, 8, 44)
            End If
            
            If str_identificador_TAG = "1100" Then
              Str_Aux = Mid(str_linhaTexto, 8, 12)
              Dim vetChave As Variant
              vetChave = Split(str_linhaTexto, ";")
              If Len(str_chave_acesso) <> 44 Then
                
                str_chave_acesso = Right(vetChave(3), 44)
              End If
              
              If Len(str_chave_acesso) <> 44 Then
                str_chave_acesso = vetChave(6)
                
              End If
              
              If Len(str_chave_acesso) <> 44 Then
                str_chave_acesso = vetChave(7)
                
              End If
              
              If Str_Aux <> "InfinitriNFe" Then
                        
                'NFe
                strSQL = "SELECT ChaveAcesso, ProtocoloAutorizacao, Numero, Sequencia, DataHoraAutorizacao, DataHoraCancelamento, ProtocoloCancelamento  "
                strSQL = strSQL & "FROM NFe "
                strSQL = strSQL & "WHERE Filial = " & gnCodFilial & " AND ChaveAcesso = '" & str_chave_acesso & "' AND TipoMovimento = 1" 'TODO: tem que ser variel   m_intTipoMovimentoPesquisa
                Set rsNFe = db.OpenRecordset(strSQL, dbOpenDynaset)
                With rsNFe
                  If (.BOF And .EOF) Then
                    MsgBox "NFe não localizada, chave de acesso número : " & str_chave_acesso, vbExclamation, "Atenção"
                    .Close
                    Set rsNFe = Nothing
                    '06/03/2013-Alexandre Afornali
                    'Alterado para ir para proximo arquivo ao inves de sair da função
                    GoTo ProximoArquivo
                  Else
                    'Validações
                    str_numero_protocolo_autorizacao = Trim(.Fields("ProtocoloAutorizacao").Value)
                    If str_numero_protocolo_autorizacao <> "" And vetChave(7) <> "110111" Then
                      MsgBox "NFe já autorizada. Este arquivo de retorno já foi processado, favor verificar. NFe número : " & .Fields("Numero").Value, vbExclamation, "Atenção"
                      .Close
                      Instream.Close
                      FSys.DeleteFile ARQUIVO, True
                      Set rsNFe = Nothing
                      '06/03/2013-Alexandre Afornali
                      'Alterado para ir para proximo arquivo ao inves de sair da função
                      GoTo ProximoArquivo
                    End If
                  End If
                End With
                '05/04/2010 - mpdea
                'Divide a linha de retorno em array de campos
                str_split_linha = Split(str_linhaTexto, ";")
                
                '05/04/2010 - mpdea
                'Verifica retorno
                If UBound(str_split_linha) > 7 Then
                  If Len(str_split_linha(4)) > 25 Then
                    str_data_hora_processamento = str_split_linha(7)
                    str_descricao_status_resposta = str_split_linha(4)
                    str_codigo_status_resposta = str_split_linha(3)
                    If Len(str_descricao_status_resposta) > 255 Then
                      str_descricao_status2_resposta = Mid(str_descricao_status_resposta, 256, 255)
                      str_descricao_status_resposta = Left(str_descricao_status_resposta, 255)
                    End If
                    str_Retorno = str_split_linha(4)
                  Else
                    str_data_hora_processamento = str_split_linha(4)
                    str_numero_protocolo_autorizacao = str_split_linha(5)
                    str_digest_value = str_split_linha(6)
                    str_codigo_status_resposta = str_split_linha(7)
                    str_descricao_status_resposta = str_split_linha(8)
                    str_Retorno = str_split_linha(8)
                    If Len(str_descricao_status_resposta) > 255 Then
                      str_descricao_status2_resposta = Mid(str_descricao_status_resposta, 256, 255)
                      str_descricao_status_resposta = Left(str_descricao_status_resposta, 255)
                    End If
                  End If
                End If
                
                If str_split_linha(4) = "Cancelamento Homologado" Then
                  str_data_hora_processamento = str_split_linha(7)
                  str_numero_protocolo_autorizacao = str_split_linha(8)
                  str_digest_value = ""
                  str_codigo_status_resposta = "101"
                  str_descricao_status_resposta = str_split_linha(4)
                  str_Retorno = str_split_linha(4)
                End If
                
                If str_descricao_status_resposta = "110111" Then
                  str_descricao_status_resposta = "101"
                  str_numero_protocolo_autorizacao = vetChave(10)
                End If
                
    '            'Ex.:1100;1;v3314;41090882268160000180010010000000083226895707;2009-08-04T17:47:47;141090016866038;RR1ixtu3pBoVF8D7KfgnO+Bohmc=;100;Autorizado o uso de NF-e (IdNFe: NFe41090800069957000194550010000362840000000015);
    '            '    0    5 7    13                                           58                  78             94                           123 127
    '            str_data_hora_processamento = Mid(str_linhaTexto, 59, 19)
    '            str_numero_protocolo_autorizacao = Mid(str_linhaTexto, 79, 15)
    '            str_digest_value = Mid(str_linhaTexto, 95, 28)
    '            str_codigo_status_resposta = Mid(str_linhaTexto, 124, 3)
    '            int_tamanho_que_falta = (Len(str_linhaTexto) - 128)
    '            str_descricao_status_resposta = Mid(str_linhaTexto, 128, int_tamanho_que_falta)
    '            str_descricao_status_resposta = RetirarParteFinal(str_descricao_status_resposta)
        
                strSQL = "SELECT * "
                strSQL = strSQL & "FROM NFeRetorno "
                strSQL = strSQL & "WHERE Filial = " & gnCodFilial & " AND Sequencia = " & rsNFe.Fields("Sequencia").Value & " AND TipoMovimento = " & m_intTipoMovimentoPesquisa
                Set rsNFeRetorno = db.OpenRecordset(strSQL, dbOpenDynaset)
                With rsNFeRetorno
                  If (.BOF And .EOF) Then
                    .AddNew
                    .Fields("Filial").Value = gnCodFilial
                    .Fields("Sequencia").Value = rsNFe.Fields("Sequencia").Value
                    .Fields("TipoMovimento").Value = m_intTipoMovimentoPesquisa
                    .Fields("DataHora").Value = Now
                    .Fields("Protocolo").Value = str_numero_protocolo_autorizacao
                    .Fields("DigestValue").Value = str_digest_value
                    .Fields("Status").Value = str_codigo_status_resposta
                    .Fields("StatusDescricao").Value = str_descricao_status_resposta
                    If Len(str_descricao_status2_resposta) > 0 Then
                      .Fields("StatusDescricao2").Value = str_descricao_status2_resposta
                    End If
                    .Update
                    
                    '05/03/2013-Alexandre Afornali
                    'Incluido tratamento de retorno para atualizar a tabela Nfe_Retorno quando codigo de
                    'erro igual a 100(Autorizado o uso da Nfe) ou 101(Cancelamento Homologado)
                  ElseIf (str_codigo_status_resposta = "100" Or str_codigo_status_resposta = "101") Then
                    .Edit
                    .Fields("DataHora").Value = Now
                    .Fields("Protocolo").Value = str_numero_protocolo_autorizacao
                    .Fields("DigestValue").Value = str_digest_value
                    .Fields("Status").Value = str_codigo_status_resposta
                    .Fields("StatusDescricao").Value = str_descricao_status_resposta
                    .Update
                  End If
                  .Close
                End With
                Set rsNFeRetorno = Nothing
        
                MsgBox "Gravado Retorno NFe Número : " & rsNFe.Fields("Numero").Value & " Retorno da Receita: " & str_Retorno, vbExclamation, "Atenção/Retorno da Receita"
                
                str_dia = Mid(str_data_hora_processamento, 9, 2)
                str_mes = Mid(str_data_hora_processamento, 6, 2)
                str_ano = Mid(str_data_hora_processamento, 1, 4)
                Str_Data = str_dia & "/" & str_mes & "/" & str_ano
        
                Select Case str_codigo_status_resposta
                
                  Case "100" 'Autorizado o uso da NF-e - Grava o protocolo de autorizacao na tbl NFe
                    rsNFe.Edit
                    rsNFe.Fields("ProtocoloAutorizacao").Value = str_numero_protocolo_autorizacao
                    rsNFe.Fields("DataHoraAutorizacao").Value = CDate(Str_Data)
                    rsNFe.Update
                    rsNFe.Close
                  Case "101" 'Cancelamento de NFe homologada. Grava o protocolo de cancelamento na tbl NFe
                    rsNFe.Edit
                    rsNFe.Fields("ProtocoloCancelamento").Value = str_numero_protocolo_autorizacao
                    rsNFe.Fields("DataHoraCancelamento").Value = CDate(Str_Data)
                    rsNFe.Update
                    rsNFe.Close
                 
                  Case "102" 'Inutilizacao de número homologado
                    'A3.DBFundamentals.Connection.DBConnection.SetFieldByGUID("tbl291_fiscal_NFe", "T291_protocolo_de_cancelamento", "T291_PK_GUID", obj_fiscal_NFe.GUID, str_numero_protocolo_autorizacao)
                  
                  Case Else
        
                End Select
              '14/09/2010 - Andrea
              Else 'Deu erro na emissao
              
                  'Gravar na tbl de retonro o erro
                  'Ex.:1100;2;InfinitriNFe;41091001820871000114550000000000010930553685;2009-10-01T17:22:29;0;;225;&lt;ValidationErrors&gt;&lt;ErrorMessage&gt;&lt;ErrorDetail&gt;The 'http://www.portalfiscal.inf.br/nfe:xNome' element is invalid - The value '3M DO BRASIL ' is invalid according to its datatype 'String' - The Pattern constraint failed.&lt;/ErrorDetail&gt;&lt;/ErrorMessage&gt;&lt;ErrorMessage&gt;&lt;ErrorDetail&gt;The 'http://www.portalfiscal.inf.br/nfe:nro' element is invalid - The value '' is invalid according to its datatype 'String' - The Pattern constraint failed.&lt;/ErrorDetail&gt;&lt;/ErrorMessage&gt;&lt;ErrorMessage&gt;&lt;ErrorDetail&gt;The 'http://www.portalfiscal.inf.br/nfe:xNome' element is invalid - The value 'SEDEX ' is invalid according to its datatype 'String' - The Pattern constraint failed.&lt;/ErrorDetail&gt;&lt;/ErrorMessage&gt;&lt;/ValidationErrors&gt;;
                  '    0    5 7            20     120:
    
                str_split_linha = Split(str_linhaTexto, ";")
                
                'Verifica retorno
                If UBound(str_split_linha) > 7 Then
                  str_chave_acesso = str_split_linha(3)
                  str_data_hora_processamento = str_split_linha(4)
                  str_numero_protocolo_autorizacao = str_split_linha(5)
                  str_digest_value = str_split_linha(6)
                  str_codigo_status_resposta = str_split_linha(7)
                  str_descricao_status_resposta = str_split_linha(8)
                  str_Retorno = str_split_linha(8)
                  If UBound(str_split_linha) > 14 Then
                    str_descricao_status_resposta = str_split_linha(14)
                    str_Retorno = str_split_linha(14)
                  End If
                  If Len(str_descricao_status_resposta) > 255 Then
                    str_descricao_status2_resposta = Mid(str_descricao_status_resposta, 256, 255)
                    str_descricao_status_resposta = Left(str_descricao_status_resposta, 255)
                  End If
                Else
                  MsgBox "Arquivo de Retorno de NFe corrompido, favor verificar arquivo : " & ARQUIVO.Name, vbExclamation, "Atenção"
                  '06/03/2013-Alexandre Afornali
                  'Alterado para ir para proximo arquivo ao inves de sair da função
                  GoTo ProximoArquivo
                End If
    
                str_dia = Mid(str_data_hora_processamento, 9, 2)
                str_mes = Mid(str_data_hora_processamento, 6, 2)
                str_ano = Mid(str_data_hora_processamento, 1, 4)
                Str_Data = str_dia & "/" & str_mes & "/" & str_ano
                
                'NFe
                strSQL = "SELECT ChaveAcesso, ProtocoloAutorizacao, Numero, Sequencia, DataHoraAutorizacao, DataHoraCancelamento  "
                strSQL = strSQL & "FROM NFe "
                strSQL = strSQL & "WHERE Filial = " & gnCodFilial & " AND ChaveAcesso = '" & str_chave_acesso & "' AND TipoMovimento = " & m_intTipoMovimentoPesquisa
                Set rsNFe = db.OpenRecordset(strSQL, dbOpenDynaset)
                With rsNFe
                  If (.BOF And .EOF) Then
                    MsgBox "NFe Não encontrada no arquivo de envio, chave de acesso número : " & str_chave_acesso, vbExclamation, "Atenção"
                    .Close
                    Set rsNFe = Nothing
                    '06/03/2013-Alexandre Afornali
                    'Alterado para ir para proximo arquivo ao inves de sair da função
                    GoTo ProximoArquivo
                  
                  Else
                  
                    'Grava o retorno
                    strSQL = "SELECT * "
                    strSQL = strSQL & "FROM NFeRetorno "
                    strSQL = strSQL & "WHERE Filial = " & gnCodFilial & " AND Sequencia = " & rsNFe.Fields("Sequencia").Value & " AND TipoMovimento = " & m_intTipoMovimentoPesquisa
                    Set rsNFeRetorno = db.OpenRecordset(strSQL, dbOpenDynaset)
                    With rsNFeRetorno
                      If (.BOF And .EOF) Then
                        .AddNew
                        .Fields("Filial").Value = gnCodFilial
                        .Fields("Sequencia").Value = rsNFe.Fields("Sequencia").Value
                        .Fields("TipoMovimento").Value = m_intTipoMovimentoPesquisa
                        .Fields("Protocolo").Value = ""
                        .Fields("DigestValue").Value = ""
                        .Fields("Status").Value = str_codigo_status_resposta
                        .Fields("StatusDescricao").Value = str_descricao_status_resposta
                        If Len(str_descricao_status2_resposta) > 0 Then
                          .Fields("StatusDescricao2").Value = str_descricao_status2_resposta
                        End If
                        .Fields("DataHora").Value = Now
                        .Update
                        
                        '05/03/2013-Alexandre Afornali
                        'Alterado tratamento ao importar um segundo retorno de erro, a descrição do erro na tabela
                        'Nfe_Retorno não era atualizada
                      ElseIf (.Fields("Status").Value <> 100 And .Fields("StatusDescricao").Value <> 101 And .Fields("StatusDescricao").Value <> 102) Then
                        .Edit
                        .Fields("StatusDescricao").Value = str_descricao_status_resposta
                        .Fields("Status").Value = str_codigo_status_resposta
                        If Len(str_descricao_status2_resposta) > 0 Then
                          .Fields("StatusDescricao2").Value = str_descricao_status2_resposta
                        End If
                        .Fields("DataHora").Value = Now
                        .Update
                      End If
                      .Close
                    End With
                    Set rsNFeRetorno = Nothing
                    MsgBox "Gravado Retorno NFe Número : " & rsNFe.Fields("Numero").Value & " Retorno da Receita: " & str_Retorno, vbExclamation, "Atenção/Retorno da Receita"
                  
                  End If
                End With
    
              End If
    
            End If
          
          Wend
DeletarArquivo:
          Instream.Close
        
          FSys.DeleteFile ARQUIVO, True
        'End If
        
ProximoArquivo:
        
      Next
  End If

  If bTemAlgumRegistroSelecionadoNaGrid = False Then
      MsgBox "Selecione uma NFe na grade.", vbInformation, "Informação"
  End If


  Exit Sub

ErrHandler:
  MsgBox "Erro ao processar Retorno : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"

End Sub

Private Function RetirarParteFinal(ByVal str_descricao_status_resposta As String) As String

    Dim int_x As Integer
    Dim str_part As String
    Dim str_return As String

    For int_x = 1 To Len(str_descricao_status_resposta)
      str_part = Mid(str_descricao_status_resposta, int_x, 1)
      If str_part <> ";" Then
        str_return = str_return & str_part
      Else
        Exit For
      End If
    Next

    RetirarParteFinal = str_return

End Function


Private Sub EnviarNFe()
  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim objNFe As New clsNFe
  Dim strSQL As String
  Dim bSelecionouRegistroNaGrade As Boolean
    
  '17/11/2010 - Andrea
  Dim rsParametros As Recordset
  Dim strVersaoLayoutEnvio As String
  Dim strPadraoArquivoIntegracao As String
  
  bSelecionouRegistroNaGrade = False
  
  On Error GoTo ErrHandler

  'Parâmetros da Filial
  strSQL = "SELECT CGC, [Razão Social], Nome, Inscrição, Endereço, EnderecoNumero, EnderecoComplemento, "
  strSQL = strSQL & "Bairro, Cidade, Estado, CEP, Pais, Fone, "
  strSQL = strSQL & "AmbienteNfe, FormatoImpressaoDanfeNfe, ModDetBaseCalculoIcms, ModDetBaseCalculoIcmsSt, "
  '16/11/2011 - Andrea
  'Incluído campo padrão de arquivo de integracao (txt ou xml)
  strSQL = strSQL & "PastaEnvioNfe, PastaRetornoNfe, VersaoLayoutEnvio, PadraoArquivoIntegracao "
  strSQL = strSQL & "FROM [Parâmetros Filial] "
  strSQL = strSQL & "WHERE Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsParametros
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
      .Close
      Set rsParametros = Nothing
      Exit Sub
    Else
      'Validações
      strVersaoLayoutEnvio = Trim(.Fields("VersaoLayoutEnvio").Value)
      If strVersaoLayoutEnvio = "" Then
        '02/12/2011 - Andrea
        strVersaoLayoutEnvio = "4.00"
      End If
      strPadraoArquivoIntegracao = Trim(.Fields("PadraoArquivoIntegracao").Value)
    End If
  End With
  
  'strVersaoLayoutEnvio = "3.10"


  '29/11/2010 - Andrea
  'Se for emissao em Contingência, valida os campos data, hora e motivo da entrada em contingência (que não podem estar vazios
  If m_intTipoEmissao <> 1 Then
  
    If Trim(txtHoraEntradaContingencia.Text) = "" Then
      MsgBox "Hora de entrada em regime de contingência inválida, favor preencher corretamente.", vbExclamation, "Atenção"
      Exit Sub
    End If
    
    If Trim(txtMotivoContingencia.Text) = "" Then
      MsgBox "Motivo da entrada em regime de contingência inválido, favor preencher corretamente.", vbExclamation, "Atenção"
      Exit Sub
    End If
      
  
  End If
    
  '02/03/2011 - Andrea
  'Se o tipo de emissão (da tela) for = 4 , muda para 5 (Contingencia com Formulário de Segurança - Documento Adicional)
  If m_intTipoEmissao = 4 Then
    m_intTipoEmissao = 5
  End If
  
  '01/10/2010 - Andrea
  Dim blnConcatenarDescricaoAdicional As Boolean
  
  If chkConcatenarDescricaoAdicional.Value = vbChecked Then
    blnConcatenarDescricaoAdicional = True
  Else
    blnConcatenarDescricaoAdicional = False
  End If
  
  Dim iRetMsgBox As Integer
  
  With grdMovimento
    'Screen.MousePointer = vbHourglass

    If strVersaoLayoutEnvio = "4.00" Then
      For nRow = 0 To .Rows - 1
        bm = .AddItemBookmark(nRow)
        
        If CBool(.Columns("Enviar").CellValue(bm)) Then
          lngSequencia = CLng(gsHandleNull(.Columns("Sequencia").CellValue(bm)))
          bSelecionouRegistroNaGrade = True
          
          If .Columns("Status").CellValue(bm) = "Cancelada" Then
              MsgBox "NFe Nº" & .Columns("Nota Fiscal").CellValue(bm) & " já esta com status Cancelada."

'              iRetMsgBox = MsgBox("NFe Nº" & .Columns("Nota Fiscal").CellValue(bm) & " já esta com status CANCELADA. Deseja mesmo assim realizar um novo envio a SEFAZ?", vbOKCancel)
'
'              If iRetMsgBox = 1 Then
'                  'Enviar mesmo assim...
'                  If Not objNFe.Enviar_4_0(gnCodFilial, lngSequencia, m_intTipoMovimentoPesquisa, blnConcatenarDescricaoAdicional, m_intTipoEmissao, dteEntradaContingencia.Text, txtHoraEntradaContingencia.Text, txtMotivoContingencia.Text, strPadraoArquivoIntegracao) Then
'                      Exit For
'                  End If
'              End If
          ElseIf .Columns("Status").CellValue(bm) = "Autorizada" Then
              MsgBox "NFe Nº" & .Columns("Nota Fiscal").CellValue(bm) & " já esta com status Autorizada."
              'Exit For
          Else
              If Not objNFe.Enviar_4_0(gnCodFilial, lngSequencia, m_intTipoMovimentoPesquisa, blnConcatenarDescricaoAdicional, m_intTipoEmissao, dteEntradaContingencia.Text, txtHoraEntradaContingencia.Text, txtMotivoContingencia.Text, strPadraoArquivoIntegracao) Then
                  Exit For
              End If
          End If
        End If
      Next nRow
    Else
      For nRow = 0 To .Rows - 1
        bm = .AddItemBookmark(nRow)
        
        If CBool(.Columns("Enviar").CellValue(bm)) Then
          lngSequencia = CLng(gsHandleNull(.Columns("Sequencia").CellValue(bm)))
          
         
          If Not objNFe.Enviar_4_0(gnCodFilial, lngSequencia, m_intTipoMovimentoPesquisa, blnConcatenarDescricaoAdicional, m_intTipoEmissao, dteEntradaContingencia.Text, txtHoraEntradaContingencia.Text, txtMotivoContingencia.Text, strPadraoArquivoIntegracao) Then
            Exit For
          End If
          
        End If
      Next nRow
    End If

    'Screen.MousePointer = vbDefault
  End With
  
  Set objNFe = Nothing
  
  If bSelecionouRegistroNaGrade = False Then
    MsgBox "Selecione uma Nota na grade", vbInformation, "Atenção"
  End If
  
  Exit Sub

ErrHandler:
  'If Screen.MousePointer = vbHourglass Then
  '  Screen.MousePointer = vbDefault
  'End If

  MsgBox "Erro ao enviar Nota Fiscal: " & Err.Number & " (" & Err.Description & "). Sequência: " & lngSequencia & ".", vbCritical, "Erro"
  
End Sub


Private Sub CartaCorrecaoNFe_XML()
  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim pNumero As Long
  Dim pSerie As Integer
  Dim pTipoDoc As String
  Dim pCNPJEmitente As String
  Dim strSQL As String
  
  Dim objNFe As New clsNFe
  
  On Error GoTo ErrHandler
     
  'Parâmetros da Filial
  Dim rsParametros As Recordset
  strSQL = "SELECT CGC FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  If (rsParametros.BOF And rsParametros.EOF) Then
      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
      Exit Sub
  Else
      pCNPJEmitente = rsParametros.Fields("CGC").Value
      pCNPJEmitente = Replace(pCNPJEmitente, "/", "")
      pCNPJEmitente = Replace(pCNPJEmitente, "-", "")
      pCNPJEmitente = Replace(pCNPJEmitente, ".", "")
      pCNPJEmitente = Replace(pCNPJEmitente, ",", "")
  End If
  rsParametros.Close
  Set rsParametros = Nothing
  
  objNFe.CartaCorrecaoXML_SEFAZ gnCodFilial, CLng(txt_numNFeCC.Text), txt_serieCC.Text, pCNPJEmitente, txt_descCC.Text, txt_tipoDocCC.Text
  
  Set objNFe = Nothing
  
  Exit Sub

ErrHandler:
  MsgBox "Erro ao realizar Carta Correção da Nota Fiscal : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"
  
End Sub


Private Sub InutilizarNFe_XML()
  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim pNumero As Long
  Dim pSerie As Integer
  Dim pTipoDoc As String
  Dim pCNPJEmitente As String
  Dim strSQL As String
  
  Dim objNFe As New clsNFe
  
  On Error GoTo ErrHandler
  
  If Len(txtMotivoInutilizacao.Text) = 0 Then
    MsgBox "Informe o Motivo de Inutilização (mínimo 15 caracteres).", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txtMotivoInutilizacao.Text) < 15 Then
    MsgBox "Motivo do Inutilização deve ter no mínimo 15 caracteres.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txtSerie.Text) = 0 Then
    MsgBox "Informe a Série da NFe/NFCe a ser inutilizada.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txtNumeroNFeInicial.Text) = 0 Then
    MsgBox "Informe o Número Inicial a ser inutilizado.", vbExclamation, "Atenção"
    Exit Sub
  End If
  If Len(txtNumeroNFeFinal.Text) = 0 Then
    MsgBox "Informe o Número Final a ser inutilizado.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txt_anoInutilizacao.Text) = 0 Then
    MsgBox "Informe o Ano. Ex: 2018", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  'Parâmetros da Filial
  Dim rsParametros As Recordset
  strSQL = "SELECT CGC FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  If (rsParametros.BOF And rsParametros.EOF) Then
      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
      Exit Sub
  Else
      pCNPJEmitente = rsParametros.Fields("CGC").Value
      pCNPJEmitente = Replace(pCNPJEmitente, "/", "")
      pCNPJEmitente = Replace(pCNPJEmitente, "-", "")
      pCNPJEmitente = Replace(pCNPJEmitente, ".", "")
      pCNPJEmitente = Replace(pCNPJEmitente, ",", "")
  End If
  rsParametros.Close
  Set rsParametros = Nothing
  
  If opt_inutilizadasNFe.Value = True Then
      objNFe.InutilizarXML_SEFAZ gnCodFilial, CLng(txtNumeroNFeInicial.Text), CLng(txtNumeroNFeFinal.Text), txtSerie.Text, CInt(txt_anoInutilizacao.Text), pCNPJEmitente, txtMotivoInutilizacao.Text, "55"
  Else
      objNFe.InutilizarXML_SEFAZ gnCodFilial, CLng(txtNumeroNFeInicial.Text), CLng(txtNumeroNFeFinal.Text), txtSerie.Text, CInt(txt_anoInutilizacao.Text), pCNPJEmitente, txtMotivoInutilizacao.Text, "65"
  End If
  
  Set objNFe = Nothing
  
  Exit Sub

ErrHandler:
  MsgBox "Erro ao inutilizar Nota Fiscal : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"
  
End Sub

Private Sub CancelarNFe_XML()
  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim pNumero As Long
  Dim pSerie As Integer
  Dim pTipoDoc As String
  Dim pCNPJEmitente As String
  Dim strSQL As String
  
  Dim objNFe As New clsNFe
  
  On Error GoTo ErrHandler
  
  If Len(txtMotivoCancelamento.Text) = 0 Then
    MsgBox "Informe o Motivo do Cancelamento da NFe.", vbExclamation, "Atenção"
    Exit Sub
  End If
      
  If Len(txtMotivoCancelamento.Text) < 15 Then
    MsgBox "Motivo do Cancelamento deve ter no mínimo 15 caracteres.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  'Parâmetros da Filial
  Dim rsParametros As Recordset
  strSQL = "SELECT CGC FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  If (rsParametros.BOF And rsParametros.EOF) Then
      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
      Exit Sub
  Else
      pCNPJEmitente = rsParametros.Fields("CGC").Value
      pCNPJEmitente = Replace(pCNPJEmitente, "/", "")
      pCNPJEmitente = Replace(pCNPJEmitente, "-", "")
      pCNPJEmitente = Replace(pCNPJEmitente, ".", "")
      pCNPJEmitente = Replace(pCNPJEmitente, ",", "")
  End If
  rsParametros.Close
  Set rsParametros = Nothing
  
  With grdMovimento
    For nRow = 0 To .Rows - 1
      bm = .AddItemBookmark(nRow)
      
      If CBool(.Columns("Enviar").CellValue(bm)) Then
        lngSequencia = CLng(gsHandleNull(.Columns("Sequencia").CellValue(bm)))
        pNumero = CLng(gsHandleNull(.Columns("Nota Fiscal").CellValue(bm)))
        pSerie = CInt(gsHandleNull(.Columns("Serie").CellValue(bm)))
        
        If m_intTipoMovimentoPesquisa = 1 Then
          pTipoDoc = "S"
        Else
          pTipoDoc = "N"
        End If
        objNFe.CancelarXML_SEFAZ lngSequencia, pNumero, pSerie, "55", gnCodFilial, pCNPJEmitente, txtMotivoCancelamento.Text, pTipoDoc
      
      End If
    Next nRow
  End With
  
  Set objNFe = Nothing
  
  Exit Sub

ErrHandler:
  MsgBox "Erro ao cancelar Nota Fiscal : " & Err.Number & " (" & Err.Description & "). Sequência: " & lngSequencia & ".", vbCritical, "Erro"
  
End Sub

'02/10/2009 - Andrea
Private Sub CancelarNFe()
  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim objNFe As New clsNFe
  
  On Error GoTo ErrHandler
  
  If Len(txtMotivoCancelamento.Text) = 0 Then
    MsgBox "Antes de Cancelar uma NFe é necessário informar o Motivo do Cancelamento.", vbExclamation, "Atenção"
    Exit Sub
  End If
      
  If Len(txtMotivoCancelamento.Text) < 15 Then
    MsgBox "Campo Motivo do Cancelamento precisa ter no mínimo 15 caracteres, favor verificar.", vbExclamation, "Atenção"
    Exit Sub
  End If

  With grdMovimento
    For nRow = 0 To .Rows - 1
      bm = .AddItemBookmark(nRow)
      
      If CBool(.Columns("Enviar").CellValue(bm)) Then
        lngSequencia = CLng(gsHandleNull(.Columns("Sequencia").CellValue(bm)))
        
        If Not objNFe.Cancelar(gnCodFilial, lngSequencia, m_intTipoMovimentoPesquisa, txtMotivoCancelamento.Text) Then
          Exit For
        End If
      End If
    Next nRow
  End With
  
  Set objNFe = Nothing
  
  Exit Sub

ErrHandler:
  MsgBox "Erro ao cancelar Nota Fiscal : " & Err.Number & " (" & Err.Description & "). Sequência: " & lngSequencia & ".", vbCritical, "Erro"
  
End Sub


Private Sub InutilizarNFe()
  
  On Error GoTo ErrHandler
  
  If Len(txtMotivoInutilizacao.Text) = 0 Then
    MsgBox "Antes de Inutilizar uma NFe é necessário informar o Motivo da Inutilização.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txtMotivoInutilizacao.Text) < 15 Then
    MsgBox "Campo Motivo da Inutilização precisa ter no mínimo 15 caracteres, favor verificar.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txtSerie.Text) = 0 Then
    MsgBox "Antes de Inutilizar uma NFe é necessário informar a Série da NFe a ser inutilizada.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txtNumeroNFeInicial.Text) = 0 Then
    MsgBox "Antes de Inutilizar uma NFe é necessário informar o Número Inicial a ser inutilizado.", vbExclamation, "Atenção"
    Exit Sub
  End If
  If Len(txtNumeroNFeFinal.Text) = 0 Then
    MsgBox "Antes de Inutilizar uma NFe é necessário informar o Número Final a ser inutilizado.", vbExclamation, "Atenção"
    Exit Sub
  End If

  Dim objNFe As New clsNFe
      
  If objNFe.Inutilizar(gnCodFilial, m_intTipoMovimentoPesquisa, txtMotivoInutilizacao.Text, txtSerie.Text, txtNumeroNFeInicial.Text, txtNumeroNFeFinal.Text) Then
    MsgBox "Inutilização terminada com sucesso.", vbExclamation, "Atenção"
  End If
 
  Set objNFe = Nothing
  
  Exit Sub



ErrHandler:
  MsgBox "Erro ao inutilizar NFe: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub CarregarOperacoes(sTipoOperacao As String)
    Dim rsOperacoes As Recordset
    Dim sCodigo As String
  
    cbo_operacao.Clear
    
    If sTipoOperacao = "SAIDA" Then
        Set rsOperacoes = db.OpenRecordset("Select * from [Operações Saída] order by Código", dbOpenDynaset)
    Else
        Set rsOperacoes = db.OpenRecordset("Select * from [Operações Entrada] order by Código", dbOpenDynaset)
    End If
    
    cbo_operacao.AddItem "TODAS"
    
    While Not rsOperacoes.EOF
        sCodigo = CStr(rsOperacoes.Fields("Código").Value)
        
        If Len(Trim(sCodigo)) = 1 Then
            sCodigo = "  " & sCodigo
        ElseIf Len(Trim(sCodigo)) = 2 Then
            sCodigo = " " & sCodigo
        End If
        
        cbo_operacao.AddItem sCodigo & " - " & rsOperacoes.Fields("Nome").Value
        rsOperacoes.MoveNext
    Wend

    rsOperacoes.Close
    Set rsOperacoes = Nothing
End Sub


Private Sub cboTipoMovimento_LostFocus()
  If cboTipoMovimento.ListIndex = 0 Then
      ' Entradas
      CarregarOperacoes "ENTRADA"
      
  ElseIf cboTipoMovimento.ListIndex = 1 Then
      ' Saídas
      CarregarOperacoes "SAIDA"
      
  End If
End Sub

Private Sub chk_nfeDevolucao_impostoDevol_Click()
  If chk_nfeDevolucao_impostoDevol.Value = 1 Then
    nfeDevolucao_impostoDevol = True
  Else
    nfeDevolucao_impostoDevol = False
  End If
End Sub

Private Sub chk_nfeInfAdProd_Click()
  If chk_nfeInfAdProd.Value = 1 Then
    nfeInfAdProd = True
  Else
    nfeInfAdProd = False
  End If
End Sub

Private Sub chk_xPed_nItemPed_Click()
  If chk_xPed_nItemPed.Value = 1 Then
    nfe_xPed_nItemPed = True
    txt_sequencia.Enabled = True
    cmd_mostrarProdutos.Enabled = True
    grid_produtos.Enabled = True
    cmd_salvarProdutos.Enabled = True
  Else
    nfe_xPed_nItemPed = False
  End If

End Sub

Private Sub cmd_ajusteStatusNoQuick_Click()
On Error GoTo Erro
  Dim sMsg As String
  Dim nStyle As String
  Dim nResponse As String
  Dim nRow As Integer
  Dim bm As Variant
  Dim strSQL As String
  Dim lngSequencia As Long
  Dim statusNaTela As String

  With grdMovimento
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)
        
          If CBool(.Columns("Enviar").CellValue(bm)) Then
              lngSequencia = CLng(gsHandleNull(.Columns("Sequencia").CellValue(bm)))
              statusNaTela = .Columns("Status").CellValue(bm)

              If statusNaTela = "Autorizada" Or statusNaTela = "Cancelada" Then
                  MsgBox "NFe Nº" & .Columns("Nota Fiscal").CellValue(bm) & " já esta com status AUTORIZADA ou CANCELADA. Então vc NÃO pode fazer ajuste de status manual."
                  Exit Sub
              End If

              sMsg = "Deseja realmente atualizar o status para AUTORIZADA da NFe Nº" & .Columns("Nota Fiscal").CellValue(bm) & " ?"
              nStyle = vbYesNo + vbQuestion + vbDefaultButton2
              nResponse = MsgBox(sMsg, nStyle, "Atenção AJUSTE DE STATUS APENAS NO QUICK (NÃO NA FAZENDA)")
              If nResponse = vbNo Then
                Exit Sub
              End If
              
              strSQL = "UPDATE NFe SET Status = 100, ProtocoloAutorizacao = '100_AUTO_MANUAL'" & _
                " WHERE Filial = " & gnCodFilial & " AND Sequencia = " & lngSequencia
              
              db.Execute strSQL, dbFailOnError
              
              MsgBox "NFe Nº" & .Columns("Nota Fiscal").CellValue(bm) & " atualizada para AUTORIZADA MANUAL."
              cmdPesquisar_Click
              Exit Sub
          End If
      Next
  End With
  
  Exit Sub
Erro:
  MsgBox "Erro ao tentar fazar ajuste de status manual " & Err.Number & " - " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_atalho_Click()
        Dim objTela As Form
        
        Set objTela = New frmXML
        
        frmXML.sXML = "Cole aqui o seu XML"
        frmXML.sXML_Erro = ""
        frmXML.xNomeArquivoXML = ""
        frmXML.sSequencia = ""
        frmXML.iOrigemChamador = 1 ' 1-Tela frmNFe Aba Erros/Críticas     2-Tela frmNFe Aba Notas Fiscais
        frmXML.Show 1
End Sub

Private Sub cmd_calendarioNFCe_Click()
    msk_dataDia.Text = frmCalendario.gsDateCalender(msk_dataDia.Text)
End Sub

Private Sub cmd_calendarioNFCeAbaCinza_Click()
    msk_dataDiaNFCe.Text = frmCalendario.gsDateCalender(msk_dataDiaNFCe.Text)
End Sub

Private Sub cmd_cancelarNFe_Click()
On Error GoTo Erro:

  Dim retMsg As Variant
  retMsg = MsgBox("Deseja realmente fazer o cancelamento da NF-e?", vbYesNo, "Cancelamento de NF-e")
  
  If retMsg = vbNo Then
      Exit Sub
  End If

  If Len(LTrim(RTrim(txtMotivoCancelamento.Text))) < 15 Then
      MsgBox "Mínimo 15 caracteres.", vbInformation, "Atenção"
      txtMotivoCancelamento.SetFocus
      Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  
  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
  
  If sPadraoArquivoIntegracao = "XML" Then
      CancelarNFe_XML
  Else
    ' TXT
    Call CancelarNFe
  End If
  
  
  cmdPesquisar_Click
  
  Screen.MousePointer = vbDefault
  
  MsgBox "NFe cancelada com sucesso!", vbInformation, "Sucesso"
  
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If
  
  MsgBox "Erro na chamada ao Método de Cancelamento de NFe. Erro:" & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub cmd_dataRetroativa_Click()
    msk_dataRetroativa.Text = frmCalendario.gsDateCalender(msk_dataRetroativa.Text)
End Sub

Private Sub cmd_enviarCC_Click()
On Error GoTo Erro:
  Screen.MousePointer = vbHourglass
  
  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
  
  If Len(txt_descCC.Text) = 0 Then
    MsgBox "Informe a Descrição da Carta Correção (mínimo 15 caracteres).", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txt_descCC.Text) < 15 Then
    MsgBox "Descrição da Carta Correção deve ter no mínimo 15 caracteres.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txt_serieCC.Text) = 0 Then
    MsgBox "Informe a Série da NFe.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txt_numNFeCC.Text) = 0 Then
    MsgBox "Informe o Número da NFe.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  If Len(txt_tipoDocCC.Text) = 0 Then
    MsgBox "Informe o tipo de documento (S para Saída  ou  E para Entrada)."
    Exit Sub
  End If
  
  If sPadraoArquivoIntegracao = "XML" Then
      CartaCorrecaoNFe_XML
      
      cmd_pesquisarCC_Click
  End If
  
  Screen.MousePointer = vbDefault
  
  MsgBox "NFe Carta Correção realizada com sucesso!", vbInformation, "Sucesso"
  
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If
  
  MsgBox "Erro na chamada ao Método de Carta Correção de NFe. Erro:" & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_enviarNFe_Click()
On Error GoTo Erro

  txt_posicaoNFeErro.Text = ""
  gstr_posicaoNFeErro_telaAcelerador = ""

  'lbl_statusProcessamento.Caption = "Transmitindo NFe para a SEFAZ..."
  lbl_statusProcessamento.Visible = True
  picture_statusProcessamento.Top = 7950

  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
  
  If chk_enviaComDataRetroativa.Value = vbChecked And msk_dataRetroativa.Text = "  /  /    " Then
      MsgBox "Você marcou para envio com data retroativa. Então informe a data menor que hoje", vbInformation, "Atenção"
      Exit Sub
  End If

  If chk_enviaComDataRetroativa.Value = vbChecked Then
      bolEnviaDataRetroativa = True
      sEnviaDataRetroativa = msk_dataRetroativa.Text
  Else
      bolEnviaDataRetroativa = False
      sEnviaDataRetroativa = ""
  End If

  EnviarNFe
  'Call EnviarNFe
  
  'Limpa retroativa: variáveis e controles
  chk_enviaComDataRetroativa.Value = False
  bolEnviaDataRetroativa = False
  sEnviaDataRetroativa = ""
  msk_dataRetroativa.Text = "  /  /    "
  
  
  cmdPesquisar_Click

  'Screen.MousePointer = vbDefault
  
  'If Screen.MousePointer = vbHourglass Then
  '  Screen.MousePointer = vbDefault
  'End If
  
  'lbl_statusProcessamento.Caption = ""
  lbl_statusProcessamento.Visible = False
  picture_statusProcessamento.Top = 8950
  
  If Trim(gstr_posicaoNFeErro_telaAcelerador) <> "" Then
      txt_posicaoNFeErro.Text = gstr_posicaoNFeErro_telaAcelerador
  End If
  
  Exit Sub
Erro:
  'If Screen.MousePointer = vbHourglass Then
  ' Screen.MousePointer = vbDefault
  'End If
  
  'Limpa retroativa: variáveis e controles
  chk_enviaComDataRetroativa.Value = False
  bolEnviaDataRetroativa = False
  sEnviaDataRetroativa = ""
  msk_dataRetroativa.Text = "  /  /    "
  
  
  'lbl_statusProcessamento.Caption = ""
  lbl_statusProcessamento.Visible = False
  picture_statusProcessamento.Top = 8950

  MsgBox "Erro na chamada ao Método de Envio de NFe. Erro:" & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmd_helpRejeicao_Click()
On Error GoTo Erro:

  If gAbreModuloXML = 0 Then
    'Caso não tenha acesso a este modulo...direcionar para tela de aquisição
    frmAquisicaoEstrategicoRel.Show
    Exit Sub
  End If
  
  Dim sNumNFe As String
  Dim sCodigoRejeicao As String
  Dim sDescricaoRejeicao As String
  Dim iIndice As Long
  
  If gridRetorno.RowSel > 0 Then
      sNumNFe = gridRetorno.TextMatrix(gridRetorno.RowSel, 1)
      sCodigoRejeicao = gridRetorno.TextMatrix(gridRetorno.RowSel, 5)
      sDescricaoRejeicao = gridRetorno.TextMatrix(gridRetorno.RowSel, 6)
      
      If sNumNFe = "" Then
        MsgBox "Selecione um registro na grade!", vbInformation, "NFe Visualizar XML"
        Exit Sub
      End If
      
      sCodigoRejeicao = Replace(sCodigoRejeicao, "-", "")
      
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      
      If sCodigoRejeicao = 1003 Then
          iIndice = InStr(1, sDescricaoRejeicao, "NCM")
          If iIndice > 0 Then
              Call objHelp.Show(strfile, "QuickStore10Help", 10010)
          Else
              iIndice = InStr(1, sDescricaoRejeicao, "cMun")
              If iIndice > 0 Then
                  Call objHelp.Show(strfile, "QuickStore10Help", 10017)
              Else
                  iIndice = InStr(1, sDescricaoRejeicao, "nro")
                  If iIndice > 0 Then
                      Call objHelp.Show(strfile, "QuickStore10Help", 10018)
                  Else
                      iIndice = InStr(1, sDescricaoRejeicao, "refNFe")
                      If iIndice > 0 Then
                          Call objHelp.Show(strfile, "QuickStore10Help", 10033)
                      Else
                      
                      End If
                  End If
              End If
          End If
      ElseIf sCodigoRejeicao = 210 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10025)
      ElseIf sCodigoRejeicao = 221 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10026)
      ElseIf sCodigoRejeicao = 232 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10055)
      ElseIf sCodigoRejeicao = 302 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10027)
      ElseIf sCodigoRejeicao = 305 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10028)
      ElseIf sCodigoRejeicao = 321 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10020)
      ElseIf sCodigoRejeicao = 531 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10053)
      ElseIf sCodigoRejeicao = 535 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10056)
      ElseIf sCodigoRejeicao = 539 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10054)
      ElseIf sCodigoRejeicao = 547 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10057)
      ElseIf sCodigoRejeicao = 690 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10021)
      ElseIf sCodigoRejeicao = 703 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10029)
      ElseIf sCodigoRejeicao = 732 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10030)
      ElseIf sCodigoRejeicao = 733 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10031)
      ElseIf sCodigoRejeicao = 770 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10024)
      ElseIf sCodigoRejeicao = 1007 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10023)
      ElseIf sCodigoRejeicao = 1425 Then
          Call objHelp.Show(strfile, "QuickStore10Help", 10032)
      Else
          ' Abre a pagina de indice geral das Rejeições
          Call objHelp.Show(strfile, "QuickStore10Help", 10013)
      End If
      
      Set objHelp = Nothing
  End If
  
  Exit Sub
Erro:
    MsgBox "Erro na chamada do Help. Cod: " & Err.Number & " Desc: " & Err.Description, vbInformation, "Erro"
End Sub

Private Sub cmd_imprimirDanfe_Click()
On Error GoTo trata_WebApiErro
  Dim url As String
  Dim objReq
  Dim objStream
  Dim nomeDanfePDF As String
  Dim nRow As Integer
  Dim bm As Variant
  
  nomeDanfePDF = ""
    
  With grdMovimento
    Screen.MousePointer = vbHourglass
    For nRow = 0 To .Rows - 1
      bm = .AddItemBookmark(nRow)
      
      If CBool(.Columns("Enviar").CellValue(bm)) Then
        nomeDanfePDF = gsHandleNull(.Columns("arquivoDanfe").CellValue(bm))
'''        Exit For
      Else
          nomeDanfePDF = ""
      End If
'''    Next nRow
'''  End With
  
'''      If LTrim(RTrim(nomeDanfePDF)) = "" Then
'''          MsgBox "Selecione uma NF-e autorizada na grade ou você selecionou alguma não autorizada.", vbInformation, "Atenção"
'''          Screen.MousePointer = vbDefault
'''          Exit Sub
'''      End If
  
      'url = "http://homolog.aclti.com.br/nfev5/GetDanfe.aspx?NomeDanfe=nfe_41180804152403000107550010000055991459682647_102735.pdf"
  
      If nomeDanfePDF <> "" Then
          '''nomeDanfePDF = App.Path & "\Danfe_NFe\" & nomeDanfePDF
          If Mid(sPastaEnvioNfe, Len(sPastaEnvioNfe), 1) <> "\" Then
              nomeDanfePDF = sPastaEnvioNfe & "\" & nomeDanfePDF
          Else
              nomeDanfePDF = sPastaEnvioNfe & nomeDanfePDF
          End If
        
          Set objReq = CreateObject("MSXML2.XMLHTTP")
          objReq.Open "GET", sCaminhoDanfe_Benefix & nomeDanfePDF, False
          objReq.Send
          
          If objReq.Status = 200 Then
              Set objStream = CreateObject("ADODB.Stream")
              objStream.Open
              objStream.Type = 1
          
              objStream.Write objReq.responseBody
              objStream.Position = 0
          
              objStream.SaveToFile nomeDanfePDF, 2
              objStream.Close
              Set objStream = Nothing
          End If
          
          Set objReq = Nothing
          Screen.MousePointer = vbDefault
        
          ShellExecute Me.hwnd, "open", nomeDanfePDF, vbNullString, vbNullString, SW_SHOWNORMAL
      End If
  
    Next nRow
  End With
  
  
  If Screen.MousePointer <> vbDefault Then
      Screen.MousePointer = vbDefault
  End If

  Exit Sub
    
trata_WebApiErro:
    If Screen.MousePointer <> vbDefault Then
        Screen.MousePointer = vbDefault
    End If

    MsgBox "Erro na impressão da Danfe! Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_imprimirDanfeCC_Click()
On Error GoTo trata_WebApiErro
  Dim url As String
  Dim objReq
  Dim objStream
  Dim nomeDanfePDF As String
  Dim nomeDanfePDF_Final As String
  Dim nRow As Integer
  Dim bm As Variant
    
  If grdCC.Row >= 0 Then
      bm = grdCC.AddItemBookmark(grdCC.Row)
      nomeDanfePDF = gsHandleNull(grdCC.Columns("arquivoDanfeCC").CellValue(bm))
      
      If nomeDanfePDF = "" Then
        Exit Sub
      End If
  End If

  'url = "http://homolog.aclti.com.br/nfev5/GetDanfe.aspx?NomeDanfe=nfe_41180804152403000107550010000055991459682647_102735.pdf"
  
  nomeDanfePDF_Final = ""
  
  If nomeDanfePDF <> "" Then
      Screen.MousePointer = vbHourglass
      
      'MsgBox "Caminho App path: " & App.Path, vbInformation, "Atenção"
      nomeDanfePDF_Final = App.Path & "\Danfe_NFe\" & nomeDanfePDF
      'MsgBox "Caminho físico do arquivo: " & nomeDanfePDF, vbInformation, "Atenção"
          
      Dim strAux As String
      Dim iAux As Integer
      iAux = Len(App.Path)
      If Mid(App.Path, iAux, 1) = "\" Then
          nomeDanfePDF_Final = App.Path & "Danfe_NFe\" & nomeDanfePDF
      End If
      'MsgBox "Caminho físico do arquivo: " & nomeDanfePDF_Final, vbInformation, "Atenção"
      
      Set objReq = CreateObject("MSXML2.XMLHTTP")
      objReq.Open "GET", sCaminhoDanfe_Benefix & nomeDanfePDF, False
      objReq.Send
      
      If objReq.Status = 200 Then
          Set objStream = CreateObject("ADODB.Stream")
          objStream.Open
          objStream.Type = 1
      
          objStream.Write objReq.responseBody
          objStream.Position = 0
      
          'objStream.SaveToFile nomeDanfePDF, 2
          objStream.SaveToFile nomeDanfePDF_Final, 2
          objStream.Close
          Set objStream = Nothing
      End If
      
      Set objReq = Nothing
      Screen.MousePointer = vbDefault
    
      'ShellExecute Me.hwnd, "open", nomeDanfePDF, vbNullString, vbNullString, SW_SHOWNORMAL
      ShellExecute Me.hwnd, "open", nomeDanfePDF_Final, vbNullString, vbNullString, SW_SHOWNORMAL
  End If

  Exit Sub
    
trata_WebApiErro:
    If Screen.MousePointer <> vbDefault Then
        Screen.MousePointer = vbDefault
    End If

    MsgBox "Erro na impressão da Danfe! Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_inutilizarNFe_Click()
On Error GoTo Erro:
  Screen.MousePointer = vbHourglass
  
  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
  
  If sPadraoArquivoIntegracao = "XML" Then
      InutilizarNFe_XML
      
      cmd_pesqInutilizadas_Click
'  Else
'    ' TXT
'    Call InutilizarNFe
  End If
  
  Screen.MousePointer = vbDefault
  
  MsgBox "NFe inutilizada(s) com sucesso!", vbInformation, "Sucesso"
  
  txt_anoInutilizacao.Text = ""
  txtSerie.Text = ""
  txtNumeroNFeInicial.Text = ""
  txtNumeroNFeFinal.Text = ""
  
  Exit Sub
Erro:
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro na chamada ao Método de Inutilização de NFe. Erro:" & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_mostrarProdutos_Click()
On Error GoTo Erro
  Dim rsProdutosAux As Recordset
  Dim strSQL As String
  
  If LTrim(RTrim(txt_sequencia.Text)) = "" Then
      MsgBox "Informe o número da sequência.", vbInformation, "Atenção"
      Exit Sub
  End If

  With grid_produtos
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With

  strSQL = "SELECT S.Código, P.Nome "
  strSQL = strSQL & " FROM [Saídas - Produtos] S, Produtos P "
  strSQL = strSQL & " Where S.Filial=" & gnCodFilial & " And S.Sequência = " & LTrim(RTrim(txt_sequencia.Text))
  strSQL = strSQL & " And S.Código = P.Código "
  strSQL = strSQL & " Order by S.Linha "

  Set rsProdutosAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsProdutosAux
    If Not (.BOF And .EOF) Then
      Do Until .EOF
          grid_produtos.AddItem .Fields("Código").Value & vbTab & _
                          .Fields("Nome").Value

        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsProdutosAux = Nothing

  Exit Sub
Erro:
    MsgBox "Erro na carga dos produtos da sequencia informada. Cod: " & Err.Number & " Desc: " & Err.Description, vbInformation, "Erro"

End Sub

Private Sub cmd_NFCe_xml_manutencao_Click()
  On Error GoTo ErrHandler
  
  Dim sChaveAcesso As String
  Dim bTemAlgumRegistroSelecionadoNaGrid As Boolean
  Dim strSQL As String
  Dim sRetorno As String
  Dim bm As Variant
  Dim nRow As Long
  Dim iIndice1 As Long
  Dim iIndice2 As Long
  Dim sStatus As String
  Dim arquivoLote As String
  Dim sDetalheAutorizacao As String
  Dim sExMessage As String
  Dim sRetNFCe As String
  Dim sSequencia As String
  Dim sStatusNaGrid As String

  bTemAlgumRegistroSelecionadoNaGrid = False
  
  ' Via SOAP
  If bSoapClient_MSSoapInit_NFCe = False Then
    Set soapclient_NFCe = New SoapClient30
    soapclient_NFCe.MSSoapInit sSoapClient_MSSoapInit_NFCe
    soapclient_NFCe.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL_NFCe
    bSoapClient_MSSoapInit_NFCe = True
  End If

  With grd_nfceNormal
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)

          If CBool(.Columns("Enviar").CellValue(bm)) Then

              sChaveAcesso = gsHandleNull(.Columns("ChaveAcesso").CellValue(bm))
              sRetNFCe = .Columns("retNFCe").CellValue(bm)
              sSequencia = .Columns("Sequencia").CellValue(bm)
              sStatusNaGrid = .Columns("Status").CellValue(bm)
              
              iIndice1 = InStr(1, sRetNFCe, "<detalheAutorizacao>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetNFCe, "</detalheAutorizacao>")
                sDetalheAutorizacao = Mid(sRetNFCe, iIndice1 + 20, iIndice2 - (iIndice1 + 20))
              Else
                sDetalheAutorizacao = ""
              End If
              
              iIndice1 = InStr(1, sRetNFCe, "<exMessage>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetNFCe, "</exMessage>")
                sExMessage = Mid(sRetNFCe, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
              Else
                sExMessage = ""
              End If

              bTemAlgumRegistroSelecionadoNaGrid = True

              'Chama WS
              sRetorno = soapclient_NFCe.GetDoc_Xml(sCNPJ, sChaveAcesso)

              Dim objTela As Form
              Set objTela = New frmXML_NFCe

              arquivoLote = PrettyPrintXML(sRetorno)

              frmXML_NFCe.bChamadorNFCeNormal = True
              frmXML_NFCe.sSequencia = sSequencia
              frmXML_NFCe.sCNPJ = sCNPJ
              frmXML_NFCe.sXML_Erro = sDetalheAutorizacao & " - " & sExMessage
              If arquivoLote <> "" Then
                frmXML_NFCe.sXML = arquivoLote
              Else
                frmXML_NFCe.sXML = sRetorno
              End If
              frmXML_NFCe.Show 1
              Set objTela = Nothing
              
              Exit For
              
          End If
      Next nRow
  End With

  If bTemAlgumRegistroSelecionadoNaGrid = False Then
      MsgBox "Selecione uma NFCe na grade.", vbInformation, "Informação"
  End If

  Exit Sub

ErrHandler:
  MsgBox "Erro ao processar Retorno : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"
End Sub

Private Sub cmd_pesqInutilizadas_Click()

  PesquisarInutilizadas
End Sub

Private Sub cmd_pesquisarCC_Click()
  PesquisarCartasCorrecao
End Sub

Private Sub cmd_pesquisarNFCe_Click()
On Error GoTo ErrHandler
  
  Dim rsMovimento As Recordset
  Dim strSQL As String
  Dim strStatus As String
  
  With grid_nfce_cont
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  'Saídas
  strSQL = "SELECT Data, Sequência, NFCe_contingencia_serie, NFCe_contingencia_num, Total, NFCe_contingencia_chave, NFCe_contingencia_status, retNFCe_contingencia "
  strSQL = strSQL & "FROM Saídas "
  strSQL = strSQL & "WHERE Saídas.Filial = " & gnCodFilial
  strSQL = strSQL & "AND Data =#" & Format(msk_dataDia.Text, "MM/DD/YYYY") & "# "
  strSQL = strSQL & "AND NFCe_contingencia_status in('Pendente','Erro','OK') "
  strSQL = strSQL & " ORDER BY Sequência DESC"

  Set rsMovimento = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsMovimento
    If Not (.BOF And .EOF) Then
      Do Until .EOF
      
          If .Fields("NFCe_contingencia_status").Value = "OK" Then
              strStatus = "Autorizado"
          Else
              strStatus = .Fields("NFCe_contingencia_status").Value
          End If
      
          'Adiciona registro
          grid_nfce_cont.AddItem "0" & vbTab & _
                          .Fields("Data").Value & vbTab & _
                          .Fields("Sequência").Value & vbTab & _
                          .Fields("NFCe_contingencia_serie").Value & vbTab & _
                          .Fields("NFCe_contingencia_num").Value & vbTab & _
                          Format(.Fields("Total").Value, FORMAT_VALUE) & vbTab & _
                          strStatus & vbTab & _
                          .Fields("NFCe_contingencia_chave").Value & vbTab & _
                          .Fields("retNFCe_contingencia").Value
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsMovimento = Nothing
    
  ssTab_contingencia.Tab = 0
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_pesquisarNFCeNormal_Click()
On Error GoTo ErrHandler
  
  Dim rsMovimento As Recordset
  Dim strSQL As String
  Dim sStatus As String
  Dim iIndice1 As Long
  Dim iIndice2 As Long
  
  With grd_nfceNormal
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  If msk_dataDiaNFCe.Text = "  /  /    " Then
      MsgBox "Informe o dia de pesquisa", vbInformation, "Atenção"
      msk_dataDiaNFCe.SetFocus
      Exit Sub
  End If
  
  'Saídas
  strSQL = "SELECT Data, Sequência, serieNF, NFCe, Total, retNFCe, ChaveNFCe "
  strSQL = strSQL & "FROM Saídas "
  strSQL = strSQL & "WHERE Saídas.Filial = " & gnCodFilial
  strSQL = strSQL & "AND Data =#" & Format(msk_dataDiaNFCe.Text, "MM/DD/YYYY") & "# "
  strSQL = strSQL & "AND NFCe_contingencia_num = 0 "
  strSQL = strSQL & " ORDER BY Sequência DESC"

  Set rsMovimento = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsMovimento
    If Not (.BOF And .EOF) Then
      Do Until .EOF
      
          sStatus = ""
          If Not IsNull(.Fields("retNFCe").Value) Then
              iIndice1 = InStr(1, .Fields("retNFCe").Value, "<detalheCancelamento>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, .Fields("retNFCe").Value, "</detalheCancelamento>")
                sStatus = Mid(.Fields("retNFCe").Value, iIndice1 + 21, iIndice2 - (iIndice1 + 21))
        
                iIndice1 = InStr(1, Mid(sStatus, 1, 4), "135")
                If iIndice1 > 0 Then
                  sStatus = "Cancelada"
                Else
                  sStatus = "Rejeitada/Erro"
                End If
'              ElseIf InStr(1, .Fields("retNFCe").Value, "<exMessage>") > 0 Then
'                iIndice2 = InStr(1, .Fields("retNFCe").Value, "</exMessage>")
'                sStatus = Mid(.Fields("retNFCe").Value, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
'
'                sStatus = "Verificar Status"
              Else
                iIndice1 = InStr(1, .Fields("retNFCe").Value, "<detalheAutorizacao>")
                If iIndice1 > 0 Then
                  iIndice2 = InStr(1, .Fields("retNFCe").Value, "</detalheAutorizacao>")
                  sStatus = Mid(.Fields("retNFCe").Value, iIndice1 + 20, iIndice2 - (iIndice1 + 20))
                
                  iIndice1 = InStr(1, Mid(sStatus, 1, 4), "100")
                  If iIndice1 > 0 Then
                    sStatus = "Autorizada"
                  Else
                    sStatus = "Rejeitada/Erro"
                  End If
                Else
                    sStatus = ""
                    iIndice1 = InStr(1, .Fields("retNFCe").Value, "<exMessage>")
                    If iIndice1 > 0 Then
                      iIndice2 = InStr(1, .Fields("retNFCe").Value, "</exMessage>")
                      sStatus = Mid(.Fields("retNFCe").Value, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
                    
                      iIndice1 = InStr(1, sStatus, "Autorizado o uso")
                      If iIndice1 > 0 Then
                        sStatus = "Autorizada"
                      Else
                        sStatus = "Verificar Status"
                      End If
                    End If
                End If
              End If
          End If
          
          'Adiciona registro
          grd_nfceNormal.AddItem "0" & vbTab & _
                          .Fields("Data").Value & vbTab & _
                          .Fields("Sequência").Value & vbTab & _
                          .Fields("serieNF").Value & vbTab & _
                          .Fields("NFCe").Value & vbTab & _
                          sStatus & vbTab & _
                          Format(.Fields("Total").Value, FORMAT_VALUE) & vbTab & _
                          .Fields("ChaveNFCe").Value & vbTab & _
                          .Fields("retNFCe").Value
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsMovimento = Nothing
    
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmd_posicaoConsulta_Click()
  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)

  ProcessarRetorno
  PesquisarRetornos
  
  cmdPesquisar_Click
  
  'Call ProcessarRetorno
  'Call PesquisarRetornos
End Sub

Private Sub cmd_posicaoNFCe_cont_Click()
  ProcessarRetornoNFCe
  
  cmd_pesquisarNFCe_Click
End Sub

Private Sub cmd_salvarProdutos_Click()
On Error GoTo Erro:
  Dim bm As Variant
  Dim iConta As Integer

  If LTrim(RTrim(txt_sequencia.Text)) = "" Then
      MsgBox "Informe o número da sequência.", vbInformation, "Atenção"
      Exit Sub
  End If

  grid_produtos.Update

  gSequenciaSaidas = CLng(txt_sequencia.Text)
  gProdutosXPed_nItemPedContador = grid_produtos.Rows

  If gProdutosXPed_nItemPedContador = 0 Then
      MsgBox "Grade sem produtos", vbInformation, "Atenção"
      Exit Sub
  End If

  For iConta = 0 To grid_produtos.Rows - 1
        bm = grid_produtos.AddItemBookmark(iConta)
        
        gProdutosXPed_nItemPed(iConta, 0) = grid_produtos.Columns(0).CellValue(bm)
        gProdutosXPed_nItemPed(iConta, 1) = grid_produtos.Columns(2).CellValue(bm)
        gProdutosXPed_nItemPed(iConta, 2) = grid_produtos.Columns(3).CellValue(bm)
  Next

  MsgBox "Informações de salvas com sucesso para a sequencia " & txt_sequencia.Text & ". Agora envie a NFe na Aba verde!", vbInformation, "Sucesso"

  Exit Sub
Erro:
  MsgBox "Erro ao salvar produtos da sequencia. Cod: " & Err.Number & " Desc: " & Err.Description, vbInformation, "Erro"

End Sub


Private Sub cmd_visualizaNFCe_cont_Click()
  On Error GoTo ErrHandler
  
  Dim sChaveAcessoCont As String
  Dim bTemAlgumRegistroSelecionadoNaGrid As Boolean
  Dim strSQL As String
  Dim sRetorno As String
  Dim bm As Variant
  Dim nRow As Long
  Dim iIndice1 As Long
  Dim iIndice2 As Long
  Dim sStatus As String
  Dim arquivoLote As String
  Dim sDetalheAutorizacao As String
  Dim sExMessage As String
  Dim sRetNFCeCont As String
  Dim sSequencia As String
  Dim sStatusNaGrid As String

  bTemAlgumRegistroSelecionadoNaGrid = False
  
  ' Via SOAP
  If bSoapClient_MSSoapInit_NFCe = False Then
    Set soapclient_NFCe = New SoapClient30
    soapclient_NFCe.MSSoapInit sSoapClient_MSSoapInit_NFCe
    soapclient_NFCe.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL_NFCe
    bSoapClient_MSSoapInit_NFCe = True
  End If

  With grid_nfce_cont
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)

          If CBool(.Columns("Enviar").CellValue(bm)) Then

              sChaveAcessoCont = gsHandleNull(.Columns("ChaveAcesso").CellValue(bm))
              sRetNFCeCont = .Columns("retNFCe").CellValue(bm)
              sSequencia = .Columns("Sequencia").CellValue(bm)
              sStatusNaGrid = .Columns("Status").CellValue(bm)
              
              iIndice1 = InStr(1, sRetNFCeCont, "<detalheAutorizacao>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetNFCeCont, "</detalheAutorizacao>")
                sDetalheAutorizacao = Mid(sRetNFCeCont, iIndice1 + 20, iIndice2 - (iIndice1 + 20))
              Else
                sDetalheAutorizacao = ""
              End If
              
              iIndice1 = InStr(1, sRetNFCeCont, "<exMessage>")
              If iIndice1 > 0 Then
                iIndice2 = InStr(1, sRetNFCeCont, "</exMessage>")
                sExMessage = Mid(sRetNFCeCont, iIndice1 + 11, iIndice2 - (iIndice1 + 11))
              Else
                sExMessage = ""
              End If

              bTemAlgumRegistroSelecionadoNaGrid = True

              'Chama WS
              sRetorno = soapclient_NFCe.GetDoc(sCNPJ, sChaveAcessoCont)

              Dim objTela As Form
              Set objTela = New frmXML_NFCe

              arquivoLote = PrettyPrintXML(sRetorno)

              objTela.bChamadorNFCeNormal = False
              objTela.sStatusDoCupomFiscalContingencia = sStatusNaGrid
              objTela.sSequencia = sSequencia
              objTela.sCNPJ = sCNPJ
              objTela.sXML_Erro = sDetalheAutorizacao & " - " & sExMessage
              If arquivoLote <> "" Then
                objTela.sXML = arquivoLote
              Else
                objTela.sXML = sRetorno
              End If
              objTela.Show 1
              Set objTela = Nothing
              
              Exit For
              
          End If
      Next nRow
  End With

  If bTemAlgumRegistroSelecionadoNaGrid = False Then
      MsgBox "Selecione uma NFCe na grade.", vbInformation, "Informação"
  End If

  Exit Sub

ErrHandler:
  MsgBox "Erro ao processar Retorno : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"
End Sub

Private Sub cmd_visualizarXML_AbaNFe_Click()
On Error GoTo Erro:
 
  If gAbreModuloXML = 0 Then
    'Caso não tenha acesso a este modulo...direcionar para tela de aquisição
    frmAquisicaoEstrategicoRel.Show
    Exit Sub
  End If
  
  
  Dim bm As Variant
  Dim sNumNFe As String
  Dim arquivoLote As String
  Dim arquivoLoteAux As String
  Dim sNomeArqXML As String
  Dim sSequenciaXML As String
  Dim sStatus As String
  Dim nRow As Integer
  
  sNumNFe = ""

  With grdMovimento

      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)
          
          If CBool(.Columns("Enviar").CellValue(bm)) Then
              sNumNFe = .Columns("NotaFiscal").CellValue(bm)
              sSequenciaXML = .Columns("Sequência").CellValue(bm)
              sStatus = .Columns("Status").CellValue(bm)
          
              If sNumNFe = "" Or sStatus = "Não Enviada" Then
                Screen.MousePointer = vbDefault
                Exit Sub
              End If
          
              'Abrir arquivo .xml
              Dim ff As Integer
              ff = FreeFile
              sNomeArqXML = sPastaEnvioNfe & "\NFeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_" & sNumNFe & ".xml"
              Open sNomeArqXML For Input As #ff
              
              Dim Linha As String
              While EOF(ff) = False
                  Linha = ""
                  Line Input #ff, Linha
                  arquivoLote = arquivoLote + Linha
              Wend
              Close #ff
              arquivoLoteAux = arquivoLote
              
              If Not IsNull(arquivoLote) And arquivoLote <> "" Then
                Dim objTela As Form
                
                Set objTela = New frmXML
                
                arquivoLote = PrettyPrintXML(arquivoLote)
                
                If arquivoLote <> "" Then
                  frmXML.sXML = arquivoLote
                Else
                  frmXML.sXML = arquivoLoteAux 'Possivelmente não abriu formatado, pois tem erro de formatacao o XML, então abrir sem formatação mesmo
                End If
                frmXML.xNomeArquivoXML = sNomeArqXML
                frmXML.sSequencia = sSequenciaXML
                frmXML.sXML_Erro = ""
                frmXML.iOrigemChamador = 2 ' 1-Tela frmNFe Aba Erros/Críticas     2-Tela frmNFe Aba Notas Fiscais
                frmXML.Show 1
              End If
              Set objTela = Nothing
          
              Exit For
          End If
      Next nRow
  End With
  
  If LTrim(RTrim(sNumNFe)) = "" Then
      MsgBox "Selecione uma Nota na grade", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If Screen.MousePointer <> vbDefault Then
      Screen.MousePointer = vbDefault
  End If
  
  Exit Sub
Erro:
    MsgBox "Erro na abertura do arquivo XML da NFe. Cod: " & Err.Number & " Desc: " & Err.Description, vbInformation, "Erro"
End Sub

Private Sub cmd_visualizarXML_nfceNormal_Click()
On Error GoTo ErrHandler
  
  Dim bTemAlgumRegistroSelecionadoNaGrid As Boolean
  Dim bm As Variant
  Dim nRow As Long
  Dim arquivoLote As String
  Dim sSequencia As String
  Dim sRetNFCeCont As String

  bTemAlgumRegistroSelecionadoNaGrid = False

  With grd_nfceNormal
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)

          If CBool(.Columns("Enviar").CellValue(bm)) Then
              sRetNFCeCont = .Columns("retNFCe").CellValue(bm)
              sSequencia = .Columns("Sequencia").CellValue(bm)
              
              bTemAlgumRegistroSelecionadoNaGrid = True
              
              Dim objTela As Form
              Set objTela = New frmXML_NFCe

              arquivoLote = PrettyPrintXML(sRetNFCeCont)
              frmXML_NFCe.bChamadorNFCeNormal = True
              frmXML_NFCe.sSequencia = sSequencia
              frmXML_NFCe.sXML = arquivoLote
              frmXML_NFCe.Show 1
              Set objTela = Nothing
              
              Exit For
          End If
      Next nRow
  End With

  If bTemAlgumRegistroSelecionadoNaGrid = False Then
      MsgBox "Selecione uma NFCe na grade.", vbInformation, "Informação"
  End If

  Exit Sub

ErrHandler:
  MsgBox "Erro ao processar Retorno : " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"
End Sub

Private Sub cmd_visualizaXML_Click()
On Error GoTo Erro:

  If gAbreModuloXML = 0 Then
    'Caso não tenha acesso a este modulo...direcionar para tela de aquisição
    frmAquisicaoEstrategicoRel.Show
    Exit Sub
  End If
  
  Dim bm As Variant
  Dim sNumNFe As String
  Dim arquivoLote As String
  Dim arquivoLoteAux As String
  Dim sDescricaoErro As String
  Dim sNomeArqXML As String
  Dim sSequenciaXML As String
  
  'If grdRetorno.Row >= 0 Then
  If gridRetorno.RowSel > 0 Then
      'bm = grdRetorno.AddItemBookmark(grdRetorno.Row)
      'sNumNFe = grdRetorno.Columns("Nota Fiscal").CellValue(bm)
      'sDescricaoErro = grdRetorno.Columns("Descrição Resposta").CellValue(bm)
      'sSequenciaXML = grdRetorno.Columns("Sequência").CellValue(bm)
      
      sNumNFe = gridRetorno.TextMatrix(gridRetorno.RowSel, 1)
      sDescricaoErro = gridRetorno.TextMatrix(gridRetorno.RowSel, 6)
      sSequenciaXML = gridRetorno.TextMatrix(gridRetorno.RowSel, 2)
      
      
      If sNumNFe = "" Then
        MsgBox "Selecione um registro na grade!", vbInformation, "NFe Visualizar XML"
        Exit Sub
      End If
      
      'Abrir arquivo .xml
      Dim ff As Integer
      ff = FreeFile
      sNomeArqXML = sPastaEnvioNfe & "\NFeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_" & sNumNFe & ".xml"
      Open sNomeArqXML For Input As #ff
  
      Dim Linha As String
      While EOF(ff) = False
          Linha = ""
          Line Input #ff, Linha
          arquivoLote = arquivoLote + Linha
      Wend
      Close #ff
      arquivoLoteAux = arquivoLote
    
      If Not IsNull(arquivoLote) And arquivoLote <> "" Then
        Dim objTela As Form
        
        Set objTela = New frmXML
        
        arquivoLote = PrettyPrintXML(arquivoLote)
        
        If arquivoLote <> "" Then
          frmXML.sXML = arquivoLote
        Else
          frmXML.sXML = arquivoLoteAux 'Possivelmente não abriu formatado, pois tem erro de formatacao o XML, então abrir sem formatação mesmo
        End If
        frmXML.sXML_Erro = sDescricaoErro
        frmXML.xNomeArquivoXML = sNomeArqXML
        frmXML.sSequencia = sSequenciaXML
        frmXML.iOrigemChamador = 1 ' 1-Tela frmNFe Aba Erros/Críticas     2-Tela frmNFe Aba Notas Fiscais
        frmXML.Show 1
      End If
  End If
  Set objTela = Nothing
  Exit Sub
Erro:
    MsgBox "Erro na abertura do arquivo XML da NFe. Cod: " & Err.Number & " Desc: " & Err.Description, vbInformation, "Erro"
End Sub

Public Function PrettyPrintXML(xml As String) As String

  Dim Reader As New SAXXMLReader30
  Dim Writer As New MXXMLWriter30

  Writer.indent = True
  Writer.standalone = False
  Writer.omitXMLDeclaration = False
  Writer.Encoding = "utf-8"

  Set Reader.contentHandler = Writer
  Set Reader.dtdHandler = Writer
  Set Reader.ErrorHandler = Writer

  Call Reader.putProperty("http://xml.org/sax/properties/declaration-handler", _
          Writer)
  Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", _
          Writer)

On Error GoTo Erro:
  Call Reader.Parse(xml)

  PrettyPrintXML = Writer.output

  Exit Function
Erro:
  MsgBox "XML **mal formatado**. Verifique a crítica/erro e faça o tratamento adequado.", vbCritical
End Function

Private Sub cmdMarcarDesmarcar_Click(Index As Integer)
  Dim nRow As Long
  
  On Error GoTo ErrHandler
  
  With grdMovimento
    .Redraw = False
    .MoveFirst
    For nRow = 1 To .Rows
      .Columns("Enviar").Value = Index
      .MoveNext
    Next nRow
    .Update
    .Scroll -32767, -32767
    .Redraw = True
  End With
  
  Exit Sub

ErrHandler:
  grdMovimento.Redraw = True
  MsgBox "Erro ao marcar/desmarcar movimentação: " & Err.Number & " (" & Err.Description & ").", vbCritical, "Erro"

End Sub

' Pilatti Outubro/17
Private Sub PesquisarPorSequencia(sSequenciaEfetivada As String)
On Error GoTo ErrHandler
  
  Dim rsMovimento As Recordset
  Dim strSQL As String
  Dim strStatus As String

  With grdMovimento
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  'Saídas
  strSQL = "SELECT Data, Sequência, [Nota Fiscal], Total, Cli_For.Código, Cli_For.Nome, [Nota Impressa], ChaveAcesso, [Operações Saída].Tipo, NFe.Status   "
  strSQL = strSQL & "FROM ((Saídas INNER JOIN Cli_For ON Saídas.Cliente = Cli_For.Código) "
  strSQL = strSQL & "LEFT JOIN NFe ON Saídas.Sequência = NFe.Sequencia) "
  strSQL = strSQL & "LEFT JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
  strSQL = strSQL & "WHERE Saídas.Sequência = " + sSequenciaEfetivada
  strSQL = strSQL & " AND (((Saídas.Filial) = " & gnCodFilial & ") "
  strSQL = strSQL & "AND ((Saídas.Efetivada)<>False) AND ( ([Operações Saída].Tipo)='V' OR ([Operações Saída].Tipo='G') OR ([Operações Saída].Tipo='E') )) "
  strSQL = strSQL & "ORDER BY Sequência DESC"
 
 
  Set rsMovimento = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsMovimento
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        '
        
        If rsMovimento.Fields("Status").Value = 3 Then
          strStatus = "Cancelada"
        Else
          If rsMovimento.Fields("Nota Impressa").Value <> 0 Then
           strStatus = "Enviada"
          Else
            strStatus = "Não Enviada"
          End If
        End If
        
        'Adiciona registro
        grdMovimento.AddItem "0" & vbTab & _
                          .Fields("Data").Value & vbTab & _
                          .Fields("Sequência").Value & vbTab & _
                          .Fields("Código").Value & vbTab & _
                          .Fields("Nome").Value & vbTab & _
                          .Fields("Nota Impressa").Value & vbTab & _
                          Format(.Fields("Total").Value, FORMAT_VALUE) & vbTab & _
                          strStatus & vbTab & _
                          .Fields("ChaveAcesso") '.Fields("Status").Value
        
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsMovimento = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"

End Sub
' Pilatti fim

Private Sub cmdPesquisar_Click()
On Error GoTo ErrHandler
  
  Dim rsMovimento As Recordset
  Dim strSQL As String
  Dim strStatus As String
  Dim boMostraRegistroNaGrade As Boolean
  
  With grdMovimento
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  'Tipo de movimentação
  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  
  '24/11/2010 - Andrea
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
   
  If m_intTipoMovimentoPesquisa = 0 Then
    'Entradas
    strSQL = "SELECT Data, Sequência, [Nota Fiscal], Total, Cli_For.Código, Cli_For.Nome, [Nota Impressa], ChaveAcesso, NFe.Status, NFe.Serie, NFe.ProtocoloAutorizacao, NFe.ProtocoloCancelamento, NFe.nomeDanfe, Entradas.Operação  "
    strSQL = strSQL & "FROM ((Entradas INNER JOIN Cli_For ON Entradas.Fornecedor = Cli_For.Código) "
    strSQL = strSQL & "LEFT JOIN NFe ON Entradas.Sequência = NFe.Sequencia and Entradas.Filial = nfe.Filial) "
    
    '14/04/2010 - Andrea
    'Relacionado com a tabela de operacoes de entradas para poder filtrar as operacoes que não vão para NFe
    strSQL = strSQL & "LEFT JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código "
    
    strSQL = strSQL & "WHERE (((Entradas.Filial) = " & gnCodFilial & ") "
    
    
    ''##############################################################
    '' PABLO - 14/10/2022
    ''##############################################################
    ' Busca pelo intervalo somente se a movimentação não tiver sido selecionada
    If param_sequencia = 0 Then
        strSQL = strSQL & "AND (Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
        strSQL = strSQL & "AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
    End If
    ''##############################################################
    
    strSQL = strSQL & "AND ((Entradas.Efetivada) <> False) AND (([Operações Entrada].Tipo)<>'A')) "
    
    
    ''##############################################################
    '' PABLO - 14/10/2022
    ''##############################################################
    ' busca pela movimentação da tela de saída
    If param_sequencia > 0 Then
        strSQL = strSQL & " AND Sequência = " & param_sequencia & " "
    End If
    ''##############################################################
    
    If chk_ordemNumNFe.Value = vbChecked Then
        strSQL = strSQL & " ORDER BY [Nota Impressa] DESC"
    Else
        strSQL = strSQL & " ORDER BY Sequência DESC"
    End If
    
  Else
    'Saídas
    strSQL = "SELECT Data, Sequência, [Nota Fiscal], Total, Cli_For.Código, Cli_For.Nome, [Nota Impressa], ChaveAcesso, [Operações Saída].Tipo, NFe.Status, NFe.ProtocoloAutorizacao, NFe.ProtocoloCancelamento, NFe.Serie, NFe.nomeDanfe, Saídas.Operação   "
    strSQL = strSQL & "FROM ((Saídas INNER JOIN Cli_For ON Saídas.Cliente = Cli_For.Código) "
    strSQL = strSQL & "LEFT JOIN NFe ON Saídas.Sequência = NFe.Sequencia and Saídas.Filial = nfe.Filial) "
    strSQL = strSQL & "LEFT JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
  
    '13/04/2010 - Andrea
    'Excluído o fitro de tipo de cliente = C
    'strSQL = strSQL & "WHERE (((Saídas.Filial) = " & gnCodFilial & ")  AND ((Cli_For.Tipo) = 'C') "
    strSQL = strSQL & "WHERE (((Saídas.Filial) = " & gnCodFilial & ") "
    
    
    ''##############################################################
    '' PABLO - 14/10/2022
    ''##############################################################
    ' Busca pelo intervalo somente se a movimentação não tiver sido selecionada
    If param_sequencia = 0 Then
        strSQL = strSQL & "AND (Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
        strSQL = strSQL & "AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
    End If
    ''##############################################################
    
    
    
    'strSQL = strSQL & "AND Efetivada "
    strSQL = strSQL & "AND ((Saídas.Efetivada)<>False) AND ( ([Operações Saída].Tipo)='V' OR ([Operações Saída].Tipo='G') OR ([Operações Saída].Tipo='E') OR ([Operações Saída].Tipo='T') )) "
    
    
    ''##############################################################
    '' PABLO - 14/10/2022
    ''##############################################################
    ' busca pela movimentação da tela de saída
    If param_sequencia > 0 Then
        strSQL = strSQL & " AND Sequência = " & param_sequencia & " "
    End If
    ''##############################################################
    
    
    
    If chk_naoNFCe.Value = 1 Then 'Se igual a 1...busca apenas as Vendas que não foram enviadas como NFCe
      strSQL = strSQL & " AND (Saídas.ChaveNFCe = '' or Saídas.ChaveNFCe = null) "
    End If
    
    If chk_ordemNumNFe.Value = vbChecked Then
        strSQL = strSQL & " ORDER BY [Nota Impressa] DESC"
    Else
        strSQL = strSQL & " ORDER BY Sequência DESC"
    End If
  End If
 
  Set rsMovimento = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsMovimento
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        If rsMovimento.Fields("Status").Value = 3 Then
          strStatus = "Cancelada"
        ElseIf rsMovimento.Fields("Status").Value = 100 Then
          strStatus = "Autorizada"
        ElseIf rsMovimento.Fields("Status").Value = 101 Then
          strStatus = "Cancelada"
        ElseIf rsMovimento.Fields("Status").Value = 135 Then
          strStatus = "Cancelada"
        ElseIf rsMovimento.Fields("Status").Value < 0 Then
          strStatus = "Erro"
        Else
          If rsMovimento.Fields("Nota Impressa").Value <> 0 Then
           strStatus = "Enviada"
          Else
            strStatus = "Não Enviada"
          End If
        End If
        
        boMostraRegistroNaGrade = False
        
        ' Visualizar:
        '   TODAS
        '   AUTORIZADAS
        '   CANCELADAS
        '   COM ERRO
        '   ENVIADAS
        '   NÃO ENVIADAS
        
        Dim sCodigoOperLimpo As String
        
        If cbo_visualizarNFeGradeVerde.Text = "TODAS" Or cbo_visualizarNFeGradeVerde.Text = "" Then
            If cbo_operacao.ListIndex > 0 Then
                sCodigoOperLimpo = Trim(Mid(cbo_operacao.Text, 1, 3))
                If sCodigoOperLimpo = rsMovimento.Fields("Operação").Value Then
                    boMostraRegistroNaGrade = True
                Else
                    boMostraRegistroNaGrade = False
                End If
            Else
                boMostraRegistroNaGrade = True
            End If
        ElseIf cbo_visualizarNFeGradeVerde.Text = "AUTORIZADAS" And strStatus = "Autorizada" Then
            If cbo_operacao.ListIndex > 0 Then
                sCodigoOperLimpo = Trim(Mid(cbo_operacao.Text, 1, 3))
                If sCodigoOperLimpo = rsMovimento.Fields("Operação").Value Then
                    boMostraRegistroNaGrade = True
                Else
                    boMostraRegistroNaGrade = False
                End If
            Else
                boMostraRegistroNaGrade = True
            End If
        ElseIf cbo_visualizarNFeGradeVerde.Text = "CANCELADAS" And strStatus = "Cancelada" Then
            If cbo_operacao.ListIndex > 0 Then
                sCodigoOperLimpo = Trim(Mid(cbo_operacao.Text, 1, 3))
                If sCodigoOperLimpo = rsMovimento.Fields("Operação").Value Then
                    boMostraRegistroNaGrade = True
                Else
                    boMostraRegistroNaGrade = False
                End If
            Else
                boMostraRegistroNaGrade = True
            End If
        ElseIf cbo_visualizarNFeGradeVerde.Text = "COM ERRO" And strStatus = "Erro" Then
            If cbo_operacao.ListIndex > 0 Then
                sCodigoOperLimpo = Trim(Mid(cbo_operacao.Text, 1, 3))
                If sCodigoOperLimpo = rsMovimento.Fields("Operação").Value Then
                    boMostraRegistroNaGrade = True
                Else
                    boMostraRegistroNaGrade = False
                End If
            Else
                boMostraRegistroNaGrade = True
            End If
        ElseIf cbo_visualizarNFeGradeVerde.Text = "ENVIADAS" And strStatus = "Enviada" Then
            If cbo_operacao.ListIndex > 0 Then
                sCodigoOperLimpo = Trim(Mid(cbo_operacao.Text, 1, 3))
                If sCodigoOperLimpo = rsMovimento.Fields("Operação").Value Then
                    boMostraRegistroNaGrade = True
                Else
                    boMostraRegistroNaGrade = False
                End If
            Else
                boMostraRegistroNaGrade = True
            End If
        ElseIf cbo_visualizarNFeGradeVerde.Text = "NÃO ENVIADAS" And strStatus = "Não Enviada" Then
            If cbo_operacao.ListIndex > 0 Then
                sCodigoOperLimpo = Trim(Mid(cbo_operacao.Text, 1, 3))
                If sCodigoOperLimpo = rsMovimento.Fields("Operação").Value Then
                    boMostraRegistroNaGrade = True
                Else
                    boMostraRegistroNaGrade = False
                End If
            Else
                boMostraRegistroNaGrade = True
            End If
        End If
        
        If boMostraRegistroNaGrade = True Then
            If m_intTipoMovimentoPesquisa = 0 Then
                grdMovimento.AddItem "0" & vbTab & _
                              .Fields("Data").Value & vbTab & _
                              .Fields("Sequência").Value & vbTab & _
                              .Fields("Código").Value & vbTab & _
                              .Fields("Nome").Value & vbTab & _
                              .Fields("Serie").Value & vbTab & _
                              .Fields("Nota Impressa").Value & vbTab & _
                              Format(.Fields("Total").Value, FORMAT_VALUE) & vbTab & _
                              strStatus & vbTab & _
                              .Fields("ChaveAcesso") & vbTab & _
                              .Fields("ProtocoloAutorizacao") & vbTab & _
                              .Fields("ProtocoloCancelamento") & vbTab & _
                              .Fields("nomeDanfe")
    
            Else
            
                'Adiciona registro
                grdMovimento.AddItem "0" & vbTab & _
                              .Fields("Data").Value & vbTab & _
                              .Fields("Sequência").Value & vbTab & _
                              .Fields("Código").Value & vbTab & _
                              .Fields("Nome").Value & vbTab & _
                              .Fields("Serie").Value & vbTab & _
                              .Fields("Nota Impressa").Value & vbTab & _
                              Format(.Fields("Total").Value, FORMAT_VALUE) & vbTab & _
                              strStatus & vbTab & _
                              .Fields("ChaveAcesso") & vbTab & _
                              .Fields("ProtocoloAutorizacao") & vbTab & _
                              .Fields("ProtocoloCancelamento") & vbTab & _
                              .Fields("nomeDanfe")
                              '.Fields("Status").Value
             End If
        End If
        
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsMovimento = Nothing
    
  tabMain.Tab = 0
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub PesquisarCartasCorrecao()
On Error GoTo ErrHandler
  
  Dim rsCC As Recordset
  Dim strSQL As String
  
  With grdCC
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  strSQL = "SELECT CNPJ, Filial, Serie, Numero, Descricao, DataHora, arquivoDanfeCC "
  strSQL = strSQL & "FROM NFeCartaCorrecao "
  strSQL = strSQL & "ORDER BY DataHora Desc "
    
  Set rsCC = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsCC
    If Not (.BOF And .EOF) Then
      Do Until .EOF
      
          'Adiciona registro
          grdCC.AddItem .Fields("DataHora").Value & vbTab & _
                  .Fields("CNPJ").Value & vbTab & _
                  .Fields("Serie").Value & vbTab & _
                  .Fields("Numero").Value & vbTab & _
                  .Fields("Descricao").Value & vbTab & _
                  .Fields("arquivoDanfeCC").Value & vbTab
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCC = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros de NFe's Cartas Correção. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub PesquisarInutilizadas()
On Error GoTo ErrHandler
  
  Dim rsInutilizadas As Recordset
  Dim strSQL As String
  
  With grdInutilizadas
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  If opt_inutilizadasNFe.Value = True Then
      strSQL = "SELECT CNPJ, Filial, Ano, Serie, NumeroInicial, NumeroFinal, Justificativa, DataHora, Modelo "
      strSQL = strSQL & " FROM NFeInutilizadas "
      strSQL = strSQL & " WHERE (Isnull(Modelo) or Modelo = 55) "    '**** daí ira trazer 55 e sem valor tb
      strSQL = strSQL & " AND Filial = " & gnCodFilial
      strSQL = strSQL & " ORDER BY DataHora Desc "
  Else
      strSQL = "SELECT CNPJ, Filial, Ano, Serie, NumeroInicial, NumeroFinal, Justificativa, DataHora, Modelo "
      strSQL = strSQL & " FROM NFeInutilizadas "
      strSQL = strSQL & " WHERE Modelo = 65 "
      strSQL = strSQL & " AND Filial = " & gnCodFilial
      strSQL = strSQL & " ORDER BY DataHora Desc "
  End If
    
  Set rsInutilizadas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsInutilizadas
    If Not (.BOF And .EOF) Then
      Do Until .EOF
      
          'Adiciona registro
          grdInutilizadas.AddItem .Fields("CNPJ").Value & vbTab & _
                  .Fields("Ano").Value & vbTab & _
                  .Fields("Serie").Value & vbTab & _
                  .Fields("NumeroInicial").Value & vbTab & _
                  .Fields("NumeroFinal").Value & vbTab & _
                  .Fields("Justificativa").Value & vbTab & _
                  .Fields("DataHora").Value & vbTab
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsInutilizadas = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros de NFe's Inutilizadas. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub PesquisarRetornos()
On Error GoTo ErrHandler
  
  Dim rsRetornos As Recordset
  Dim strSQL As String
  Dim strTipoMovimentacao As String
  Dim strVazio As String
  strVazio = ""

  gridRetorno.Rows = 1
  gridRetorno.Row = 0

'  With grdRetorno
'    .Redraw = False
'    .RemoveAll
'    .Redraw = True
'  End With

  'Tipo de movimentação
  m_intTipoMovimentoPesquisa = cboTipoMovimento.ListIndex
  
  m_intTipoEmissao = (cboTipoEmissao.ListIndex + 1)
   
  If m_intTipoMovimentoPesquisa = 0 Then
    'Entradas
''    strSQL = "SELECT Data, Sequência, [Nota Fiscal], Cli_For.Código, Cli_For.Nome, [Nota Impressa], DataHora, Protocolo,DigestValue, NFeRetorno.Status,StatusDescricao,ProtocoloCancelamento  "
''    strSQL = strSQL & "FROM ((Entradas INNER JOIN Cli_For ON Entradas.Fornecedor = Cli_For.Código) "
''    strSQL = strSQL & "INNER JOIN NFeRetorno ON Entradas.Sequência = NFeRetorno.Sequencia and Entradas.Filial = NFeRetorno.Filial) INNER JOIN NFe ON Entradas.Sequência = NFe.Sequencia "
''    strSQL = strSQL & "WHERE (((Entradas.Filial) = " & gnCodFilial & ") "
''    strSQL = strSQL & "AND (Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
''    strSQL = strSQL & "AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#)) "
''    strSQL = strSQL & "ORDER BY Sequência DESC"

      strSQL = "SELECT Data, Sequência, [Nota Fiscal], Cli_For.Código, Cli_For.Nome, [Nota Impressa],  DataHora, Protocolo,DigestValue, NFeRetorno.Status,StatusDescricao,ProtocoloCancelamento "
      strSQL = strSQL & " From Cli_For, Entradas, nfe, NfeRetorno "
      strSQL = strSQL & " Where Entradas.Filial = " & gnCodFilial
      strSQL = strSQL & " AND (Entradas.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
      strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
      strSQL = strSQL & " and Entradas.Fornecedor = Cli_For.Código "
      strSQL = strSQL & " and Entradas.Sequência = Nfe.sequencia "
      strSQL = strSQL & " and Entradas.Filial = NFe.filial "
      strSQL = strSQL & " and NFe.sequencia = NFeRetorno.sequencia "
      strSQL = strSQL & " and NFe.filial = nferetorno.filial "

    strTipoMovimentacao = "Entrada"

  Else
    'Saídas
    
''    strSQL = "SELECT Data, Sequência, [Nota Fiscal], Cli_For.Código, Cli_For.Nome, [Nota Impressa],  DataHora, Protocolo,DigestValue, NFeRetorno.Status,StatusDescricao,ProtocoloCancelamento   "
''    strSQL = strSQL & "FROM ((Saídas INNER JOIN Cli_For ON Saídas.Cliente = Cli_For.Código) "
''    strSQL = strSQL & "INNER JOIN NFeRetorno ON Saídas.Sequência = NFeRetorno.Sequencia and Saídas.Filial = NFeRetorno.Filial)  INNER JOIN NFe ON Saídas.Sequência = NFe.Sequencia "
''    strSQL = strSQL & "WHERE (((Saídas.Filial) = " & gnCodFilial & ") "
''    strSQL = strSQL & "AND (Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
''    strSQL = strSQL & "AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#)) "
''    strSQL = strSQL & "ORDER BY Sequência DESC"

      strSQL = "SELECT Data, Sequência, [Nota Fiscal], Cli_For.Código, Cli_For.Nome, [Nota Impressa],  DataHora, Protocolo,DigestValue, NFeRetorno.Status,StatusDescricao,ProtocoloCancelamento "
      strSQL = strSQL & " From Cli_For, Saídas, nfe, NfeRetorno "
      strSQL = strSQL & " Where Saídas.Filial = " & gnCodFilial
      strSQL = strSQL & " AND (Saídas.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
      strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
      strSQL = strSQL & " and Saídas.Cliente = Cli_For.Código "
      strSQL = strSQL & " and Saídas.Sequência = Nfe.sequencia "
      strSQL = strSQL & " and Saídas.Filial = NFe.filial "
      strSQL = strSQL & " and NFe.sequencia = NFeRetorno.sequencia "
      strSQL = strSQL & " and NFe.filial = nferetorno.filial "
      strSQL = strSQL & " ORDER BY Sequência DESC "

      strTipoMovimentacao = "Saida"

  End If
 
 
  Dim sStatusDescricaoAux As String
 
  Set rsRetornos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsRetornos
    If Not (.BOF And .EOF) Then
    
      Do Until .EOF
      
        sStatusDescricaoAux = TratarCaracteresEspeciais0001(.Fields("StatusDescricao").Value)
      
        If m_intTipoMovimentoPesquisa = 0 Then
          
                gridRetorno.AddItem vbTab & .Fields("Nota Impressa").Value & vbTab & _
                          .Fields("Sequência").Value & vbTab & _
                          strTipoMovimentacao & vbTab & _
                          .Fields("DataHora").Value & vbTab & _
                          .Fields("Status").Value & vbTab & _
                          sStatusDescricaoAux & vbTab & _
                          .Fields("Protocolo").Value & vbTab & _
                          .Fields("ProtocoloCancelamento").Value & vbTab & _
                          .Fields("DigestValue").Value & vbTab & _
                          .Fields("Código").Value & vbTab & _
                          .Fields("Nome").Value
        Else
        
            If .Fields("Status").Value <> "0" Then
              'Adiciona registro
              gridRetorno.AddItem vbTab & .Fields("Nota Impressa").Value & vbTab & _
                          .Fields("Sequência").Value & vbTab & _
                          strTipoMovimentacao & vbTab & _
                          .Fields("DataHora").Value & vbTab & _
                          .Fields("Status").Value & vbTab & _
                          sStatusDescricaoAux & vbTab & _
                          .Fields("Protocolo").Value & vbTab & _
                          .Fields("ProtocoloCancelamento").Value & vbTab & _
                          .Fields("DigestValue").Value & vbTab & _
                          .Fields("Código").Value & vbTab & _
                          .Fields("Nome").Value

            End If
         End If
        
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsRetornos = Nothing
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao exibir registros de retorno. Cod: " & Err.Number & " Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmdPesquisarRetornos_Click()
  Call PesquisarRetornos
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


Private Sub dteEntradaContingencia_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      dteEntradaContingencia.Text = frmCalendario.gsDateCalender(dteEntradaContingencia.Text)
  End Select
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub Form_Load()
  If Not IsNull(origemTelaSaidasParaTelaNFe) And Len(origemTelaSaidasParaTelaNFe) > 0 Then
    'MsgBox "Veio da tela de Saidas!!"
        
    tabMain.Top = 120
    tabMain.Left = 120
    'fraDetalhes.Top = 4320
    'fraDetalhes.Left = 120
    'fraPesquisa.Visible = False
    'Frame1.Visible = False
    Me.Height = 7155
    
    PesquisarPorSequencia origemTelaSaidasParaTelaNFe
    origemTelaSaidasParaTelaNFe = ""
  Else
    '24/11/2010 - Andrea
    cboTipoEmissao.ListIndex = 0
    
    cboTipoMovimento.ListIndex = 1
    Data_Ini.Text = Format(Data_Atual, "dd/MM/yyyy")
    Data_Fim.Text = Format(Data_Atual, "dd/MM/yyyy")
    
    '25/11/2010 - Andrea
    dteEntradaContingencia.Text = Format(Data_Atual, "dd/MM/yyyy")
  End If
  
  msk_dataDia.Text = Format(Data_Atual, "dd/MM/yyyy")
  msk_dataDiaNFCe.Text = Format(Data_Atual, "dd/MM/yyyy")
  
  txt_anoInutilizacao.Text = Year(Now)
  txt_tipoDocCC.Text = "S"   ' S de Saída
  
  'Parâmetros da Filial
  Dim rsParamX As Recordset
  Set rsParamX = db.OpenRecordset("SELECT PastaEnvioNfe, PadraoArquivoIntegracao FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset, dbReadOnly)
  sPastaEnvioNfe = rsParamX.Fields("PastaEnvioNfe").Value
  sPadraoArquivoIntegracao = rsParamX.Fields("PadraoArquivoIntegracao").Value
  rsParamX.Close
  Set rsParamX = Nothing
 
  txtMotivoInutilizacao.Text = "Erro na digitacao dos dados"
  txtMotivoCancelamento.Text = "Erro na digitacao dos dados"
  
  
  ' Grade de Erros e Críticas do XML
  gridRetorno.ColWidth(0) = 280
  gridRetorno.ColWidth(1) = 1000
  gridRetorno.ColWidth(2) = 1000
  gridRetorno.ColWidth(3) = 1000
  gridRetorno.ColWidth(4) = 1800
  gridRetorno.ColWidth(5) = 1200
  gridRetorno.ColWidth(6) = 9500
  gridRetorno.ColWidth(7) = 2200
  gridRetorno.ColWidth(8) = 2200
  gridRetorno.ColWidth(9) = 2000
  gridRetorno.ColWidth(10) = 2000
  gridRetorno.ColWidth(11) = 2500

  gridRetorno.Row = 0
  gridRetorno.TextMatrix(0, 1) = "Nota Fiscal"
  gridRetorno.TextMatrix(0, 2) = "Sequência"
  gridRetorno.TextMatrix(0, 3) = "Tipo"
  gridRetorno.TextMatrix(0, 4) = "Data Hora Arquivo"
  gridRetorno.TextMatrix(0, 5) = "Nº Retorno"
  gridRetorno.TextMatrix(0, 6) = "Descrição Resposta"
  gridRetorno.TextMatrix(0, 7) = "Protocolo Autorização"
  gridRetorno.TextMatrix(0, 8) = "Protocolo Cancelamento"
  gridRetorno.TextMatrix(0, 9) = "Digest Value"
  gridRetorno.TextMatrix(0, 10) = "Cliente"
  gridRetorno.TextMatrix(0, 11) = "Nome Cliente"

  'Tratamento de dados da aba preta
  If nfeInfAdProd = True Then
    chk_nfeInfAdProd.Value = 1
  Else
    chk_nfeInfAdProd.Value = 0
  End If
  
  If nfeDevolucao_impostoDevol = True Then
    chk_nfeDevolucao_impostoDevol.Value = 1
  Else
    chk_nfeDevolucao_impostoDevol.Value = 0
  End If
  
  If nfe_xPed_nItemPed = True Then
    chk_xPed_nItemPed.Value = 1
    txt_sequencia.Enabled = True
    cmd_mostrarProdutos.Enabled = True
    grid_produtos.Enabled = True
    cmd_salvarProdutos.Enabled = True
  Else
    chk_xPed_nItemPed.Value = 0
  End If
  
  If nfe_xPed_nItemPed = False Then
      txt_sequencia.Enabled = False
      cmd_mostrarProdutos.Enabled = False
      grid_produtos.Enabled = False
      cmd_salvarProdutos.Enabled = False
  End If
  'fim tratamento aba preta
  
  Dim strSQL As String
  Dim rsParametros As Recordset
  'Parâmetros da Filial
  strSQL = "SELECT CGC FROM [Parâmetros Filial] "
  strSQL = strSQL & "WHERE Filial = " & gnCodFilial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsParametros
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
      .Close
      Set rsParametros = Nothing
      Exit Sub
    Else
      'Validações
      sCNPJ = Trim(.Fields("CGC").Value)
      sCNPJ = Replace(sCNPJ, "/", "")
      sCNPJ = Replace(sCNPJ, "-", "")
      sCNPJ = Replace(sCNPJ, " ", "")
      sCNPJ = Replace(sCNPJ, ".", "")
      sCNPJ = Replace(sCNPJ, ",", "")
    End If
  End With
  rsParametros.Close
  Set rsParametros = Nothing
  
  grd_nfceNormal.StyleSets("Vermelho").ForeColor = RGB(255, 0, 0)
  grd_nfceNormal.StyleSets("Preto").ForeColor = RGB(0, 0, 0)
  grdMovimento.StyleSets("Vermelho").ForeColor = RGB(255, 0, 0)
  grdMovimento.StyleSets("Preto").ForeColor = RGB(0, 0, 0)

  tabMain.Tab = 0
  
  
  ''##############################################################
  '' PABLO - 14/10/2022
  ''##############################################################
  ' pesquisar assim que a janela for aberta
  Call cmdPesquisar_Click
  
  ' selecionar o movimento da janela de venda
  Call SelecionarMovimentoUnico(True)
  ''##############################################################
 
End Sub

''##############################################################
'' PABLO - 03/11/2022
''##############################################################
Private Sub SelecionarMovimentoUnico(p_enviar As Boolean)
  ' selecionar o movimento da janela de venda
  If grdMovimento.Rows = 1 And param_sequencia > 0 Then
      grdMovimento.Columns("Enviar").Value = 1
      grdMovimento.Update
      If p_enviar Then
          Call cmd_enviarNFe_Click
      End If
  End If
End Sub

''##############################################################
'' PABLO - 03/11/2022
''##############################################################
Private Sub Form_Paint()
    Call SelecionarMovimentoUnico(False)
End Sub

''##############################################################
'' PABLO - 14/10/2022
''##############################################################
Private Sub Form_Unload(Cancel As Integer)
    param_sequencia = 0
End Sub
''##############################################################



Private Sub grd_nfceNormal_Change()
    grd_nfceNormal.Update
End Sub

Private Sub grd_nfceNormal_RowLoaded(ByVal Bookmark As Variant)
  Dim sStatus As String
 
  If IsEmpty(Bookmark) Then Exit Sub
 
  sStatus = grd_nfceNormal.Columns("Status").CellText(Bookmark)
  
  If sStatus = "Rejeitada/Erro" Then
      grd_nfceNormal.Columns("Data").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("Sequencia").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("Serie").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("Nota Fiscal").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("Total").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("retNFCe").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("Status").CellStyleSet "Vermelho"
      grd_nfceNormal.Columns("ChaveAcesso").CellStyleSet "Vermelho"
  Else
      grd_nfceNormal.Columns("Data").CellStyleSet "Preto"
      grd_nfceNormal.Columns("Sequencia").CellStyleSet "Preto"
      grd_nfceNormal.Columns("Serie").CellStyleSet "Preto"
      grd_nfceNormal.Columns("Nota Fiscal").CellStyleSet "Preto"
      grd_nfceNormal.Columns("Total").CellStyleSet "Preto"
      grd_nfceNormal.Columns("retNFCe").CellStyleSet "Preto"
      grd_nfceNormal.Columns("Status").CellStyleSet "Preto"
      grd_nfceNormal.Columns("ChaveAcesso").CellStyleSet "Preto"
  End If
  
End Sub

Private Sub grdMovimento_Change()
  grdMovimento.Update
End Sub

Private Sub grdMovimento_RowLoaded(ByVal Bookmark As Variant)
Dim sStatus As String
 
  If IsEmpty(Bookmark) Then Exit Sub
 
  sStatus = grdMovimento.Columns("Status").CellText(Bookmark)
  
  If sStatus = "Erro" Then
      grdMovimento.Columns("Data").CellStyleSet "Vermelho"
      grdMovimento.Columns("Sequencia").CellStyleSet "Vermelho"
      grdMovimento.Columns("Serie").CellStyleSet "Vermelho"
      grdMovimento.Columns("Nota Fiscal").CellStyleSet "Vermelho"
      grdMovimento.Columns("Total").CellStyleSet "Vermelho"
      grdMovimento.Columns("Código").CellStyleSet "Vermelho"
      grdMovimento.Columns("Nome Cliente/Fornecedor").CellStyleSet "Vermelho"
      grdMovimento.Columns("ChaveAcesso").CellStyleSet "Vermelho"
      grdMovimento.Columns("Status").CellStyleSet "Vermelho"
  Else
      grdMovimento.Columns("Data").CellStyleSet "Preto"
      grdMovimento.Columns("Sequencia").CellStyleSet "Preto"
      grdMovimento.Columns("Serie").CellStyleSet "Preto"
      grdMovimento.Columns("Nota Fiscal").CellStyleSet "Preto"
      grdMovimento.Columns("Total").CellStyleSet "Preto"
      grdMovimento.Columns("Código").CellStyleSet "Preto"
      grdMovimento.Columns("Nome Cliente/Fornecedor").CellStyleSet "Preto"
      grdMovimento.Columns("ChaveAcesso").CellStyleSet "Preto"
      grdMovimento.Columns("Status").CellStyleSet "Preto"
  End If
End Sub

Private Sub grid_nfce_cont_Change()
  grid_nfce_cont.Update
End Sub

Private Sub gridRetorno_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  gridRetorno.Redraw = False
End Sub

Private Sub gridRetorno_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  gridRetorno.RowSel = gridRetorno.Row
  gridRetorno.Redraw = True
End Sub

Private Sub msk_dataDia_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      msk_dataDia.Text = frmCalendario.gsDateCalender(msk_dataDia.Text)
  End Select
End Sub

Private Sub opt_inutilizadasNFCe_Click()
  If opt_inutilizadasNFCe.Value = True Then
      cmd_pesqInutilizadas.Caption = "Pesquisar NFC-e/Cupom Fiscal Inutilizados"
      lblNumeroNFeInicial(0).Caption = "NFCe Número Inicial"
      lblNumeroNFeFinal(1).Caption = "NFCe Número Final"
      cmd_inutilizarNFe.Caption = "Inutilizar NFCe/Cupom Fiscal"
  Else
      cmd_pesqInutilizadas.Caption = "Pesquisar NF-e Inutilizadas"
      lblNumeroNFeInicial(0).Caption = "NFe Número Inicial"
      lblNumeroNFeFinal(1).Caption = "NFe Número Final"
      cmd_inutilizarNFe.Caption = "Inutilizar NFe"
  End If

  With grdInutilizadas
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With

End Sub

Private Sub opt_inutilizadasNFe_Click()
  If opt_inutilizadasNFe.Value = True Then
      cmd_pesqInutilizadas.Caption = "Pesquisar NF-e Inutilizadas"
      lblNumeroNFeInicial(0).Caption = "NFe Número Inicial"
      lblNumeroNFeFinal(1).Caption = "NFe Número Final"
      cmd_inutilizarNFe.Caption = "Inutilizar NFe"
  Else
      cmd_pesqInutilizadas.Caption = "Pesquisar NFC-e/Cupom Fiscal Inutilizados"
      lblNumeroNFeInicial(0).Caption = "NFCe Número Inicial"
      lblNumeroNFeFinal(1).Caption = "NFCe Número Final"
      cmd_inutilizarNFe.Caption = "Inutilizar NFCe/Cupom Fiscal"
  End If

  With grdInutilizadas
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With

End Sub
