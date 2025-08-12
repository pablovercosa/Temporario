VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmInformacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Informações"
   ClientHeight    =   6360
   ClientLeft      =   270
   ClientTop       =   480
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Informacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   14910
   Begin TabDlg.SSTab sstInfo 
      Height          =   5700
      Left            =   60
      TabIndex        =   2
      Top             =   585
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   8
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
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "Informacoes.frx":4E95A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Comentários2"
      Tab(0).Control(1)=   "lblFax"
      Tab(0).Control(2)=   "lblFone2"
      Tab(0).Control(3)=   "lblFone1"
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(5)=   "Label21"
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(7)=   "Fantasia2"
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(9)=   "Tipo2"
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(11)=   "Bloqueado2"
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(13)=   "Inativo2"
      Tab(0).Control(14)=   "Label6"
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(16)=   "Cidade2"
      Tab(0).Control(17)=   "Label7"
      Tab(0).Control(18)=   "Estado2"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Compras/Vendas"
      TabPicture(1)   =   "Informacoes.frx":4E976
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdApagarRegistros"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "B1_Remonta2"
      Tab(1).Control(3)=   "Data1"
      Tab(1).Control(4)=   "Grade12"
      Tab(1).Control(5)=   "Label39"
      Tab(1).Control(6)=   "Label38"
      Tab(1).Control(7)=   "Line9"
      Tab(1).Control(8)=   "Line8"
      Tab(1).Control(9)=   "Line7"
      Tab(1).Control(10)=   "Label37"
      Tab(1).Control(11)=   "lblProdutoCor"
      Tab(1).Control(12)=   "lblProdutoTamanho"
      Tab(1).Control(13)=   "Tot_Unid12"
      Tab(1).Control(14)=   "Tot_Unid22"
      Tab(1).Control(15)=   "Tot_Val12"
      Tab(1).Control(16)=   "Tot_Val22"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Pagar"
      TabPicture(2)   =   "Informacoes.frx":4E992
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Shape3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Line4"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Line5"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Line6"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Grade22"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "B_Monta_Pagar"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame2"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Data2"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Frame5"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Receber"
      TabPicture(3)   =   "Informacoes.frx":4E9AE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Shape6"
      Tab(3).Control(1)=   "Shape5"
      Tab(3).Control(2)=   "Shape4"
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(4)=   "Label12"
      Tab(3).Control(5)=   "Label11"
      Tab(3).Control(6)=   "Line1"
      Tab(3).Control(7)=   "Line2"
      Tab(3).Control(8)=   "Line3"
      Tab(3).Control(9)=   "Grade32"
      Tab(3).Control(10)=   "Frame3"
      Tab(3).Control(11)=   "B_Monta_Receber"
      Tab(3).Control(12)=   "Data3"
      Tab(3).Control(13)=   "Frame6"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Cheques e Cartões"
      TabPicture(4)   =   "Informacoes.frx":4E9CA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grade_Cartões2"
      Tab(4).Control(1)=   "Grade_Cheques2"
      Tab(4).Control(2)=   "Data_Cartões"
      Tab(4).Control(3)=   "Data_Cheques"
      Tab(4).Control(4)=   "pnlTotalChequesBand"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Conta do Cliente"
      TabPicture(5)   =   "Informacoes.frx":4E9E6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grade42"
      Tab(5).Control(1)=   "Data4"
      Tab(5).Control(2)=   "B_Monta_Conta"
      Tab(5).Control(3)=   "Frame4"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Outras Informações"
      TabPicture(6)   =   "Informacoes.frx":4EA02
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Grade_Efetuados2"
      Tab(6).Control(1)=   "Grade_Contatos2"
      Tab(6).Control(2)=   "Data_Efetuados"
      Tab(6).Control(3)=   "Data_Contatos"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Serviços"
      TabPicture(7)   =   "Informacoes.frx":4EA1E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Grade_Serv"
      Tab(7).Control(1)=   "Data5"
      Tab(7).ControlCount=   2
      Begin Threed.SSPanel pnlTotalChequesBand 
         Height          =   315
         Left            =   -74880
         TabIndex        =   91
         Top             =   5310
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.Frame Frame8 
            Caption         =   "Resumo dos cartões"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   5400
            TabIndex        =   106
            Top             =   360
            Width           =   5295
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Vencido e Não Recebido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3360
               TabIndex        =   120
               Top             =   420
               Width           =   1800
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido Atrasado :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   119
               Top             =   2220
               Width           =   1860
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido em Dia :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   118
               Top             =   1860
               Width           =   1725
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "Total a Receber :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   117
               Top             =   1500
               Width           =   1245
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   116
               Top             =   1140
               Width           =   1185
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Qtde Total de Contas :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   420
               Width           =   1605
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total das Contas :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   114
               Top             =   780
               Width           =   1695
            End
            Begin VB.Label lblQtdeTotaldeContasCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   113
               Top             =   360
               Width           =   1380
            End
            Begin VB.Label lblValorTotaldasContasCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   112
               Top             =   720
               Width           =   1380
            End
            Begin VB.Label lblTotalRecebidoCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   111
               Top             =   1080
               Width           =   1380
            End
            Begin VB.Label lblTotalaReceberCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   110
               Top             =   1440
               Width           =   1380
            End
            Begin VB.Label lblTotalRecebidoemDiaCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1950
               TabIndex        =   109
               Top             =   1800
               Width           =   1380
            End
            Begin VB.Label lblTotalRecebidoAtrasadoCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1950
               TabIndex        =   108
               Top             =   2160
               Width           =   1380
            End
            Begin VB.Label lblVencidoeNaoRecebidoCA 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3600
               TabIndex        =   107
               Top             =   720
               Width           =   1380
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Resumo dos cheques"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   0
            TabIndex        =   93
            Top             =   360
            Width           =   5295
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Vencido e Não Recebido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3360
               TabIndex        =   122
               Top             =   420
               Width           =   1800
            End
            Begin VB.Label lblQtdeTotaldeContasCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   121
               Top             =   360
               Width           =   1380
            End
            Begin VB.Label lblVencidoeNaoRecebidoCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3600
               TabIndex        =   105
               Top             =   720
               Width           =   1380
            End
            Begin VB.Label lblTotalRecebidoAtrasadoCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1950
               TabIndex        =   104
               Top             =   2160
               Width           =   1380
            End
            Begin VB.Label lblTotalRecebidoemDiaCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1950
               TabIndex        =   103
               Top             =   1800
               Width           =   1380
            End
            Begin VB.Label lblTotalaReceberCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   102
               Top             =   1440
               Width           =   1380
            End
            Begin VB.Label lblTotalRecebidoCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   101
               Top             =   1080
               Width           =   1380
            End
            Begin VB.Label lblValorTotaldasContasCH 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   100
               Top             =   720
               Width           =   1380
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total das Contas :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   99
               Top             =   780
               Width           =   1695
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Qtde Total de Contas :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   98
               Top             =   420
               Width           =   1605
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   97
               Top             =   1140
               Width           =   1185
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Total a Receber :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   96
               Top             =   1500
               Width           =   1245
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido em Dia :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   95
               Top             =   1860
               Width           =   1725
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido Atrasado :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   94
               Top             =   2220
               Width           =   1860
            End
         End
         Begin Threed.SSPanel sspTtileTotlizadores 
            Height          =   285
            Left            =   240
            TabIndex        =   92
            Top             =   0
            Width           =   10215
            _Version        =   65536
            _ExtentX        =   18018
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Totalizadores"
            ForeColor       =   16777215
            BackColor       =   13604704
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin VB.Image Image3 
            Height          =   285
            Left            =   10440
            MouseIcon       =   "Informacoes.frx":4EA3A
            MousePointer    =   99  'Custom
            Picture         =   "Informacoes.frx":4F304
            Top             =   0
            Width           =   360
         End
         Begin VB.Image Image2 
            Height          =   285
            Left            =   10440
            MouseIcon       =   "Informacoes.frx":4F89E
            MousePointer    =   99  'Custom
            Picture         =   "Informacoes.frx":50168
            Top             =   0
            Width           =   360
         End
         Begin VB.Image Image1 
            Height          =   285
            Left            =   0
            Picture         =   "Informacoes.frx":50702
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdApagarRegistros 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Apagar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -62430
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -67080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin SSDataWidgets_B.SSDBGrid Grade_Serv 
         Bindings        =   "Informacoes.frx":50B20
         Height          =   4950
         Left            =   -74790
         TabIndex        =   84
         Top             =   525
         Width           =   14280
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
         ExtraHeight     =   53
         Columns(0).Width=   3200
         _ExtentX        =   25188
         _ExtentY        =   8731
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Frame Frame6 
         Caption         =   "Resumo"
         Height          =   1710
         Left            =   -72555
         TabIndex        =   64
         Top             =   390
         Width           =   7170
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Vencido e Não Recebido"
            Height          =   195
            Left            =   3720
            TabIndex        =   83
            Top             =   660
            Width           =   1725
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Total Recebido Atrasado"
            Height          =   195
            Left            =   3720
            TabIndex        =   80
            Top             =   1020
            Width           =   1770
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Total Recebido em Dia"
            Height          =   195
            Left            =   3720
            TabIndex        =   79
            Top             =   300
            Width           =   1590
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Total a Receber"
            Height          =   195
            Left            =   240
            TabIndex        =   77
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total Recebido"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   1020
            Width           =   1065
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Total de Contas"
            Height          =   195
            Left            =   240
            TabIndex        =   73
            Top             =   1380
            Width           =   1545
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total das Contas"
            Height          =   195
            Left            =   240
            TabIndex        =   72
            Top             =   660
            Width           =   1620
         End
         Begin VB.Label Qtde_Total_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   71
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Valor_Total_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   70
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label Total_Pago_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   69
            Top             =   967
            Width           =   1380
         End
         Begin VB.Label Total_Pagar_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   68
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Total_Dia_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5670
            TabIndex        =   67
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Total_Atrasado_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   5670
            TabIndex        =   66
            Top             =   967
            Width           =   1380
         End
         Begin VB.Label Total_Vencido_Receber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   5670
            TabIndex        =   65
            Top             =   600
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Resumo"
         Height          =   1710
         Left            =   2445
         TabIndex        =   54
         Top             =   390
         Width           =   7170
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Total Vencido e Não Pago"
            Height          =   195
            Left            =   3570
            TabIndex        =   82
            Top             =   660
            Width           =   1830
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Total Pago Atrasado"
            Height          =   195
            Left            =   3570
            TabIndex        =   81
            Top             =   1020
            Width           =   1470
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Total Pago em Dia"
            Height          =   195
            Left            =   3570
            TabIndex        =   78
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Total a Pagar"
            Height          =   195
            Left            =   240
            TabIndex        =   76
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Total Pago"
            Height          =   195
            Left            =   240
            TabIndex        =   75
            Top             =   1020
            Width           =   765
         End
         Begin VB.Label Total_Vencido_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5670
            TabIndex        =   63
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label Total_Atrasado_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   5670
            TabIndex        =   62
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Total_Dia_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5670
            TabIndex        =   61
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Total_Pagar_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   60
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Total_Pago_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   59
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Valor_Total_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   58
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label Qtde_Total_Pagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E5E5E5&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   57
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total das Contas"
            Height          =   195
            Left            =   240
            TabIndex        =   56
            Top             =   660
            Width           =   1620
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Qtde Total de Contas"
            Height          =   195
            Left            =   240
            TabIndex        =   55
            Top             =   1380
            Width           =   1545
         End
      End
      Begin VB.Data Data_Contatos 
         Caption         =   "Data_Contatos"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -66960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data Data_Efetuados 
         Caption         =   "Data_Efetuados"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -66840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contas"
         Height          =   795
         Left            =   -74670
         TabIndex        =   47
         Top             =   780
         Width           =   10635
         Begin VB.OptionButton O_Conta_Todas 
            Appearance      =   0  'Flat
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   420
            TabIndex        =   50
            Top             =   330
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton O_Conta_Recebidas 
            Appearance      =   0  'Flat
            Caption         =   "Recebidas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1620
            TabIndex        =   49
            Top             =   330
            Width           =   1095
         End
         Begin VB.OptionButton O_Conta_Receber 
            Appearance      =   0  'Flat
            Caption         =   "A Receber"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3090
            TabIndex        =   48
            Top             =   330
            Width           =   1095
         End
      End
      Begin VB.CommandButton B_Monta_Conta 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -63930
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1710
         Width           =   2955
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data Data_Cheques 
         Caption         =   "Data_Cheques"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Height          =   300
         Left            =   -71355
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1095
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data_Cartões 
         Caption         =   "Data_Cartões"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -70320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -66360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton B_Monta_Receber 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Atualizar"
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
         Left            =   -63930
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2160
         Width           =   2985
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contas"
         Height          =   1725
         Left            =   -65280
         TabIndex        =   35
         Top             =   390
         Width           =   1245
         Begin VB.OptionButton O_Receber 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "A Receber"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   38
            Top             =   1170
            Width           =   1095
         End
         Begin VB.OptionButton O_Recebidas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Recebidas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   37
            Top             =   750
            Width           =   1095
         End
         Begin VB.OptionButton O_Receber_Recebidas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   36
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   8760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contas"
         Height          =   1725
         Left            =   9660
         TabIndex        =   27
         Top             =   390
         Width           =   1305
         Begin VB.OptionButton O_Pagas_Pagar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   330
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton O_Pagas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Pagas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   750
            Width           =   825
         End
         Begin VB.OptionButton O_Pagar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "A Pagar"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.CommandButton B_Monta_Pagar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Atualizar"
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2160
         Width           =   3165
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ordem"
         Height          =   855
         Left            =   -65040
         TabIndex        =   19
         Top             =   480
         Width           =   2505
         Begin VB.OptionButton O1_Data 
            Appearance      =   0  'Flat
            Caption         =   "Data"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   330
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton O1_Produto 
            Appearance      =   0  'Flat
            Caption         =   "Produto"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.CommandButton B1_Remonta2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -62430
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2220
         Width           =   2115
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
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
         Left            =   -66240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Comentários2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   3660
         Left            =   -73815
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1740
         Width           =   11280
      End
      Begin SSDataWidgets_B.SSDBGrid Grade_Contatos2 
         Bindings        =   "Informacoes.frx":50B34
         Height          =   1965
         Left            =   -74760
         TabIndex        =   53
         Top             =   600
         Width           =   14235
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         _ExtentX        =   25109
         _ExtentY        =   3466
         _StockProps     =   79
         Caption         =   "Contatos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SSDataWidgets_B.SSDBGrid Grade_Efetuados2 
         Bindings        =   "Informacoes.frx":50B50
         Height          =   2655
         Left            =   -74760
         TabIndex        =   52
         Top             =   2760
         Width           =   14235
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   318
         Columns(0).Width=   3200
         _ExtentX        =   25109
         _ExtentY        =   4683
         _StockProps     =   79
         Caption         =   "Contatos Efetuados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SSDataWidgets_B.SSDBGrid Grade42 
         Bindings        =   "Informacoes.frx":50B6D
         Height          =   3765
         Left            =   -74670
         TabIndex        =   51
         Top             =   1680
         Width           =   10635
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   18759
         _ExtentY        =   6641
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SSDataWidgets_B.SSDBGrid Grade_Cheques2 
         Bindings        =   "Informacoes.frx":50B81
         Height          =   2175
         Left            =   -74880
         TabIndex        =   45
         Top             =   480
         Width           =   14295
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         _ExtentX        =   25215
         _ExtentY        =   3836
         _StockProps     =   79
         Caption         =   "Cheques"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SSDataWidgets_B.SSDBGrid Grade_Cartões2 
         Bindings        =   "Informacoes.frx":50B9C
         Height          =   2415
         Left            =   -74880
         TabIndex        =   44
         Top             =   2760
         Width           =   14325
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         _ExtentX        =   25268
         _ExtentY        =   4260
         _StockProps     =   79
         Caption         =   "Cartões"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SSDataWidgets_B.SSDBGrid Grade32 
         Bindings        =   "Informacoes.frx":50BB7
         Height          =   3360
         Left            =   -74790
         TabIndex        =   43
         Top             =   2160
         Width           =   10755
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
         AllowUpdate     =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   18971
         _ExtentY        =   5927
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin SSDataWidgets_B.SSDBGrid Grade22 
         Bindings        =   "Informacoes.frx":50BCB
         Height          =   3360
         Left            =   210
         TabIndex        =   34
         Top             =   2160
         Width           =   10755
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   16777152
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   18971
         _ExtentY        =   5927
         _StockProps     =   79
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
      Begin SSDataWidgets_B.SSDBGrid Grade12 
         Bindings        =   "Informacoes.frx":50BDF
         Height          =   4095
         Left            =   -74850
         TabIndex        =   17
         Top             =   1440
         Width           =   12330
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
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         MaxSelectedRows =   0
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   53
         Columns(0).Width=   3200
         UseDefaults     =   0   'False
         _ExtentX        =   21749
         _ExtentY        =   7223
         _StockProps     =   79
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
      Begin VB.Label Label39 
         Caption         =   "Cor"
         Height          =   255
         Left            =   -69390
         TabIndex        =   127
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label38 
         Caption         =   "Tamanho"
         Height          =   255
         Left            =   -69780
         TabIndex        =   126
         Top             =   660
         Width           =   735
      End
      Begin VB.Line Line9 
         X1              =   -70830
         X2              =   -70830
         Y1              =   1140
         Y2              =   1380
      End
      Begin VB.Line Line8 
         X1              =   -70830
         X2              =   -70710
         Y1              =   1380
         Y2              =   1260
      End
      Begin VB.Line Line7 
         X1              =   -70830
         X2              =   -70950
         Y1              =   1380
         Y2              =   1260
      End
      Begin VB.Label Label37 
         Caption         =   "Clique na linha para ver Tamanho e Cor"
         Height          =   495
         Left            =   -71550
         TabIndex        =   125
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label lblProdutoCor 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -69000
         TabIndex        =   124
         Top             =   990
         Width           =   2955
      End
      Begin VB.Label lblProdutoTamanho 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -69000
         TabIndex        =   123
         Top             =   630
         Width           =   2955
      End
      Begin VB.Line Line6 
         X1              =   1260
         X2              =   1140
         Y1              =   2130
         Y2              =   2010
      End
      Begin VB.Line Line5 
         X1              =   1260
         X2              =   1380
         Y1              =   2130
         Y2              =   2010
      End
      Begin VB.Line Line4 
         X1              =   1260
         X2              =   1260
         Y1              =   1620
         Y2              =   2130
      End
      Begin VB.Line Line3 
         X1              =   -73740
         X2              =   -73860
         Y1              =   2130
         Y2              =   2010
      End
      Begin VB.Line Line2 
         X1              =   -73740
         X2              =   -73620
         Y1              =   2130
         Y2              =   2010
      End
      Begin VB.Line Line1 
         X1              =   -73740
         X2              =   -73740
         Y1              =   1620
         Y2              =   2130
      End
      Begin VB.Label lblFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -64350
         TabIndex        =   89
         Top             =   1245
         Width           =   1815
      End
      Begin VB.Label lblFone2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -64350
         TabIndex        =   88
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblFone1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -64350
         TabIndex        =   87
         Top             =   442
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "Telefones"
         Height          =   255
         Left            =   -65310
         TabIndex        =   86
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Fax"
         Height          =   255
         Left            =   -64890
         TabIndex        =   85
         Top             =   1290
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Recebido com atraso"
         Height          =   255
         Left            =   -74400
         TabIndex        =   42
         Top             =   1395
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Vencido e não recebido"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74400
         TabIndex        =   41
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Em dia ou a receber"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74400
         TabIndex        =   40
         Top             =   675
         Width           =   1575
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -74760
         Top             =   1395
         Width           =   255
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -74760
         Top             =   1050
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -74760
         Top             =   675
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000A&
         Caption         =   "Em dia ou a pagar"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   675
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Top             =   675
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Vencido e não pago"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   1050
         Width           =   1530
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Top             =   1395
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Pago com atraso"
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   1395
         Width           =   1395
      End
      Begin VB.Label Tot_Unid12 
         Caption         =   "Total Unidades"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   630
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Tot_Unid22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -73440
         TabIndex        =   24
         Top             =   592
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Tot_Val12 
         Caption         =   "Total Valores"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   990
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Tot_Val22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -73440
         TabIndex        =   22
         Top             =   952
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fantasia"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Fantasia2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -73800
         TabIndex        =   15
         Top             =   442
         Width           =   7965
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Tipo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -73800
         TabIndex        =   13
         Top             =   855
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Bloqueado"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Bloqueado2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -73800
         TabIndex        =   11
         Top             =   1245
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Inativo"
         Height          =   255
         Left            =   -71130
         TabIndex        =   10
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label Inativo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -70410
         TabIndex        =   9
         Top             =   1245
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Comentários"
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   -71130
         TabIndex        =   7
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Cidade2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -70410
         TabIndex        =   6
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Estado"
         Height          =   255
         Left            =   -67080
         TabIndex        =   5
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Estado2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -66330
         TabIndex        =   4
         Top             =   855
         Width           =   495
      End
   End
   Begin VB.Label Código 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   13125
   End
End
Attribute VB_Name = "frmInformacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Dim rsTamanhos  As Recordset
  Dim rsCores     As Recordset

  Dim rsCliFor     As Recordset
  Dim rsUsuarios   As Recordset
  Dim rsParametros As Recordset
  
  Dim Compras As Integer
  Dim Pagar   As Integer
  Dim Receber As Integer
  Dim Cheques As Integer
  Dim Cartões As Integer
  Dim Outras  As Integer
  Dim Conta_Cliente As Integer
  Dim Serviços      As Boolean
  Dim gbDescrAdicional As Boolean
  
  Dim DBCNAB            As Database
  Dim rsDescCNAB        As Recordset
  Dim intCodRetornoCNAB As Integer
  Dim strTipoRetorno    As String     'Identifica se analisa o retorno A, B ou C
  
  '11/06/2004 - Daniel
  'Variáveis para o tratamento de totalizadores de
  'Cheques e
  Dim m_QtdeTotaldeContasCH     As Integer
  Dim m_ValorTotaldasContasCH   As Double
  Dim m_TotalRecebidoCH         As Double
  Dim m_TotalaReceberCH         As Double
  Dim m_TotalRecebidoemDiaCH    As Double
  Dim m_TotalRecebidoAtrasadoCH As Double
  Dim m_VencidoeNaoRecebidoCH   As Double
  'Cartões
  Dim m_QtdeTotaldeContasCA     As Integer
  Dim m_ValorTotaldasContasCA   As Double
  Dim m_TotalRecebidoCA         As Double
  Dim m_TotalaReceberCA         As Double
  Dim m_TotalRecebidoemDiaCA    As Double
  Dim m_TotalRecebidoAtrasadoCA As Double
  Dim m_VencidoeNaoRecebidoCA   As Double
  
Private Sub cmdApagarRegistros_Click()
  frmApagaInformacoes.Show vbModal
End Sub

Private Sub Grade12_Click()
On Error GoTo Erro

  Dim bkmrk As Variant ' Bookmarks are always defined as variants
  Dim lTamanho As Long
  Dim lCor As Long
  Dim lContador As Long
  
  lTamanho = Grade12.Columns("Tamanho").CellValue(bkmrk)
  lCor = Grade12.Columns("Cor").CellValue(bkmrk)
  
  lContador = 0
  If rsTamanhos.RecordCount > 0 Then
    rsTamanhos.MoveFirst
    
    For lContador = 0 To rsTamanhos.RecordCount - 1
        If lTamanho = rsTamanhos.Fields("Código") Then
          lblProdutoTamanho.Caption = rsTamanhos.Fields("Nome")
          Exit For
        End If
        rsTamanhos.MoveNext
    Next
  End If
  
  lContador = 0
  If rsCores.RecordCount > 0 Then
    rsCores.MoveFirst
    
    For lContador = 0 To rsCores.RecordCount - 1
        If lCor = rsCores.Fields("Código") Then
          lblProdutoCor.Caption = rsCores.Fields("Nome")
          Exit For
        End If
        rsCores.MoveNext
    Next
  End If
  
  Exit Sub
Erro:
    MsgBox "Erro na visualização da Grade do produto selecionado " & Err.Number & " " & Err.Description, vbInformation, "Grade"

End Sub

'------------------------------------------------------------------------------
'28/10/2005 - mpdea
'Corrigido sobreposição de objetos e modificado layout para que não
'oculte os dados de cartões por padrão
'------------------------------------------------------------------------------
Private Sub Image2_Click()
  pnlTotalChequesBand.Height = 315
  pnlTotalChequesBand.Top = 5310
  Grade_Cartões2.Visible = True
  
  Image2.Visible = False
  Image3.Visible = True
  
  MontaTotalizadores
End Sub

Private Sub Image3_Click()
  Grade_Cartões2.Visible = False
  pnlTotalChequesBand.Top = 2670
  pnlTotalChequesBand.Height = 2955
  
  Image2.Visible = True
  Image3.Visible = False
End Sub
'------------------------------------------------------------------------------

Private Sub sstInfo_Click(PreviousTab As Integer)
  Select Case sstInfo.Tab
    Case 1
      If Compras = False Then
        Monta_Grade_Compras
      End If
    Case 2
      If Pagar = False Then
        Monta_Grade_Pagar
      End If
    Case 3
      If Receber = False Then
        Monta_Grade_Receber
      End If
    Case 4
    '28/10/2005 - mpdea
    'Comentado código para permanecer visível as informações de cartões
'      '11/06/2004 - Daniel
'      'Para carregar enrolado a panel de totalizadores
'      Image3_Click
'      '------------------------------------------------
      
      If Cheques = False Then
        Monta_Grade_Cheques
      End If
      If Cartões = False Then
        Monta_Grade_Cartões
      End If
    Case 5
      If Conta_Cliente = False Then
        Monta_Grade_Conta
      End If
    Case 6
      If Outras = False Then
        Monta_Grade_Contatos
        Monta_Grade_Efetuados
      End If
    Case 7
      If Serviços = False Then
        Monta_Grade_Serviços
      End If
  End Select
End Sub

Sub Monta_Grade_Serviços()
  Dim Rec_Serviços As Recordset
  Dim sSql
  
  sSql = "SELECT Contador, Sequência, Data, Serviço, Descrição, Tempo, Valor FROM [Comissão Serviços]"
  'sSql = sSql + " INNER JOIN Produtos On [Resumo Clientes].Produto = Produtos.Código"
  sSql = sSql + " WHERE Cliente = " + Código.Caption
  'If O1_Data.Value = True Then sSql = sSql + " ORDER By Sequência, Contador"
  'If O1_Produto.Value = True Then sSql = sSql + " ORDER By Produtos.Nome, Dia"
  
  Set Rec_Serviços = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade_Serv.DataMode = 1

  Set Data5.Recordset = Rec_Serviços

  Grade_Serv.Visible = False
  
  Grade_Serv.DataMode = 0

  Grade_Serv.ReBind
  
  Grade_Serv.Columns(0).Visible = False
  
  Grade_Serv.Columns(1).Width = 1000
  Grade_Serv.Columns(2).Width = 1100
  Grade_Serv.Columns(3).Width = 900
  Grade_Serv.Columns(4).Width = 5200
  Grade_Serv.Columns(5).Width = 1000
  Grade_Serv.Columns(6).Width = 1000
  Grade_Serv.Columns(6).NumberFormat = "###,###,##0.00"
  
  Serviços = True

  Grade_Serv.Visible = True

End Sub

Private Sub B_Monta_Conta_Click()
 Monta_Grade_Conta
End Sub

Private Sub B_Monta_Pagar_Click()
 Monta_Grade_Pagar
End Sub

Private Sub B_Monta_Receber_Click()
  Monta_Grade_Receber
End Sub

Private Sub B1_Remonta2_Click()
 Monta_Grade_Compras
End Sub

Private Sub Form_Activate()

  Call RefreshForm
  
  sstInfo.Tab = 3

End Sub

Private Sub RefreshForm()
  
  Compras = False
  Pagar = False
  Receber = False
  Cheques = False
  Cartões = False
  Outras = False
  Conta_Cliente = False
  Serviços = False
  
  Screen.MousePointer = vbHourglass
 
  Código.Caption = gsCodCliente
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(Código.Caption)
  If rsCliFor.NoMatch Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  rsUsuarios.Index = "Código"
  rsUsuarios.Seek "=", gnUserCode
  If rsUsuarios.NoMatch Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  sstInfo.TabEnabled(1) = True
  sstInfo.TabEnabled(2) = True
  sstInfo.TabEnabled(3) = True
  sstInfo.TabEnabled(4) = True
  sstInfo.TabEnabled(5) = True
  sstInfo.TabEnabled(6) = True
  sstInfo.TabEnabled(7) = True
  
  gbDescrAdicional = False
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
     Screen.MousePointer = vbDefault
     Exit Sub
  End If
  gbDescrAdicional = rsParametros("Saida Descr Adicional")
  
  If rsUsuarios("Pasta Compras") = False Then sstInfo.TabEnabled(1) = False
  If rsUsuarios("Pasta Pagar") = False Then sstInfo.TabEnabled(2) = False
  If rsUsuarios("Pasta Receber") = False Then sstInfo.TabEnabled(3) = False
  If rsUsuarios("Pasta Cheques") = False Then sstInfo.TabEnabled(4) = False
  If rsUsuarios("Pasta Conta") = False Then sstInfo.TabEnabled(5) = False
  If rsUsuarios("Pasta Outras") = False Then sstInfo.TabEnabled(6) = False
  If rsUsuarios("Pasta Serviços") = False Then sstInfo.TabEnabled(7) = False
  
  Nome.Caption = rsCliFor("Nome") & ""
  Fantasia2.Caption = rsCliFor("Fantasia") & ""
  If rsCliFor("Tipo") = "C" Then Tipo2.Caption = "Cliente"
  If rsCliFor("Tipo") = "F" Then Tipo2.Caption = "Fornecedor"
  If rsCliFor("Tipo") = "R" Then Tipo2.Caption = "Revendedor"
  If rsCliFor("Tipo") = "O" Then Tipo2.Caption = "Outros"
  
  If rsCliFor("Bloqueado") = True Then Bloqueado2.Caption = "Sim"
  If rsCliFor("Bloqueado") = False Then Bloqueado2.Caption = "Não"
  If rsCliFor("Inativo") = True Then Inativo2.Caption = "Sim"
  If rsCliFor("Inativo") = False Then Inativo2.Caption = "Não"
  
  Cidade2.Caption = rsCliFor("Cidade") & ""
  Estado2.Caption = rsCliFor("Estado") & ""
  
  lblFone1.Caption = rsCliFor("Fone 1") & ""
  lblFone2.Caption = rsCliFor("Fone 2") & ""
  lblFax.Caption = rsCliFor("Fax") & ""
  
  Comentários2.Text = rsCliFor("Comentários") & ""
  
  If gbServico = False Then sstInfo.TabVisible(7) = False
  
  Select Case sstInfo.Tab
    Case 1
      If Compras = False Then
        Monta_Grade_Compras
      End If
    Case 2
      If Pagar = False Then
        Monta_Grade_Pagar
      End If
    Case 3
      If Receber = False Then
        Monta_Grade_Receber
      End If
    Case 4
      If Cheques = False Then
        Monta_Grade_Cheques
      End If
      If Cartões = False Then
        Monta_Grade_Cartões
      End If
    Case 5
      If Conta_Cliente = False Then
        Monta_Grade_Conta
      End If
    Case 6
      If Outras = False Then
        Monta_Grade_Contatos
        Monta_Grade_Efetuados
      End If
    Case 7
      If Serviços = False Then
        Monta_Grade_Serviços
      End If
  End Select
  
'  If Compras = False Then
'    Monta_Grade_Compras
'  End If
'
'  If Pagar = False Then
'    Monta_Grade_Pagar
'  End If
'
'  If Receber = False Then
'    Monta_Grade_Receber
'  End If
'
'  If Cheques = False Then
'    Monta_Grade_Cheques
'  End If
'  If Cartões = False Then
'    Monta_Grade_Cartões
'  End If
'
'  If Conta_Cliente = False Then
'    Monta_Grade_Conta
'  End If
'
'  If Outras = False Then
'    Monta_Grade_Contatos
'    Monta_Grade_Efetuados
'  End If
'
'  If Serviços = False Then
'    Monta_Grade_Serviços
'  End If
  
  Screen.MousePointer = vbDefault

End Sub

Sub Monta_Grade_Cartões()
  Dim Rec_Cartões As Recordset
  Dim sSql
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Cliente, Cartões.Nome, Cartão, Vencimento, [Valor Cartão], Valor, Processado, Devolvido, Sequência, Acréscimo FROM [Contas a Receber]"
  sSql = sSql + " INNER JOIN Cartões On [Contas a Receber].Administradora = Cartões.Código"
  sSql = sSql + " WHERE Cliente = " + Código.Caption + " AND [Contas a Receber].Tipo = '" + "O" + "'"
  sSql = sSql + " ORDER By Vencimento"
  
  Set Rec_Cartões = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade_Cartões2.DataMode = 1

  Set Data_Cartões.Recordset = Rec_Cartões

  Grade_Cartões2.Visible = False
  
  Grade_Cartões2.DataMode = 0

  Grade_Cartões2.ReBind
  Grade_Cartões2.Columns(0).Visible = False
'  Grade_Cheques.Columns(1).Width = 590
  Grade_Cartões2.Columns(2).Width = 1100
'  Grade_Cheques.Columns(3).Width = 1350
  Grade_Cartões2.Columns(3).Width = 1100
''  Grade_Cartões2.Columns(4).NumberFormat = "###,###,##0.00"
  Grade_Cartões2.Columns(4).Width = 1000
''  Grade_Cartões2.Columns(5).NumberFormat = "###,###,##0.00"
  Grade_Cartões2.Columns(5).Width = 1200
  Grade_Cartões2.Columns(6).Width = 1100
'  Grade_Cheques.Columns(5).Width = 1000
'  Grade_Cheques.Columns(6).Width = 1000
  Grade_Cartões2.Columns(6).Style = 2
'  Grade_Cheques.Columns(7).Style = 2
  Cartões = True

  Grade_Cartões2.Visible = True

  '11/06/2004 - Daniel
  'Tratamento para Totalizadores
  Call LimparVarsTotalizadoresCA
  
  If Rec_Cartões.RecordCount = 0 Then Exit Sub
  
  With Rec_Cartões
    .MoveLast
    .MoveFirst
    
    m_QtdeTotaldeContasCA = .RecordCount
  
    Do Until .EOF
      
      m_ValorTotaldasContasCA = m_ValorTotaldasContasCA + .Fields("Valor").Value
      
      If .Fields("Processado").Value = True Then 'Recebeu
        m_TotalRecebidoCA = m_TotalRecebidoCA + .Fields("Valor").Value
        
        If .Fields("Acréscimo").Value = 0 Then 'Em dia
          m_TotalRecebidoemDiaCA = m_TotalRecebidoemDiaCA + .Fields("Valor").Value
        Else 'Com Atraso
          m_TotalRecebidoAtrasadoCA = m_TotalRecebidoAtrasadoCA + .Fields("Valor").Value
        End If
      
      Else
        If .Fields("Devolvido") = False Then
          m_TotalaReceberCA = m_TotalaReceberCA + .Fields("Valor").Value
        End If
      End If
    
      If .Fields("Vencimento").Value <= Data_Atual And .Fields("Processado").Value = False Then
        m_VencidoeNaoRecebidoCA = m_VencidoeNaoRecebidoCA + .Fields("Valor").Value
      End If
    
      .MoveNext
    Loop
  
  End With
  
End Sub

Sub Monta_Grade_Cheques()
  Dim Rec_Cheques As Recordset
  Dim sSql
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Cliente, Banco, Bancos.Nome, Cheque, Vencimento, Valor, Processado, Devolvido, Sequência, Acréscimo FROM [Contas a Receber]"
  sSql = sSql + " INNER JOIN Bancos On [Contas a Receber].Banco = Bancos.Código"
  sSql = sSql + " WHERE Cliente = " + Código.Caption + " AND [Contas a Receber].Tipo = '" + "C" + "'"
  sSql = sSql + " ORDER By Vencimento"
  
  Set Rec_Cheques = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade_Cheques2.DataMode = 1

  Set Data_Cheques.Recordset = Rec_Cheques

  Grade_Cheques2.Visible = False
  
  Grade_Cheques2.DataMode = 0

  Grade_Cheques2.ReBind
  Grade_Cheques2.Columns(0).Visible = False
  Grade_Cheques2.Columns(1).Width = 590
  Grade_Cheques2.Columns(2).Width = 1350
  Grade_Cheques2.Columns(3).Width = 1350
  Grade_Cheques2.Columns(4).Width = 1200
  'Grade_Cheques2.Columns(5).NumberFormat = "###,###,##0.00"
  Grade_Cheques2.Columns(5).Width = 1200
  Grade_Cheques2.Columns(6).Width = 1000
  Grade_Cheques2.Columns(6).Style = 2
  Grade_Cheques2.Columns(7).Style = 2
  Cheques = True

  Grade_Cheques2.Visible = True

  '11/06/2004 - Daniel
  'Tratamento para totalizadores
  
  Call LimparVarsTotalizadoresCH
  
  If Rec_Cheques.RecordCount = 0 Then Exit Sub
  
  With Rec_Cheques
    .MoveLast
    .MoveFirst
  
    m_QtdeTotaldeContasCH = .RecordCount
  
    Do Until .EOF
      
      m_ValorTotaldasContasCH = m_ValorTotaldasContasCH + .Fields("Valor").Value
      
      If .Fields("Processado").Value = True Then 'Recebeu
        m_TotalRecebidoCH = m_TotalRecebidoCH + .Fields("Valor").Value
        
        If .Fields("Acréscimo").Value = 0 Then 'Em dia
          m_TotalRecebidoemDiaCH = m_TotalRecebidoemDiaCH + .Fields("Valor").Value
        Else 'Com Atraso
          m_TotalRecebidoAtrasadoCH = m_TotalRecebidoAtrasadoCH + .Fields("Valor").Value
        End If
      
      Else
        If .Fields("Devolvido") = False Then
          m_TotalaReceberCH = m_TotalaReceberCH + .Fields("Valor").Value
        End If
      End If
    
      If .Fields("Vencimento").Value <= Data_Atual And .Fields("Processado").Value = False Then
        m_VencidoeNaoRecebidoCH = m_VencidoeNaoRecebidoCH + .Fields("Valor").Value
      End If
    
      .MoveNext
    Loop
  
  End With
  
End Sub

Sub Monta_Grade_Compras()
  Dim Rec_Compras As Recordset
  Dim sSql As String
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Cliente, Dia, Produto, Tamanho, Cor, Produtos.Nome, Qtde, [Valor Total],[Descricao Adicional] FROM [Resumo Clientes]"
  sSql = sSql + " INNER JOIN Produtos On [Resumo Clientes].Produto = Produtos.Código "
  sSql = sSql + " WHERE Cliente = " + Código.Caption
  If O1_Data.Value = True Then sSql = sSql + " ORDER By Dia"
  If O1_Produto.Value = True Then sSql = sSql + " ORDER By Produtos.Nome, Dia"
  
  Set Rec_Compras = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade12.DataMode = 1

  Set Data1.Recordset = Rec_Compras

  Grade12.Visible = False
  
  Grade12.DataMode = 0

  Grade12.ReBind
  Grade12.Columns(0).Visible = False
  ''Grade12.Columns(5).NumberFormat = FORMAT_VALUE
  Grade12.Columns(1).Width = 1000
  Grade12.Columns(2).Width = 1900
  
  Grade12.Columns(3).Width = 800  'tamanho
  Grade12.Columns(4).Width = 800  'cor
  
  Grade12.Columns(5).Width = 4800
  Grade12.Columns(6).Width = 750
  ''Grade12.Columns(6).NumberFormat = FORMAT_VALUE
  Grade12.Columns(7).Width = 1100
  If gbDescrAdicional = True Then
      Grade12.Columns(8).Width = 10000
      Grade12.Columns(8).Caption = "Adicional"
      Grade12.Columns(8).Visible = True
  Else
      Grade12.Columns(8).Visible = False
  End If
      
  Compras = True

  Grade12.Visible = True
End Sub

Sub Monta_Grade_Conta()
  Dim Rec_Conta As Recordset
  Dim sSql As String
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Filial, Cliente, Data, Produto, Descrição, Qtde, Valor, [Valor Pago], [Data Pagamento] FROM [Conta Cliente]"
  'sSql = sSql + " INNER JOIN Produtos On [Conta Cliente].Produto = Produtos.Código"
  sSql = sSql + " WHERE Cliente = " + Código.Caption
  If O_Conta_Recebidas = True Then sSql = sSql + " AND Valor = [Valor Pago]"
  If O_Conta_Receber = True Then sSql = sSql + " AND Valor <> [Valor Pago]"
  sSql = sSql + " ORDER By Data"
'  If O1_Data.Value = True Then sSql = sSql + " ORDER By Dia"
'  If O1_Produto.Value = True Then sSql = sSql + " ORDER By Produtos.Nome"
  
  Set Rec_Conta = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade42.DataMode = 1

  Set Data4.Recordset = Rec_Conta

  Grade42.Visible = False

  Grade42.DataMode = 0

  Grade42.ReBind
  Grade42.Columns(0).Visible = False
  Grade42.Columns(1).Visible = False
  Grade42.Columns(2).Width = 1200
  Grade42.Columns(4).Width = 3200
  Grade42.Columns(5).Width = 500
  'Grade42.Columns(6).NumberFormat = "###,###,##0.00"
  'Grade42.Columns(7).NumberFormat = "###,###,##0.00"
  Grade42.Columns(7).Width = 1000
  Grade42.Columns(8).Width = 1350

'  Grade1.Columns(1).Width = 1000
'  Grade1.Columns(2).Width = 1200
'  Grade1.Columns(3).Width = 3300
'  Grade1.Columns(4).Width = 750
'  Grade1.Columns(5).Width = 1100
  Grade42.Visible = True
  Conta_Cliente = True

End Sub

Sub Monta_Grade_Contatos()
  Dim Rec_Contatos As Recordset
  Dim sSql
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Cliente, Contato, Cargo, [Dia Aniversário], [Mês Aniversário], Ramal, Email FROM [Contatos]"
  sSql = sSql + " WHERE Cliente = " + Código.Caption
  
  Set Rec_Contatos = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade_Contatos2.DataMode = 1

  Set Data_Contatos.Recordset = Rec_Contatos

  Grade_Contatos2.Visible = False
  
  Grade_Contatos2.DataMode = 0

  Grade_Contatos2.ReBind
  Grade_Contatos2.Columns(0).Visible = False
  Grade_Contatos2.Columns(1).Width = 2290
  Grade_Contatos2.Columns(2).Width = 2150
  Grade_Contatos2.Columns(3).Width = 1400
  Grade_Contatos2.Columns(4).Width = 1400
  Grade_Contatos2.Columns(5).Width = 1400
  Grade_Contatos2.Columns(6).Width = 1400
  
'  Grade_Cheques.Columns(5).NumberFormat = "###,###,##0.00"
'  Grade_Cheques.Columns(5).Width = 1000
'  Grade_Cheques.Columns(6).Width = 1000
'   Grade_Cheques.Columns(6).Style = 2
'  Grade_Cheques.Columns(7).Style = 2
  Outras = True

  Grade_Contatos2.Visible = True

End Sub

Sub Monta_Grade_Efetuados()
  Dim Rec_Efetuados As Recordset
  Dim sSql
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Cliente, Data, Seqüência, Descrição, Pendência, [Data Aviso]  FROM [Contatos Efetuados]"
  sSql = sSql + " WHERE Cliente = " + Código.Caption
  sSql = sSql + " ORDER BY Seqüência"
  
  Set Rec_Efetuados = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade_Efetuados2.DataMode = 1

  Set Data_Efetuados.Recordset = Rec_Efetuados

  Grade_Efetuados2.Visible = False
  
  Grade_Efetuados2.DataMode = 0

  Grade_Efetuados2.ReBind
  Grade_Efetuados2.Columns(0).Visible = False
  Grade_Efetuados2.Columns(2).Visible = False
  Grade_Efetuados2.Columns(1).Width = 1000
  Grade_Efetuados2.Columns(3).Width = 7250
  Grade_Efetuados2.Columns(3).VertScrollBar = True
  Grade_Efetuados2.Columns(4).Style = 2
  Grade_Efetuados2.Columns(4).Width = 1000
  Grade_Efetuados2.Columns(5).Width = 1000
  
  Outras = True

  Grade_Efetuados2.Visible = True

End Sub

Sub Monta_Grade_Pagar()
  Dim Rec_Pagar As Recordset
  Dim sSql As String
  Dim Aux_Total_Contas As Integer
  Dim Aux_Valor_Contas As Double
  Dim Aux_Total_Pago As Double
  Dim Aux_Total_Pagar As Double
  Dim Aux_Pago_Dia As Double
  Dim Aux_Atrasado As Double
  Dim Aux_Vencido As Double
  
  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = "SELECT Filial, Fornecedor, Vencimento, Descrição, Valor, Desconto, Acréscimo, [Valor Pago], Pagamento, Nota, Sequência FROM [Contas a Pagar]"
'  sSql = sSql + " INNER JOIN Cli_For On [Contas a Pagar].Fornecedor = Cli_For.Código"
  sSql = sSql + " WHERE Fornecedor = " + Código.Caption
  sSql = sSql + " And Filial = " + str(gnCodFilial)
  If O_Pagas.Value = True Then sSql = sSql + " And Valor = [Valor Pago]"
  If O_Pagar.Value = True Then sSql = sSql + " And [Valor Pago] = 0"
  sSql = sSql + " ORDER By Vencimento"
'  If O1_Produto.Value = True Then sSql = sSql + " ORDER By Produtos.Nome"
  
  Set Rec_Pagar = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade22.DataMode = 1

  Set Data2.Recordset = Rec_Pagar
  
  Grade22.Visible = False
  
  Grade22.DataMode = 0

  Grade22.ReBind
  Grade22.Columns(0).Visible = False 'Filial
  Grade22.Columns(1).Visible = False 'Fornecedor
  Grade22.Columns(2).Width = 1300
  'Grade22.Columns(4).NumberFormat = "###,###,##0.00"
  Grade22.Columns(4).Width = 1000
  'Grade22.Columns(5).NumberFormat = "###,###,##0.00"
  Grade22.Columns(5).Width = 1000
  'Grade22.Columns(6).NumberFormat = "###,###,##0.00"
  Grade22.Columns(6).Width = 1000
  'Grade22.Columns(7).NumberFormat = "###,###,##0.00"
  Grade22.Columns(7).Width = 1000
  Grade22.Columns(8).Width = 1000
  Grade22.Columns(9).Width = 1400
  
  Pagar = True

  Grade22.Visible = True
  
  If Rec_Pagar.RecordCount = 0 Then Exit Sub
  
  Aux_Total_Contas = 0
  Aux_Valor_Contas = 0
  Aux_Total_Pago = 0
  Aux_Total_Pagar = 0
  Aux_Pago_Dia = 0
  Aux_Atrasado = 0
  Aux_Vencido = 0
  Qtde_Total_Pagar.Caption = ""
  Valor_Total_Pagar.Caption = ""
  Total_Pago_Pagar.Caption = ""
  Total_Pagar_Pagar.Caption = ""
  Total_Dia_Pagar.Caption = ""
  Total_Atrasado_Pagar.Caption = ""
  Total_Vencido_Pagar.Caption = ""
  
  
  Rec_Pagar.MoveFirst
  
  Do While Not Rec_Pagar.EOF
    Aux_Total_Contas = Aux_Total_Contas + 1
    If Rec_Pagar("Valor Pago") = 0 Then
      Aux_Valor_Contas = Aux_Valor_Contas + Rec_Pagar("Valor")
      Aux_Total_Pagar = Aux_Total_Pagar + Rec_Pagar("Valor")
      If CDate(Rec_Pagar("Vencimento")) < Date Then
        Aux_Vencido = Aux_Vencido + Rec_Pagar("Valor")
      End If
    End If
    If Rec_Pagar("Valor Pago") <> 0 Then
      Aux_Valor_Contas = Aux_Valor_Contas + Rec_Pagar("Valor Pago")
      Aux_Total_Pago = Aux_Total_Pago + Rec_Pagar("Valor Pago")
      If Rec_Pagar("Vencimento") >= Rec_Pagar("Pagamento") Then
        Aux_Pago_Dia = Aux_Pago_Dia + Rec_Pagar("Valor Pago")
      Else
        Aux_Atrasado = Aux_Atrasado + Rec_Pagar("Valor Pago")
      End If
      
    End If
    
    Rec_Pagar.MoveNext
  Loop
  
  Qtde_Total_Pagar.Caption = Aux_Total_Contas
  Valor_Total_Pagar.Caption = Format(Aux_Valor_Contas, "###,###,###,##0.00")
  Total_Pago_Pagar.Caption = Format(Aux_Total_Pago, "###,###,###,##0.00")
  Total_Pagar_Pagar.Caption = Format(Aux_Total_Pagar, "###,###,###,##0.00")
  Total_Dia_Pagar.Caption = Format(Aux_Pago_Dia, "###,###,###,##0.00")
  Total_Atrasado_Pagar.Caption = Format(Aux_Atrasado, "###,###,###,##0.00")
  Total_Vencido_Pagar.Caption = Format(Aux_Vencido, "###,###,###,##0.00")
  
End Sub

Sub Monta_Grade_Receber()
  Dim Rec_Receber As Recordset
  Dim sSql
  Dim Aux_Total_Contas As Integer
  Dim Aux_Valor_Contas As Double
  Dim Aux_Total_Pago As Double
  Dim Aux_Total_Pagar As Double
  Dim Aux_Pago_Dia As Double
  Dim Aux_Atrasado As Double
  Dim Aux_Vencido As Double

  If Len(Trim(Código.Caption)) = 0 Then
    Exit Sub
  End If
  
  sSql = " SELECT Filial, Cliente, Vencimento, Descrição, Valor, Desconto, Acréscimo, [Valor Recebido], [Data Recebimento], Nota, Sequência"
  
  '17/10/2007 - Anderson
  'Alteração realizada para considerar o caminho padrão do Quick Store
  If Len(Trim(Dir$(gsDefaultPath & "QuickCNAB.mdb"))) > 0 Then
    sSql = sSql & ", CNAB_NossoNumero, CNAB_CodMovRet, CNAB_CodIdComplementar "
  End If
  
  sSql = sSql & " FROM [Contas a Receber] "
  sSql = sSql + " WHERE Cliente = " + Código.Caption + " AND [Contas a Receber].Tipo = '" + "R" + "'"
  sSql = sSql + " And Filial = " + str(gnCodFilial)
  If O_Recebidas.Value = True Then sSql = sSql + " And Valor = [Valor Recebido]"
  If O_Receber.Value = True Then sSql = sSql + " And [Valor Recebido] = 0"
  sSql = sSql + " ORDER By Vencimento"
  
  Set Rec_Receber = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade32.DataMode = 1

  Set Data3.Recordset = Rec_Receber

  Grade32.Visible = False
  
  Grade32.DataMode = 0

  Grade32.ReBind
  Grade32.Columns(0).Visible = False 'Filial
  Grade32.Columns(1).Visible = False 'Fornecedor
  Grade32.Columns(2).Width = 1150
  'Grade32.Columns(4).NumberFormat = "###,###,###.##"
  Grade32.Columns(4).Width = 1000
  'Grade32.Columns(5).NumberFormat = "###,###,##0.00"
  Grade32.Columns(5).Width = 1000
  'Grade32.Columns(6).NumberFormat = "###,###,##0.00"
  Grade32.Columns(6).Width = 1000
  'Grade32.Columns(7).NumberFormat = "###,###,##0.00"
  Grade32.Columns(7).Width = 1250
    
  Grade32.Columns(8).Caption = "Recebimento"
  Grade32.Columns(8).Width = 1100
  
  Grade32.Columns(9).Width = 1000
  
  '17/10/2007 - Anderson
  'Alteração realizada para considerar o caminho padrão do Quick Store
  If Len(Trim(Dir$(gsDefaultPath & "QuickCNAB.mdb"))) > 0 Then
    Grade32.Columns(11).Caption = "Nosso Número"
    
    Grade32.Columns(12).Width = 3000
    Grade32.Columns(12).Caption = "Retorno CNAB"
    
    Grade32.Columns(13).Width = 3000
    Grade32.Columns(13).Caption = "Código Complementar"
  End If
  
  Receber = True

 If Rec_Receber.RecordCount = 0 Then
    Grade32.Visible = True
    Exit Sub
 End If
  
  Aux_Total_Contas = 0
  Aux_Valor_Contas = 0
  Aux_Total_Pago = 0
  Aux_Total_Pagar = 0
  Aux_Pago_Dia = 0
  Aux_Atrasado = 0
  Aux_Vencido = 0
  Qtde_Total_Pagar.Caption = ""
  Valor_Total_Pagar.Caption = ""
  Total_Pago_Pagar.Caption = ""
  Total_Pagar_Pagar.Caption = ""
  Total_Dia_Pagar.Caption = ""
  Total_Atrasado_Pagar.Caption = ""
  Total_Vencido_Pagar.Caption = ""
  
  Rec_Receber.MoveFirst
  
  Do While Not Rec_Receber.EOF
    Aux_Total_Contas = Aux_Total_Contas + 1
    If Rec_Receber("Valor Recebido") = 0 Then
      Aux_Valor_Contas = Aux_Valor_Contas + Rec_Receber("Valor")
      Aux_Total_Pagar = Aux_Total_Pagar + Rec_Receber("Valor")
      If CDate(Rec_Receber("Vencimento")) < Date Then
        Aux_Vencido = Aux_Vencido + Rec_Receber("Valor")
      End If
    End If
    If Rec_Receber("Valor Recebido") <> 0 Then
      Aux_Valor_Contas = Aux_Valor_Contas + Rec_Receber("Valor Recebido")
      Aux_Total_Pago = Aux_Total_Pago + Rec_Receber("Valor Recebido")
      If Rec_Receber("Vencimento") >= Rec_Receber("Data Recebimento") Then
        Aux_Pago_Dia = Aux_Pago_Dia + Rec_Receber("Valor Recebido")
      Else
        Aux_Atrasado = Aux_Atrasado + Rec_Receber("Valor Recebido")
      End If
      
    End If
    
    Rec_Receber.MoveNext
  Loop
  
  Qtde_Total_Receber.Caption = Aux_Total_Contas
  Valor_Total_Receber.Caption = Format(Aux_Valor_Contas, "###,###,###,##0.00")
  Total_Pago_Receber.Caption = Format(Aux_Total_Pago, "###,###,###,##0.00")
  Total_Pagar_Receber.Caption = Format(Aux_Total_Pagar, "###,###,###,##0.00")
  Total_Dia_Receber.Caption = Format(Aux_Pago_Dia, "###,###,###,##0.00")
  Total_Atrasado_Receber.Caption = Format(Aux_Atrasado, "###,###,###,##0.00")
  Total_Vencido_Receber.Caption = Format(Aux_Vencido, "###,###,###,##0.00")
  
  Rec_Receber.MoveFirst
  
  Grade32.Visible = True

  Rec_Receber.Close
  If Not Rec_Receber Is Nothing Then Set Rec_Receber = Nothing
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsUsuarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  '17/10/2007 - Anderson
  'Alteração realizada para considerar o caminho padrão do Quick Store
  If Len(Trim(Dir$(gsDefaultPath & "QuickCNAB.mdb"))) > 0 Then
    '17/10/2007 - Anderson
    'Alteração realizada para considerar o caminho padrão do Quick Store
    Set DBCNAB = ws.OpenDatabase(gsDefaultPath & "QuickCNAB.mdb", False, False, ";pwd=" & gsGetPValue())
  End If
  
  Grade22.StyleSets("Verde").ForeColor = RGB(0, 128, 0)
'  Grade22.StyleSets("Verde").Font.Bold = True
  Grade22.StyleSets("Vermelho").ForeColor = RGB(255, 0, 0)
 ' Grade22.StyleSets("Vermelho").Font.Bold = True
  Grade22.StyleSets("Preto").ForeColor = RGB(0, 0, 0)
'  Grade22.StyleSets("Preto").Font.Bold = True
  
  Grade32.StyleSets("Verde").ForeColor = RGB(0, 128, 0)
'  Grade32.StyleSets("Verde").Font.Bold = True
  Grade32.StyleSets("Vermelho").ForeColor = RGB(255, 0, 0)
'  Grade32.StyleSets("Vermelho").Font.Bold = True
  Grade32.StyleSets("Preto").ForeColor = RGB(0, 0, 0)
'  Grade32.StyleSets("Preto").Font.Bold = True

  Call RefreshForm
  
  If gbSuperUser Then
    cmdApagarRegistros.Visible = True
  End If
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rsCliFor = Nothing
  Set rsUsuarios = Nothing
  Set rsParametros = Nothing
  
  rsTamanhos.Close
  rsCores.Close
  Set rsTamanhos = Nothing
  Set rsCores = Nothing
End Sub

Private Sub Grade12_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

Dim TotalSelec As Integer
Dim i As Integer
Dim bkmrk As Variant ' Bookmarks are always defined as variants
Dim Unidades As Double
Dim Valor As Double
 
  Unidades = 0
  Valor = 0
 
  Tot_Unid12.Visible = False
  Tot_Unid22.Visible = False
  Tot_Val12.Visible = False
  Tot_Val22.Visible = False
 
 
  If SelType <> 2 Then Exit Sub

  TotalSelec = Grade12.SelBookmarks.Count
  If TotalSelec < 2 Then Exit Sub

  For i = 0 To (TotalSelec - 1)
    bkmrk = Grade12.SelBookmarks(i)
    Unidades = Unidades + Grade12.Columns("Qtde").CellValue(bkmrk)
    Valor = Valor + Grade12.Columns("Valor Total").CellValue(bkmrk)
  Next i

  Tot_Unid22.Caption = Unidades
  'Tot_Val22.Caption = Format(Valor, "###,###,###,##0.00")

  Tot_Unid12.Visible = True
  Tot_Unid22.Visible = True
  Tot_Val12.Visible = True
  Tot_Val22.Visible = True

End Sub

Private Sub Grade22_RowLoaded(ByVal Bookmark As Variant)
 Dim Vencimento As Date
 Dim Pago As Double
 Dim Pagamento As Variant
 
 If IsEmpty(Bookmark) Then Exit Sub

 On Error GoTo Deu_Erro
 
 Vencimento = Grade22.Columns(2).CellText(Bookmark)
 Pago = Grade22.Columns(7).CellText(Bookmark)
 Pagamento = Grade22.Columns(8).CellText(Bookmark)
  
  If Pago = 0 Then  'não foi pago
    If Vencimento < Date Then Grade22.Columns(2).CellStyleSet "Vermelho"
    If Vencimento > Date Then Grade22.Columns(2).CellStyleSet "Verde"
    Exit Sub
  End If
    
  If CDate(Pagamento) <= Vencimento Then Grade22.Columns(2).CellStyleSet "Verde"
  If CDate(Pagamento) > Vencimento Then Grade22.Columns(2).CellStyleSet "Preto"

  Exit Sub
  
Deu_Erro:
  Exit Sub
  
End Sub

Private Sub Grade32_RowLoaded(ByVal Bookmark As Variant)
 Dim Vencimento As Date
 Dim Pago As Double
 Dim Pagamento As Variant
 Dim Aux As Variant
 
 If IsEmpty(Bookmark) Then Exit Sub
 
 'Aux = Bookmark
 'Debug.Print Bookmark
 'Aux = Grade3.Columns(2).CellText(Bookmark)
  On Error GoTo Deu_Erro
 
 Vencimento = Grade32.Columns(2).CellText(Bookmark)
 Pago = Grade32.Columns(7).CellText(Bookmark)
 Pagamento = Grade32.Columns(8).CellText(Bookmark)
  
  If Pago = 0 Then  'não foi pago
    If CDate(Vencimento) < Date Then Grade32.Columns(2).CellStyleSet "Vermelho"
    If CDate(Vencimento) > Date Then Grade32.Columns(2).CellStyleSet "Verde"
    Exit Sub
  End If
    
  If CDate(Pagamento) <= Vencimento Then Grade32.Columns(2).CellStyleSet "Verde"
  If CDate(Pagamento) > Vencimento Then Grade32.Columns(2).CellStyleSet "Preto"
  
  ' 05/05/2003 - Maikel
  '              Linhas inseridas para exebição da coluna de descrição dos retornos do banco
  '--------------------------------------------------------
    '17/10/2007 - Anderson
    'Alteração realizada para considerar o caminho padrão do Quick Store
    If Dir$(gsDefaultPath & "QuickCnab.mdb") <> "" Then
      If IsNumeric(Grade32.Columns("CNAB_CodMovRet").Text) Then
        intCodRetornoCNAB = Grade32.Columns("CNAB_CodMovRet").Text
        Set rsDescCNAB = DBCNAB.OpenRecordset("SELECT * FROM Movimentos_Retorno WHERE Mov_CodigoMovimento = " & Grade32.Columns("CNAB_CodMovRet").Text, dbOpenSnapshot)
        
        With rsDescCNAB
          If (.BOF And .EOF) Then
            Grade32.Columns("CNAB_CodMovRet").Text = "Descrição inválida"
          Else
            Grade32.Columns("CNAB_CodMovRet").Text = Grade32.Columns("CNAB_CodMovRet").Text & " - " & _
                                       .Fields("Mov_DescricaoMovimento")
          End If
        End With
      End If
      
      Select Case intCodRetornoCNAB
        Case 3, 26, 30
          strTipoRetorno = "A"    'Procura tipo A
        Case 28
          strTipoRetorno = "B"    'Procura tipo B
        Case 6, 9, 17
          strTipoRetorno = "C"    'Procura tipo C
      End Select
      
      If IsNumeric(Grade32.Columns("CNAB_CodIdComplementar").Text) Then
        Set rsDescCNAB = DBCNAB.OpenRecordset(" SELECT * FROM IdentificacaoComplementar WHERE Idc_CodigoRejeicao = " & Grade32.Columns("CNAB_CodIdComplementar").Text & _
                                              " AND Idc_Tipo = '" & Trim(strTipoRetorno) & "'", dbOpenSnapshot)
        
        With rsDescCNAB
          If (.BOF And .EOF) Then
            Grade32.Columns("CNAB_CodIdComplementar").Text = "Descrição inválida"
          Else
            Grade32.Columns("CNAB_CodIdComplementar").Text = Grade32.Columns("CNAB_CodIdComplementar").Text & " - " & _
                                       .Fields("Idc_Descricao")
          End If
        End With
      End If
    End If
  '--------------------------------------------------------
  Exit Sub
  
Deu_Erro:
 Exit Sub

End Sub

Private Sub MontaTotalizadores()
  '11/06/2004 - Daniel
  
  lblQtdeTotaldeContasCH.Caption = m_QtdeTotaldeContasCH
  lblValorTotaldasContasCH = Format(m_ValorTotaldasContasCH, "###,###,###,##0.00")
  lblTotalRecebidoCH = Format(m_TotalRecebidoCH, "###,###,###,##0.00")
  lblTotalaReceberCH = Format(m_TotalaReceberCH, "###,###,###,##0.00")
  lblTotalRecebidoemDiaCH = Format(m_TotalRecebidoemDiaCH, "###,###,###,##0.00")
  lblTotalRecebidoAtrasadoCH = Format(m_TotalRecebidoAtrasadoCH, "###,###,###,##0.00")
  lblVencidoeNaoRecebidoCH = Format(m_VencidoeNaoRecebidoCH, "###,###,###,##0.00")
  
  lblQtdeTotaldeContasCA.Caption = m_QtdeTotaldeContasCA
  lblValorTotaldasContasCA = Format(m_ValorTotaldasContasCA, "###,###,###,##0.00")
  lblTotalRecebidoCA = Format(m_TotalRecebidoCA, "###,###,###,##0.00")
  lblTotalaReceberCA = Format(m_TotalaReceberCA, "###,###,###,##0.00")
  lblTotalRecebidoemDiaCA = Format(m_TotalRecebidoemDiaCA, "###,###,###,##0.00")
  lblTotalRecebidoAtrasadoCA = Format(m_TotalRecebidoAtrasadoCA, "###,###,###,##0.00")
  lblVencidoeNaoRecebidoCA = Format(m_VencidoeNaoRecebidoCA, "###,###,###,##0.00")
End Sub

Private Sub LimparVarsTotalizadoresCH()
  '11/06/2004 - Daniel

  m_QtdeTotaldeContasCH = 0
  m_ValorTotaldasContasCH = 0
  m_TotalRecebidoCH = 0
  m_TotalaReceberCH = 0
  m_TotalRecebidoemDiaCH = 0
  m_TotalRecebidoAtrasadoCH = 0
  m_VencidoeNaoRecebidoCH = 0
End Sub

Private Sub LimparVarsTotalizadoresCA()
  '11/06/2004 - Daniel

  m_QtdeTotaldeContasCA = 0
  m_ValorTotaldasContasCA = 0
  m_TotalRecebidoCA = 0
  m_TotalaReceberCA = 0
  m_TotalRecebidoemDiaCA = 0
  m_TotalRecebidoAtrasadoCA = 0
  m_VencidoeNaoRecebidoCA = 0
End Sub
  

