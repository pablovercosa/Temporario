VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmImprimeEtiq 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"ImprimeEtiquetas.frx":0000
   ClientHeight    =   8580
   ClientLeft      =   2925
   ClientTop       =   825
   ClientWidth     =   16125
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
   HelpContextID   =   1290
   Icon            =   "ImprimeEtiquetas.frx":0101
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8580
   ScaleWidth      =   16125
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab_impressoras 
      Height          =   4755
      Left            =   30
      TabIndex        =   39
      Top             =   3750
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   8387
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Impressoras Normais"
      TabPicture(0)   =   "ImprimeEtiquetas.frx":4EA5B
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Impressoras Específicas de Etiquetas"
      TabPicture(1)   =   "ImprimeEtiquetas.frx":4EA77
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "Tamanho de Etiquetas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   90
         TabIndex        =   60
         Top             =   420
         Width           =   13695
         Begin VB.CheckBox chk_colunaDaEsquerdaSemPreco 
            Caption         =   "Coluna da esquerda sem preço"
            Height          =   225
            Left            =   510
            TabIndex        =   79
            Top             =   1470
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.OptionButton opt_modelo_02_argox214 
            Appearance      =   0  'Flat
            Caption         =   "PPLA Argox 214"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   78
            Top             =   1170
            Width           =   1575
         End
         Begin VB.CheckBox chk_geraArquivo 
            Appearance      =   0  'Flat
            Caption         =   "Gerar Arquivo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   780
            TabIndex        =   76
            Top             =   3390
            Width           =   1365
         End
         Begin VB.OptionButton opt_zpl_epl_bematechLB1000 
            Appearance      =   0  'Flat
            Caption         =   "ZPL EPL Bematech LB-1000"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   70
            Top             =   810
            Width           =   2385
         End
         Begin VB.TextBox txtTemperatura 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4020
            TabIndex        =   67
            Text            =   "12"
            Top             =   2325
            Width           =   1560
         End
         Begin VB.ComboBox cboSaida 
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
            ItemData        =   "ImprimeEtiquetas.frx":4EA93
            Left            =   4020
            List            =   "ImprimeEtiquetas.frx":4EAB8
            TabIndex        =   66
            Text            =   "LPT1"
            Top             =   1440
            Width           =   1560
         End
         Begin VB.CommandButton cmd_imagem3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3825
            Left            =   5700
            Picture         =   "ImprimeEtiquetas.frx":4EB04
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   210
            Width           =   2835
         End
         Begin VB.CommandButton cmd_imagem4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3825
            Left            =   8610
            Picture         =   "ImprimeEtiquetas.frx":72092
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   210
            Width           =   2835
         End
         Begin VB.CommandButton cmd_imprimirEmImpressoraEtiquetas 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Imprimir Etiquetas"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   210
            Width           =   2085
         End
         Begin VB.OptionButton opt_modelo_01 
            Appearance      =   0  'Flat
            Caption         =   "PPLB Modelo 1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   61
            Top             =   450
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            Caption         =   "* Obrigatoriamente deverá utilizar o ResultPrintFile.exe"
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
            Height          =   255
            Left            =   780
            TabIndex        =   77
            Top             =   3690
            Width           =   4545
         End
         Begin VB.Label LBLSaida 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Saída"
            Height          =   195
            Left            =   4020
            TabIndex        =   69
            Top             =   1170
            Width           =   390
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Temperatura"
            Height          =   195
            Left            =   4020
            TabIndex        =   68
            Top             =   2070
            Width           =   930
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "* Clique sobre as imagens"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3180
            TabIndex        =   65
            Top             =   210
            Width           =   2475
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tamanho de Etiquetas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74910
         TabIndex        =   40
         Top             =   420
         Width           =   13695
         Begin VB.CheckBox chk_pequena_fonteNomeMenor 
            Appearance      =   0  'Flat
            Caption         =   "Com fonte do nome do produto menor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   930
            TabIndex        =   71
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton cmd_imprimir_ImpressoraNormal 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Gerar relatório para impressão"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   210
            Width           =   2085
         End
         Begin VB.OptionButton O_Pequena 
            Appearance      =   0  'Flat
            Caption         =   "Pequena  *Pimaco 6187 ou 6287"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   58
            Top             =   330
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton O_Média 
            Appearance      =   0  'Flat
            Caption         =   "Média      *Pimaco 6180 ou 6280"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   55
            Top             =   870
            Width           =   2865
         End
         Begin VB.OptionButton O_Grande 
            Appearance      =   0  'Flat
            Caption         =   "Grande para roupa - padrão Infopar - 9 etiq/folha"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   54
            Top             =   2190
            Width           =   4185
         End
         Begin VB.OptionButton O_Grande2 
            Appearance      =   0  'Flat
            Caption         =   "Grande - Impressão Descr. Produto (Pimaco 6195/6295)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -150
            TabIndex        =   53
            Top             =   3840
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.OptionButton O_Grande3 
            Appearance      =   0  'Flat
            Caption         =   "Grande para roupa - padrão Infopar - 12 etiq/folha"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   52
            Top             =   2520
            Width           =   4215
         End
         Begin VB.OptionButton O_Grande4 
            Appearance      =   0  'Flat
            Caption         =   "Grande para roupa - padrão PAULIMAQ LJC 253"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   51
            Top             =   2850
            Width           =   3975
         End
         Begin VB.ComboBox cboPosicao 
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
            Height          =   330
            ItemData        =   "ImprimeEtiquetas.frx":9573C
            Left            =   4740
            List            =   "ImprimeEtiquetas.frx":95746
            TabIndex        =   50
            Text            =   "12"
            Top             =   3060
            Width           =   885
         End
         Begin VB.OptionButton O_GrandeProcon 
            Appearance      =   0  'Flat
            Caption         =   "Grande     *Pimaco 6182 - Padrão PROCON"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   49
            Top             =   1530
            Width           =   3555
         End
         Begin VB.OptionButton O_Pequena2 
            Appearance      =   0  'Flat
            Caption         =   "Pequena com Lote (Pimaco 6187 / 6287)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -120
            TabIndex        =   48
            Top             =   3540
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.OptionButton O_GrandeProcon2 
            Appearance      =   0  'Flat
            Caption         =   "Grande - Padrão PROCON - Modelo Pimaco 8296"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   47
            Top             =   1860
            Width           =   4095
         End
         Begin VB.OptionButton optMedia6081 
            Appearance      =   0  'Flat
            Caption         =   "Média      *Pimaco 6081"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   46
            Top             =   1200
            Width           =   2145
         End
         Begin VB.CommandButton cmd_imagem1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3825
            Left            =   8610
            Picture         =   "ImprimeEtiquetas.frx":95752
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   210
            Width           =   2835
         End
         Begin VB.CommandButton cmd_imagem2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3825
            Left            =   5700
            Picture         =   "ImprimeEtiquetas.frx":B82C4
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   210
            Width           =   2835
         End
         Begin VB.Frame Frame4 
            Caption         =   "Saída"
            BeginProperty Font 
               Name            =   "WeblySleek UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2760
            TabIndex        =   41
            Top             =   3390
            Width           =   2865
            Begin VB.OptionButton O_Vídeo 
               Appearance      =   0  'Flat
               Caption         =   "Vídeo"
               BeginProperty Font 
                  Name            =   "WeblySleek UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   330
               TabIndex        =   43
               Top             =   300
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton O_Impressora 
               Appearance      =   0  'Flat
               Caption         =   "Impressora"
               BeginProperty Font 
                  Name            =   "WeblySleek UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1380
               TabIndex        =   42
               Top             =   300
               Width           =   1185
            End
         End
         Begin VB.Label lblPaulimaq 
            Caption         =   "Nº de caracteres para código de barras:  12 ou 09"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   930
            TabIndex        =   57
            Top             =   3120
            Width           =   3840
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "* Clique sobre as imagens"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3180
            TabIndex        =   56
            Top             =   210
            Width           =   2475
         End
      End
   End
   Begin VB.CommandButton cmd_limparTabelaDeEtiquetas 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Apagar a sua lista"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Apagar a lista de produtos e quantidades de etiquetas para impressão"
      Top             =   870
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Detalhes_PROCON 
      Caption         =   "Detalhes etiqueta padrão PROCON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   30
      TabIndex        =   20
      Top             =   2670
      Width           =   11145
      Begin VB.TextBox Num_Parcelas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   8
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox Msg_Linha2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   6690
         MaxLength       =   25
         TabIndex        =   12
         Top             =   600
         Width           =   4365
      End
      Begin VB.TextBox Msg_Linha1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   11
         Top             =   600
         Width           =   4515
      End
      Begin VB.ComboBox Tab_Prazo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8730
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   210
         Width           =   2355
      End
      Begin VB.ComboBox Tab_Vista 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ImprimeEtiquetas.frx":DB29E
         Left            =   5700
         List            =   "ImprimeEtiquetas.frx":DB2A0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   2385
      End
      Begin VB.OptionButton Financiamento 
         Appearance      =   0  'Flat
         Caption         =   "Financiamento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1590
         TabIndex        =   7
         Top             =   255
         Width           =   1455
      End
      Begin VB.OptionButton Parcelamento 
         Appearance      =   0  'Flat
         Caption         =   "Parcelamento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   255
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "A Prazo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8130
         TabIndex        =   27
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Msg linha 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5700
         TabIndex        =   26
         Top             =   645
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Msg linha 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Nºparcelas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3060
         TabIndex        =   22
         Top             =   255
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Tabela a Vista"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   255
         Width           =   1155
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ajustar margem da impressora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   11220
      TabIndex        =   17
      Top             =   30
      Width           =   2685
      Begin ComctlLib.Slider sldEsquerda 
         Height          =   495
         Left            =   420
         TabIndex        =   23
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   1
         Min             =   -7
         Max             =   7
      End
      Begin ComctlLib.Slider sldSuperior 
         Height          =   1335
         Left            =   1140
         TabIndex        =   25
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   2355
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1
         Min             =   -7
         Max             =   7
      End
      Begin VB.Label lblSuperior 
         Alignment       =   2  'Center
         Caption         =   "Superior = padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   510
         TabIndex        =   18
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label lblEsquerda 
         Alignment       =   2  'Center
         Caption         =   "Esquerda = padrão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   390
         TabIndex        =   19
         Top             =   1140
         Width           =   1815
      End
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
      Height          =   345
      Left            =   5100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   -210
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Etiqueta será composta por"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   30
      TabIndex        =   16
      Top             =   450
      Width           =   11145
      Begin VB.Frame Frame8 
         Height          =   945
         Left            =   9300
         TabIndex        =   73
         Top             =   150
         Width           =   1755
         Begin VB.TextBox txt_etiquetasEmBranco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   285
            Left            =   360
            MaxLength       =   2
            TabIndex        =   75
            Text            =   "1"
            Top             =   570
            Width           =   405
         End
         Begin VB.Label Label11 
            Caption         =   "Começar a imprimir da posição da etiqueta nº"
            Height          =   675
            Left            =   90
            TabIndex        =   74
            Top             =   150
            Width           =   1575
         End
      End
      Begin VB.CheckBox chk_fonteNomeLowerCase 
         Appearance      =   0  'Flat
         Caption         =   "Letras minúsculas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   660
         TabIndex        =   72
         Top             =   780
         Width           =   1725
      End
      Begin VB.Frame Frame3 
         Caption         =   "Etiqueta para Roupa "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         TabIndex        =   34
         Top             =   1590
         Width           =   11025
         Begin VB.CheckBox O_Troca 
            Appearance      =   0  'Flat
            Caption         =   "Mensagem de troca"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   37
            Top             =   240
            Width           =   1875
         End
         Begin VB.CheckBox O_Tamanho 
            Appearance      =   0  'Flat
            Caption         =   "Tamanho"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2340
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox O_Cor 
            Appearance      =   0  'Flat
            Caption         =   "Cor"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7230
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Código de Barras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         TabIndex        =   29
         Top             =   1020
         Width           =   11025
         Begin VB.OptionButton O_Imprime_Barras 
            Appearance      =   0  'Flat
            Caption         =   "Grande"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton O_Não_Imprime 
            Appearance      =   0  'Flat
            Caption         =   "Sem Código de barras"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7200
            TabIndex        =   31
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton O_Imprime_Pequena 
            Appearance      =   0  'Flat
            Caption         =   "Pequeno *Não funciona em algumas impressoras/leitores"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2310
            TabIndex        =   30
            Top             =   240
            Width           =   4755
         End
      End
      Begin VB.CheckBox chkAppendSubClasse 
         Appearance      =   0  'Flat
         Caption         =   "Nome da Sub-classe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7290
         TabIndex        =   5
         Top             =   525
         Width           =   1905
      End
      Begin VB.CheckBox Imprime_Nome 
         Appearance      =   0  'Flat
         Caption         =   "Nome do produto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   525
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.ComboBox Lista 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ImprimeEtiquetas.frx":DB2A2
         Left            =   3660
         List            =   "ImprimeEtiquetas.frx":DB2A4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   3315
      End
      Begin VB.CheckBox Imprime_preço 
         Appearance      =   0  'Flat
         Caption         =   "Preço do produto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   1665
      End
      Begin VB.CheckBox chkAppendClasse 
         Appearance      =   0  'Flat
         Caption         =   "Nome da Classe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2370
         TabIndex        =   4
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Tabela de Preços"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2340
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton B_Emite 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar Formatação de Etiquetas"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   13950
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   150
      Width           =   2115
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   6540
      Top             =   -300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "ImprimeEtiquetas.frx":DB2A6
      DataSource      =   "Data1"
      Height          =   345
      Left            =   1050
      TabIndex        =   0
      ToolTipText     =   "Digite 0 (zero) para todos."
      Top             =   75
      Width           =   1395
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
      Columns(0).Width=   3200
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   93
      Text            =   "0"
      BackColor       =   12648447
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
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1050
      Left            =   15570
      TabIndex        =   28
      Top             =   1230
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1852
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   16711680
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Funcionário"
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
      Height          =   255
      Left            =   30
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Nome_func 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   2490
      TabIndex        =   14
      Top             =   75
      Width           =   8685
   End
End
Attribute VB_Name = "frmImprimeEtiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsEtiquetas As Recordset
Private rsFuncionarios As Recordset
Private rsProdutos As Recordset
Private rsPreços As Recordset
Private rsEtiquetas_Tempo As Recordset
Private rsParametros As Recordset
Private rsTamanhos As Recordset
Private rsCores As Recordset
Private rsClasses As Recordset
Private rsSubclasses As Recordset
Private rsProdutosParaEtip As Recordset

Private Ajusta_Ver As Integer
Private Ajusta_Hor As Integer
Private Margem_Sup As Long
Private Margem_Inf As Long
Private Margem_Dir As Long
Private Margem_Esq As Long
Private sCaminhoArquivo As String

Public ParamCodigoUsuario As Long

Private Function nCheckValues(ByVal nType As TipoMargem) As Integer
  Dim nPosition As Integer
  Dim lblText As Label
  
  If nType = tmEsquerda Then
    Set lblText = lblEsquerda
    nPosition = sldEsquerda.Value
  Else
    Set lblText = lblSuperior
    nPosition = sldSuperior.Value
  End If
  
  If nPosition = 0 Then
    nCheckValues = 0
  Else
    nCheckValues = CInt(nPosition / 10 * 567)
  End If
  lblText.Caption = IIf(nType = tmSuperior, "Superior", "Esquerda") & _
    IIf(nPosition = 0, " = padrão", " = " & IIf(nPosition > 0, "+", "") & nPosition & " mm")
End Function

Private Function nTrataDetalhesPROCON(ByVal bStatus As Boolean) As Integer
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  
  Detalhes_PROCON.Enabled = bStatus
  Parcelamento.Enabled = bStatus
  Financiamento.Enabled = bStatus
  Num_Parcelas.Enabled = bStatus
  Tab_Vista.Enabled = bStatus
  Tab_Prazo.Enabled = bStatus
  Msg_Linha1.Enabled = bStatus
  Msg_Linha2.Enabled = bStatus
  Label1.Enabled = bStatus
  Label4.Enabled = bStatus
  Label5.Enabled = bStatus
  Label6.Enabled = bStatus
  Label7.Enabled = bStatus
    
End Function


Private Sub cmd_imagem1_Click()
    Dim objTelaImg As frmImprimeEtiquetaImagem
    
    Set objTelaImg = New frmImprimeEtiquetaImagem
    
    objTelaImg.sTituloPagina = "Sugestão de Modelo de Página de Etiquetas para compra"
    
    If O_Pequena.Value = True Then
        ' imagem tem 449x610 pixels
        objTelaImg.lAltura = 610 * 15.64
        objTelaImg.lLargura = 449 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6187_grande.bmp"
    ElseIf O_Média.Value = True Then
        objTelaImg.lAltura = 605 * 15.64
        objTelaImg.lLargura = 447 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6180_grande.bmp"
    ElseIf optMedia6081.Value = True Then
        objTelaImg.lAltura = 609 * 15.64
        objTelaImg.lLargura = 451 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6081_grande.bmp"
    ElseIf O_GrandeProcon.Value = True Then
        objTelaImg.lAltura = 608 * 15.64
        objTelaImg.lLargura = 449 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6182_grande.bmp"
    End If
 
    objTelaImg.Show 1
End Sub

Private Sub cmd_imagem2_Click()
    Dim objTelaImg As frmImprimeEtiquetaImagem
    
    Set objTelaImg = New frmImprimeEtiquetaImagem
    
    
    objTelaImg.sTituloPagina = "Modelo/exemplo IMPRESSO de Página de Etiqueta"
    
    ' Esta imagem tem as seguintes dimensões:
    If O_Pequena.Value = True Then
        ' imagem tem 726x610 pixels
        objTelaImg.lAltura = 610 * 15.64
        objTelaImg.lLargura = 726 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6187_grandeIMPRESSA.bmp"
    ElseIf O_Média.Value = True Then
        objTelaImg.lAltura = 605 * 15.64
        objTelaImg.lLargura = 746 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6180_grandeIMPRESSA.bmp"
    ElseIf optMedia6081.Value = True Then
        objTelaImg.lAltura = 646 * 15.64
        objTelaImg.lLargura = 781 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6081_grandeIMPRESSA.bmp"
    ElseIf O_GrandeProcon.Value = True Then
        objTelaImg.lAltura = 671 * 15.64
        objTelaImg.lLargura = 753 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6182_grandeIMPRESSA.bmp"
    End If
 
    objTelaImg.Show 1
End Sub

Private Sub cmd_imagem3_Click()
    Dim objTelaImg As frmImprimeEtiquetaImagem
    
    Set objTelaImg = New frmImprimeEtiquetaImagem
    
    
    objTelaImg.sTituloPagina = "Modelo/exemplo IMPRESSO de Página de Etiqueta"
    
    ' Esta imagem tem as seguintes dimensões:
    If opt_modelo_01.Value = True Then
        ' imagem tem 726x610 pixels
        objTelaImg.lAltura = 608 * 15.64
        objTelaImg.lLargura = 451 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaArgoxPPLB_01_etiqueta_grande.bmp"
    ElseIf opt_zpl_epl_bematechLB1000.Value = True Then
        objTelaImg.lAltura = 601 * 15.64
        objTelaImg.lLargura = 450 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaBematechLB1000_01_etiqueta_grande.bmp"
'''    ElseIf optMedia6081.Value = True Then
'''        objTelaImg.lAltura = 646 * 15.64
'''        objTelaImg.lLargura = 781 * 15.24
'''        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6081_grandeIMPRESSA.bmp"
'''    ElseIf O_GrandeProcon.Value = True Then
'''        objTelaImg.lAltura = 671 * 15.64
'''        objTelaImg.lLargura = 753 * 15.24
'''        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6182_grandeIMPRESSA.bmp"
    End If
 
    objTelaImg.Show 1
End Sub

Private Sub cmd_imagem4_Click()
    Dim objTelaImg As frmImprimeEtiquetaImagem
    
    Set objTelaImg = New frmImprimeEtiquetaImagem
    
    objTelaImg.sTituloPagina = "Sugestão de Modelo de Página de Etiquetas para compra"
    
    If opt_modelo_01.Value = True Then
        ' imagem tem 449x610 pixels
        objTelaImg.lAltura = 608 * 15.64
        objTelaImg.lLargura = 451 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaArgoxPPLB_01_grande.bmp"
    ElseIf opt_zpl_epl_bematechLB1000.Value = True Then
        objTelaImg.lAltura = 601 * 15.64
        objTelaImg.lLargura = 450 * 15.24
        objTelaImg.sImagem = "\Imagens\etiquetaBematechLB1000_01_grande.bmp"
'''    ElseIf optMedia6081.Value = True Then
'''        objTelaImg.lAltura = 609 * 15.64
'''        objTelaImg.lLargura = 451 * 15.24
'''        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6081_grande.bmp"
'''    ElseIf O_GrandeProcon.Value = True Then
'''        objTelaImg.lAltura = 608 * 15.64
'''        objTelaImg.lLargura = 449 * 15.24
'''        objTelaImg.sImagem = "\Imagens\etiquetaPimaco6182_grande.bmp"
    End If
 
    objTelaImg.Show 1
End Sub

Private Sub cmd_imprimir_ImpressoraNormal_Click()
  Dim sSql As String
  Dim Linhas As Long
  Dim sCodFunc As Integer
  Dim Prod As String
  Dim Tamanho As Integer
  Dim Tamanho_Str As String
  Dim Cor As Integer
  Dim Cor_Str As String
  Dim Erro As Integer
  Dim Descrição As String
  Dim Sem_Preço As String
  Dim i As Long
  Dim Str1 As String
  Dim Cod_Str As String
  Dim Cod_Cor_Str As String
  Dim Cod_Tam_Str As String
  Dim sObs As String
  '17/02/2005 - Daniel
  Dim strDescricao2 As String
  '20/05/2005 - Daniel
  'Tratamento de novos campos
  Dim blnImprimirUmaEtiq   As Boolean
  Dim blnImprimirPrecoEtiq As Boolean
  '08/06/2005 - Daniel
  'Tratamento para o campo [Etiquetas - Tempo].Funcionario
  Dim intFuncionario As Integer
  '30/10/2007 - Celso
  'Modelo novo de etiqueta com lote e validade
  Dim Lote As String
  Dim DataValidade As String
  
  '19/02/2008 - Celso
  'Implementação PROCON
  Dim nPreco_Vista As Double
  Dim nPreco_Prazo As Double
  Dim sTipo As String
  
  
  '----------------[Fim das declarações de variáveis]----------------
  
  On Error GoTo TratarErro '31/03/2005 - Daniel (Tratamento para erros)
  
  Call StatusMsg("")
  
  If Combo.Text = "" And Combo.Text = "0" Then
      MsgBox "Selecione um usuário dono da lista", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If O_Pequena.Value = True Then
    Margem_Sup = 720  '1.27 cm * 567 twips por cm
    Margem_Inf = Margem_Sup
    Margem_Esq = 794 '1.40 cm * 567 twips por com
    Margem_Dir = Margem_Esq
  End If
  
  '------------------------------------------------------------------
  '30/10/2007 - Celso
  If O_Pequena2.Value = True Then
    Margem_Sup = 720  '1.27 cm * 567 twips por cm
    Margem_Inf = Margem_Sup
    Margem_Esq = 794 '1.40 cm * 567 twips por com
    Margem_Dir = Margem_Esq
  End If
  '------------------------------------------------------------------
  
  If O_Média.Value = True Then
    Margem_Sup = 765  '1.35 cm * 567 twips por cm
    Margem_Inf = 624  '1.10 * 567
    Margem_Esq = 600  '1.10 cm * 567 twips por com
    Margem_Dir = 100  '1.00 * 567
  End If
  
  '17/06/2009 - mpdea
  If optMedia6081.Value Then
    Margem_Sup = 720  '1.27 cm * 567 twips por cm
    Margem_Inf = Margem_Sup
    Margem_Esq = 213  '0.377 cm * 567 twips por cm (231,759)
    Margem_Dir = Margem_Esq
  End If
  
  If O_Grande2.Value = True Then 'grande com descrição
    Margem_Sup = 1304  '2.30 cm * 567 twips por cm
    Margem_Inf = 1021  '1.80 * 567
    Margem_Esq = 1049  '1.85 cm * 567 twips por com
    Margem_Dir = 964   '1.70 * 567
  End If
  
  If O_Grande.Value = True Or O_Grande3.Value = True Then  'roupa
    Margem_Sup = 680   '1.20 cm * 567 twips por cm
    Margem_Inf = 567   '1.00 * 567
    Margem_Esq = 567   '1.00 cm * 567 twips por com
    Margem_Dir = 567   '1.00 * 567
  End If
        
  '19/02/2008 - Celso
  'Implementação PROCON
  If O_GrandeProcon2.Value = True Then
    Margem_Sup = 720    '1.27 cm * 567 twips por cm
    Margem_Inf = 720    '1.27 cm * 567
    Margem_Esq = 181    '0.32 cm * 567 twips por com
    Margem_Dir = 181    '0.32 * 567
  End If
    
  'Verifica opções selecionadas
  If Imprime_preço.Value = 1 Then
    If Lista.Text = "" Then
      DisplayMsg "Lista de preços incorreta, verifique."
      ' Lista.SetFocus
      Exit Sub
    End If
  End If
  
  If O_Imprime_Barras.Value = False Then
    If Imprime_preço.Value = False Then
      If Imprime_Nome.Value = False Then
        DisplayMsg "Escolha pelo menos uma opção de impressão (nome, preço, código)."
        Exit Sub
      End If
    End If
  End If
  
  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  If O_Grande4.Value Then
    If Not IsNumeric(cboPosicao.Text) Then
      MsgBox "Qtde. de caracteres para o Cód. de barras inválida, verifique.", vbExclamation, "Atenção"
      cboPosicao.SetFocus
      Exit Sub
    End If
  End If
  '--------------------------------------------------------
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  If O_GrandeProcon2 Then
    If Not IsNumeric(Num_Parcelas.Text) Then
      MsgBox "Número de parcelas deve ser numérico!", vbExclamation, "Atenção"
      Num_Parcelas.SetFocus
      Exit Sub
    End If
    If CInt(Num_Parcelas.Text) < 1 Or CInt(Num_Parcelas.Text) > 36 Then
      MsgBox "Número de parcelas deve ser entre 1 e 36. Verifique!", vbExclamation, "Atenção"
      Num_Parcelas.SetFocus
      Exit Sub
    End If
    If Tab_Vista.Text = "" Then
      DisplayMsg "Tabela de preços incorreta. Verifique!"
      Tab_Vista.SetFocus
      Exit Sub
    End If
    If Tab_Prazo.Text = "" Then
      DisplayMsg "Tabela de preços incorreta. Verifique!"
      Tab_Prazo.SetFocus
      Exit Sub
    End If
  End If
  '--------------------------------------------------
  
  'Coloca as Etiquetas do Funcionário escolhido
  'no arquivo Etiquetas - Tempo
  Call StatusMsg("Aguarde, gerando etiquetas...")
  DoEvents
  
''''''''''  Set rsEtiquetas_Tempo = db.OpenRecordset("Etiquetas - Tempo")
''''''''''  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
''''''''''  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
''''''''''  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
''''''''''
''''''''''  '04/11/2005 - mpdea
''''''''''  'Adicionado Sub Classe na etiqueta
''''''''''  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
''''''''''
''''''''''  rsPreços.Index = "Tabela"
''''''''''  rsProdutos.Index = "Código"
''''''''''  rsTamanhos.Index = "Código"
''''''''''  rsCores.Index = "Código"
''''''''''  rsClasses.Index = "Código"
''''''''''
''''''''''  '04/11/2005 - mpdea
''''''''''  'Adicionado Sub Classe na etiqueta
''''''''''  rsSubclasses.Index = "Código"
  
  sCodFunc = CInt(gsHandleNull(Combo.Text))
  Prod = ""
  Erro = False
  
'''''''''''''''  If Val(sCodFunc) > 0 Then
'''''''''''''''    Set rsEtiquetas = db.OpenRecordset("SELECT * FROM Etiquetas WHERE Funcionário = " & sCodFunc & " ORDER BY Funcionário, Produto, Tamanho, Cor", dbOpenDynaset)
'''''''''''''''  Else
'''''''''''''''    Set rsEtiquetas = db.OpenRecordset("SELECT * FROM Etiquetas ORDER BY Funcionário, Produto, Tamanho, Cor", dbOpenDynaset)
'''''''''''''''  End If
'''''''''''''''
'''''''''''''''  Do While Not rsEtiquetas.EOF
'''''''''''''''    Prod = rsEtiquetas("Produto")
'''''''''''''''    Tamanho = rsEtiquetas("Tamanho")
'''''''''''''''    Cor = rsEtiquetas("Cor")
'''''''''''''''    Tamanho_Str = ""
'''''''''''''''    Cor_Str = ""
'''''''''''''''    If Cor <> 0 Then
'''''''''''''''      If O_Cor.Value = 1 Then
'''''''''''''''        rsCores.Seek "=", Cor
'''''''''''''''        If Not rsCores.NoMatch Then Cor_Str = rsCores("Nome")
'''''''''''''''      End If
'''''''''''''''    End If
'''''''''''''''    If Tamanho <> 0 Then
'''''''''''''''      If O_Tamanho.Value = 1 Then
'''''''''''''''        rsTamanhos.Seek "=", Tamanho
'''''''''''''''        If Not rsTamanhos.NoMatch Then Tamanho_Str = rsTamanhos("Nome")
'''''''''''''''      End If
'''''''''''''''    End If
'''''''''''''''
'''''''''''''''    '  If O_Grande = False Then 'não é etiqueta de roupa
'''''''''''''''    Descrição = ""
'''''''''''''''    Sem_Preço = ""
'''''''''''''''    sObs = ""
'''''''''''''''
'''''''''''''''    If Imprime_preço.Value = 1 Then
'''''''''''''''      rsPreços.Seek "=", Lista.Text, rsEtiquetas("Produto")
'''''''''''''''      '19/04/2007 - Anderson
'''''''''''''''      'Incluído clausula para verificar se é etiqueta padrão Procon
'''''''''''''''      'If Not rsPreços.NoMatch Then
'''''''''''''''      If Not rsPreços.NoMatch And Not O_GrandeProcon Then
'''''''''''''''        Descrição = Descrição + "R$ " + Format$(rsPreços("Preço"), "###,###,##0.00") + " "
'''''''''''''''      End If
'''''''''''''''    End If
'''''''''''''''
'''''''''''''''    If Imprime_Nome.Value = 1 Then
'''''''''''''''       rsProdutos.Seek "=", rsEtiquetas("Produto")
'''''''''''''''       If Not rsProdutos.NoMatch Then
'''''''''''''''          Descrição = Descrição + rsProdutos("Nome")
'''''''''''''''          Sem_Preço = rsProdutos("Nome")
'''''''''''''''          sObs = rsProdutos("Obs") & ""
'''''''''''''''          If chkAppendClasse.Value = 1 Then
'''''''''''''''            rsClasses.Seek "=", rsProdutos("Classe")
'''''''''''''''            If Not rsClasses.NoMatch Then
'''''''''''''''              Descrição = Trim(Descrição) & " " & Trim(rsClasses("Nome"))
'''''''''''''''            End If
'''''''''''''''          End If
'''''''''''''''          '04/11/2005 - mpdea
'''''''''''''''          'Adicionado Sub Classe na etiqueta
'''''''''''''''          If chkAppendSubClasse.Value = 1 Then
'''''''''''''''            rsSubclasses.Seek "=", rsProdutos("Sub Classe")
'''''''''''''''            If Not rsSubclasses.NoMatch Then
'''''''''''''''              Descrição = Trim(Descrição) & " " & Trim(rsSubclasses("Nome"))
'''''''''''''''            End If
'''''''''''''''          End If
'''''''''''''''       Else
'''''''''''''''          Descrição = rsEtiquetas("Produto")
'''''''''''''''          Sem_Preço = ""
'''''''''''''''          sObs = ""
'''''''''''''''       End If
'''''''''''''''    End If
'''''''''''''''
'''''''''''''''    '17/02/2005 - Daniel
'''''''''''''''    '
'''''''''''''''    'Tratamento para o campo [Etiquetas - Tempo].Descricao2
'''''''''''''''    'Código de barras
'''''''''''''''    strDescricao2 = rsProdutos("Nome") & ""
'''''''''''''''    '
'''''''''''''''    '20/05/2005 - Daniel
'''''''''''''''    'Tratamento para os campos ImprimirUmaEtiq e ImprimirPrecoEtiq
'''''''''''''''    blnImprimirUmaEtiq = rsProdutos("ImprimirUmaEtiq").Value
'''''''''''''''    blnImprimirPrecoEtiq = rsProdutos("ImprimirPrecoEtiq").Value
'''''''''''''''    '
'''''''''''''''    '08/06/2005 - Daniel
'''''''''''''''    'Tratamento para o campo [Etiquetas - Tempo].Funcionario
'''''''''''''''    If Len(Nome_func.Caption) > 0 Then 'Funcionário preenchido
'''''''''''''''      intFuncionario = CInt(Combo.Text)
'''''''''''''''    Else
'''''''''''''''      intFuncionario = 0
'''''''''''''''    End If
'''''''''''''''
'''''''''''''''  ' 19/02/2008 - Celso
'''''''''''''''  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
'''''''''''''''    If O_GrandeProcon2.Value Then
'''''''''''''''      rsPreços.Seek "=", Tab_Vista.Text, rsEtiquetas("Produto")
'''''''''''''''      If Not rsPreços.NoMatch Then
'''''''''''''''         nPreco_Vista = CDbl(rsPreços.Fields("Preço"))
'''''''''''''''      Else
'''''''''''''''         nPreco_Vista = 0
'''''''''''''''      End If
'''''''''''''''      rsPreços.Seek "=", Tab_Prazo.Text, rsEtiquetas("Produto")
'''''''''''''''      If Not rsPreços.NoMatch Then
'''''''''''''''         nPreco_Prazo = CDbl(rsPreços.Fields("Preço"))
'''''''''''''''      Else
'''''''''''''''         nPreco_Prazo = 0
'''''''''''''''      End If
'''''''''''''''    End If
'''''''''''''''
'''''''''''''''    If Parcelamento.Value = True Then
'''''''''''''''      sTipo = "Tipo ='1'"   'Parcelamento
'''''''''''''''    Else
'''''''''''''''      sTipo = "Tipo ='2'"   'Financiamento
'''''''''''''''    End If
'''''''''''''''
'''''''''''''''  '----------------------------------------------
'''''''''''''''
'''''''''''''''    For i = 1 To rsEtiquetas("Qtde")
'''''''''''''''      Cod_Str = rsEtiquetas("Produto")
'''''''''''''''      Cod_Str = Trim(Cod_Str)
'''''''''''''''
'''''''''''''''      Cod_Cor_Str = ""
'''''''''''''''      If rsEtiquetas("Cor") <> 0 Then
'''''''''''''''        Cod_Cor_Str = "000" + Trim(str(rsEtiquetas("Cor")))
'''''''''''''''        Cod_Cor_Str = Right$(Cod_Cor_Str, 3)
'''''''''''''''      End If
'''''''''''''''
'''''''''''''''      Cod_Tam_Str = ""
'''''''''''''''      If rsEtiquetas("Tamanho") <> 0 Then
'''''''''''''''        Cod_Tam_Str = "000" + Trim(str(rsEtiquetas("Tamanho")))
'''''''''''''''        Cod_Tam_Str = Right$(Cod_Tam_Str, 3)
'''''''''''''''      End If
'''''''''''''''
'''''''''''''''      Cod_Str = Cod_Str + Cod_Tam_Str + Cod_Cor_Str
'''''''''''''''
'''''''''''''''      With rsEtiquetas_Tempo
'''''''''''''''        .AddNew
'''''''''''''''        .Fields("Código") = Cod_Str
'''''''''''''''        .Fields("Código Barra") = "*" + UCase(Cod_Str) + "*"
'''''''''''''''        .Fields("Código Produto") = rsEtiquetas("Produto")
'''''''''''''''        .Fields("Sem Preço") = Left(Sem_Preço, .Fields("Sem Preço").Size)
'''''''''''''''
'''''''''''''''        '16/08/2005 - mpdea
'''''''''''''''        'Corrigido RT-3021
'''''''''''''''        'Incluído verificação de preço localizado
'''''''''''''''        If Not rsPreços.NoMatch Then
'''''''''''''''          .Fields("Preco") = CDbl(rsPreços.Fields("Preço"))
'''''''''''''''        End If
'''''''''''''''
'''''''''''''''        .Fields("Tamanho") = Left(Tamanho_Str, .Fields("Tamanho").Size)
'''''''''''''''        .Fields("Cor") = Left(Cor_Str, .Fields("Cor").Size)
'''''''''''''''        .Fields("Texto Grande") = sObs
'''''''''''''''        '23/04/2007 - Anderson
'''''''''''''''        'Incluída geração de código completo quando selecionado etiqueta modelo Procon sem impressão do código de barras
'''''''''''''''        If O_Não_Imprime And O_GrandeProcon Then
'''''''''''''''          Descrição = Cod_Str & " - " & Descrição
'''''''''''''''        End If
'''''''''''''''        If O_GrandeProcon Then
'''''''''''''''          If Len(.Fields("Tamanho")) > 0 And O_Tamanho.Value Then
'''''''''''''''            Descrição = Descrição & " - TAM: " & .Fields("Tamanho")
'''''''''''''''          End If
'''''''''''''''          If Len(.Fields("Cor")) > 0 And O_Cor.Value Then
'''''''''''''''            Descrição = Descrição & " - COR: " & .Fields("Cor")
'''''''''''''''          End If
'''''''''''''''        End If
'''''''''''''''        .Fields("Descrição") = Left(Descrição, .Fields("Descrição").Size)
'''''''''''''''        '22/05/2007 - Anderson
'''''''''''''''        'Incluído o código para evitar a repetição do código do produto na descrição
'''''''''''''''        If O_Não_Imprime And O_GrandeProcon Then
'''''''''''''''          Descrição = Mid(Descrição, Len(Cod_Str & " - ") + 1)
'''''''''''''''        End If
'''''''''''''''        '17/02/2005 - Daniel
'''''''''''''''        '
'''''''''''''''        'Tratamento para o campo [Etiquetas - Tempo].Descricao2
'''''''''''''''        'neste campo será armazenado apenas o Nome do Produto que
'''''''''''''''        'terá finalidade de atender etiquetas com 09 caracteres no
'''''''''''''''        'código de barras
'''''''''''''''        .Fields("Descricao2").Value = Left(strDescricao2 & "", .Fields("Descricao2").Size)
'''''''''''''''        '20/05/2005 - Daniel
'''''''''''''''        'Tratamento para os campos ImprimirUmaEtiq e ImprimirPrecoEtiq
'''''''''''''''        .Fields("ImprimirUmaEtiq").Value = blnImprimirUmaEtiq
'''''''''''''''        .Fields("ImprimirPrecoEtiq").Value = blnImprimirPrecoEtiq
'''''''''''''''        '08/06/2005 - Daniel
'''''''''''''''        'Tratamento para o campo [Etiquetas - Tempo].Funcionario
'''''''''''''''        .Fields("Funcionario").Value = intFuncionario
'''''''''''''''        '-------------------------------------------------------------
'''''''''''''''
'''''''''''''''        '19/04/2007 - Anderson
'''''''''''''''        'Implementação da divisão do preço na etiqueta atendendo as exigências do PROCON
'''''''''''''''        .Fields("DividirPrecoEtiqueta").Value = rsProdutos("DividirPrecoEtiqueta")
'''''''''''''''        '30/10/2007 - Celso
'''''''''''''''        'Modelo novo de etiqueta com lote e data validade
'''''''''''''''        .Fields("Lote").Value = rsProdutos("Lote") + " "
'''''''''''''''        .Fields("DataValidade").Value = rsProdutos("DataValidade")
'''''''''''''''
'''''''''''''''        '19/02/2008 - Celso
'''''''''''''''        'Implementação PROCON
'''''''''''''''        If O_GrandeProcon2.Value Then
'''''''''''''''          .Fields("Preco") = nPreco_Vista
'''''''''''''''          .Fields("PrecoPrazo") = nPreco_Prazo
'''''''''''''''        End If
'''''''''''''''
'''''''''''''''        .Update
'''''''''''''''      End With
'''''''''''''''    Next i
'''''''''''''''
'''''''''''''''    rsEtiquetas.MoveNext
'''''''''''''''  Loop
  
  'Seta Valores e Manda Relatório
  
  '31/03/2005 - Daniel
  'Resetando o objeto Rel
  'Estava ocorrendo o erro do Crystal "Error opening file"
  'File could not be found "Etiquetas - Tempo", at file location "Etiquetas - Tempo"
  'e em seguida estava sendo tratado o erro através do TratarErro e era exibida a
  'mensagem: Run-time error '20533' Unable to open database
  Rel.Reset
  '--------------------------------------------------------------------------------
  
  '04/04/2005 - Daniel
  'Correção para evitar a mensagem: Nº de tabelas inválidas
  If O_Grande4.Value Then
    'Nome do BD
    Str1 = gsQuickDBFileName
    Rel.DataFiles(0) = Str1
    '31/03/2005 - Daniel
    'Adicionado linha abaixo para evitar bug
    Rel.DataFiles(1) = Str1
  Else
    'Nome do BD
    Str1 = gsQuickDBFileName
    Rel.DataFiles(0) = Str1
  End If
  
  'Saída
  If O_Vídeo = True Then Rel.Destination = 0
  If O_Impressora = True Then Rel.Destination = 1
  
  'Estado da janela
  Rel.WindowState = crptMaximized
  
  '12/05/2005 - Daniel
  'Correção para exibição dos botões de Configuração
  'de Impressoras e Botão de Pesquisas
  Rel.WindowShowPrintSetupBtn = True
  Rel.WindowShowSearchBtn = True
  
  'Nome do relatório
  If O_Imprime_Barras.Value Then
    If O_Pequena.Value = True And chk_pequena_fonteNomeMenor.Value = vbUnchecked Then
      Str1 = gsReportPath & "ETIQP1.RPT"              ' ATUALIZADO
    ElseIf O_Pequena.Value = True And chk_pequena_fonteNomeMenor.Value = vbChecked Then
      Str1 = gsReportPath & "Etiqp1_letra_tam6.RPT"   ' ATUALIZADO
    ElseIf O_Pequena2.Value Then
      Str1 = gsReportPath & "ETIQP4.RPT"              ' ATUALIZADO
    ElseIf O_Média.Value Then
      '07/03/2003 - Maikel Cordeiro
      'Adicionadas as clausulas abaixo para a etiqueta média com tamanho
      '---------------------------------------------------------------
      If O_Tamanho.Value = vbChecked Then
        Str1 = gsReportPath & "ETIQM1T.RPT"           ' ATUALIZADO
      Else
        Str1 = gsReportPath & "ETIQM1.RPT"            ' ATUALIZADO
      End If
      '---------------------------------------------------------------
    ElseIf optMedia6081.Value Then
      '17/06/2009 - mpdea
      Str1 = gsReportPath & "Etiq6081.RPT"            ' ATUALIZADO
    ElseIf O_Grande.Value Then
      Str1 = gsReportPath & "ETIQROUP.RPT"            ' ATUALIZADO
    ElseIf O_Grande3.Value Then
      Str1 = gsReportPath & "ETIQROU2.RPT"            ' ATUALIZADO
    ElseIf O_Grande2.Value Then
      Str1 = gsReportPath & "ETIQG1.RPT"              ' ATUALIZADO
    ElseIf O_GrandeProcon.Value Then '19/04/2007 - Anderson - Implementação de etiqueta - Exigência Procon
      Str1 = gsReportPath & "ETIQG4.RPT"              ' ATUALIZADO
    End If
  ElseIf O_Imprime_Pequena.Value Then
    If O_Pequena.Value Then
      Str1 = gsReportPath & "ETIQP3.RPT"              ' ATUALIZADO
    ElseIf O_Pequena2.Value Then
      Str1 = gsReportPath & "ETIQP4.RPT"              ' ATUALIZADO
    ElseIf O_Média.Value Then
      '07/03/2003 - Maikel Cordeiro
      'Adicionadas as clausulas abaixo para a etiqueta média com tamanho
      '---------------------------------------------------------------
      If O_Tamanho.Value = vbChecked Then
        Str1 = gsReportPath & "ETIQM3T.RPT"              ' ATUALIZADO
      Else
        Str1 = gsReportPath & "ETIQM3.RPT"              ' ATUALIZADO
      End If
    ElseIf O_Grande.Value Then
      Str1 = gsReportPath & "ETIQROU3.RPT"              ' ATUALIZADO
    ElseIf O_Grande3.Value Then
      Str1 = gsReportPath & "ETIQROU4.RPT"              ' ATUALIZADO
    ElseIf O_Grande2.Value Then
      Str1 = gsReportPath & "ETIQG3.RPT"                ' ATUALIZADO
    ElseIf O_GrandeProcon.Value Then '19/04/2007 - Anderson - Implementação de etiqueta - Exigência Procon
      Str1 = gsReportPath & "ETIQG5.RPT"                ' ATUALIZADO
    End If
  ElseIf O_Não_Imprime.Value Then
    If O_Pequena.Value Then
      Str1 = gsReportPath & "ETIQP2.RPT"              ' ATUALIZADO
    ElseIf O_Pequena2.Value Then ' 30/10/2007 - Celso - Novo modelo etiqueta
      Str1 = gsReportPath & "ETIQP4.RPT"              ' ATUALIZADO

    ElseIf O_Média.Value Then
      '07/03/2003 - Maikel Cordeiro
      'Adicionadas as clausulas abaixo para a etiqueta média com tamanho
      '---------------------------------------------------------------
      If O_Tamanho.Value = vbChecked Then
        Str1 = gsReportPath & "ETIQM2T.RPT"            ' ATUALIZADO
      Else
        Str1 = gsReportPath & "ETIQM2.RPT"            ' ATUALIZADO
      End If
    ElseIf O_Grande.Value Then
      Str1 = gsReportPath & "ETIQROUP.RPT"            ' ATUALIZADO
    ElseIf O_Grande3.Value Then
      Str1 = gsReportPath & "ETIQROU2.RPT"            ' ATUALIZADO
    ElseIf O_Grande2.Value Then
      Str1 = gsReportPath & "ETIQG2.RPT"            ' ATUALIZADO
    ElseIf O_GrandeProcon.Value Then '19/04/2007 - Anderson - Implementação de etiqueta - Exigência Procon
      Str1 = gsReportPath & "ETIQG6.RPT"            ' ATUALIZADO
    End If
  End If
  
  If O_Grande4.Value Then
    '16/02/2005 - Daniel
    '
    'Solicitante..: Consultora Marineida
    '
    'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
    '
    'Tratamento para Qtde. de caracteres a serem impressas no
    'Código de barras
    If cboPosicao.Text = "12" Then
      Str1 = gsReportPath & "EtiqRou5.rpt"            ' ATUALIZADO
    Else
      Str1 = gsReportPath & "EtiqRou5B.rpt"            ' ATUALIZADO
    End If
  End If
  
  If O_GrandeProcon2.Value Then '20/02/2008 - Celso - Nova Implementação PROCON
     Str1 = gsReportPath & "ETIQG7.RPT"              ' ATUALIZADO
  End If
    
  Rel.ReportFileName = Str1

'''  Rel.Formulas(0) = ""
'''  Rel.Formulas(1) = ""
'''  Rel.Formulas(2) = ""
  Dim bSair As Boolean
  bSair = False

  If Combo.Text <> "" And Combo.Text <> "0" Then
      Dim rsEtiquetasAux As Recordset
      Set rsEtiquetasAux = db.OpenRecordset("SELECT min(seq), max(seq) FROM [Etiquetas - Tempo] where Funcionario=" & Combo.Text, dbOpenDynaset)
      
      If rsEtiquetasAux.RecordCount > 0 Then
          rsEtiquetasAux.MoveFirst
          
          If Not IsNull(rsEtiquetasAux.Fields(0).Value) Then
              Rel.ParameterFields(0) = "pSeqInicial;" & rsEtiquetasAux.Fields(0).Value & ";true"
              Rel.ParameterFields(1) = "pSeqFinal;" & rsEtiquetasAux.Fields(1).Value & ";true"
          Else
              Rel.ParameterFields(0) = "pSeqInicial;0;true"
              Rel.ParameterFields(1) = "pSeqFinal;0;true"
              
              bSair = True
          End If
      End If
      rsEtiquetasAux.Close
      Set rsEtiquetasAux = Nothing
  End If
  
  If bSair = True Then
      MsgBox "Não existe lista de etiquetas para este usuário", vbInformation, "Atenção"
      Exit Sub
  End If
  
  '17/06/2009 - mpdea
  If optMedia6081.Value Then
    Str1 = "Mensagem = '" + (rsParametros("Mensagem Troca") & "") + "'"
    If O_Troca = 0 Then Str1 = "Mensagem = ''"
    Rel.Formulas(0) = Str1
  End If
  
  '06/10/2004 - Daniel
  'Adicionado a linha: Or O_Grande4.Value
  'Case: Paulimaq não imprimia nunca a mensagem de troca
  If O_Grande.Value = True Or O_Grande3.Value = True Or O_Grande4.Value Then
    Str1 = "Mensagem = '" + (rsParametros("Mensagem Troca") & "") + "'"
    If O_Troca = 0 Then Str1 = "Mensagem = ''"
    Rel.Formulas(0) = Str1
    Str1 = "Fone1 = '" + (rsParametros("Mensagem Etiq 1") & "") + "'"
    Rel.Formulas(1) = Str1
    Str1 = "Fone2 = '" + (rsParametros("Mensagem Etiq 2") & "") + "'"
    Rel.Formulas(2) = Str1
  End If
  
  '19/02/2008 - Celso
  'Implementação PROCON
  If O_GrandeProcon2.Value = True Then
    Rel.Formulas(0) = sTipo
    Rel.Formulas(1) = "Msg1 ='" & Msg_Linha1.Text & "" & "'"
    Rel.Formulas(2) = "Msg2 ='" & Msg_Linha2.Text & "" & "'"
    Rel.Formulas(3) = "NumParc ='" & Num_Parcelas.Text & "" & "'"
  End If
  
  'Margens
  Margem_Sup = Margem_Sup + Ajusta_Hor
  Margem_Inf = Margem_Inf - Ajusta_Hor
  
  Margem_Esq = Margem_Esq + Ajusta_Ver
  Margem_Dir = Margem_Dir - Ajusta_Ver
  
  Rel.MarginTop = Margem_Sup
  Rel.MarginBottom = Margem_Inf
  Rel.MarginLeft = Margem_Esq
  Rel.MarginRight = Margem_Dir
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  
  Rel.Action = 1
  
''''''''''  rsEtiquetas_Tempo.Close
''''''''''  rsTamanhos.Close
''''''''''  rsCores.Close
''''''''''  rsEtiquetas.Close
''''''''''  Set rsEtiquetas_Tempo = Nothing
''''''''''  Set rsTamanhos = Nothing
''''''''''  Set rsCores = Nothing
''''''''''  Set rsEtiquetas = Nothing
''''''''''
''''''''''  '04/11/2005 - mpdea
''''''''''  'Fecha tabelas
''''''''''  rsClasses.Close
''''''''''  rsSubclasses.Close
''''''''''  Set rsClasses = Nothing
''''''''''  Set rsSubclasses = Nothing
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
End Sub


Private Sub cmd_imprimirEmImpressoraEtiquetas_Click()
On Error GoTo Erro
  Dim rstEtiqueta   As Recordset
  Dim strAuxi       As String
  Dim intbuffer     As Integer
  Dim sCaminhoAux As String

    
  If opt_modelo_01.Value = True Then
     
     intbuffer = FreeFile
  
     If Combo.Text = "" And Combo.Text = "0" Then
         MsgBox "Selecione um usuário dono da lista", vbInformation, "Atenção"
         Exit Sub
     End If
 
  '''    If arrUsuario(0) = "XX" Then
  '''       Set rstEtiqueta = db.OpenRecordset("SELECT * FROM [Etiquetas - Tempo] ORDER BY Código", dbOpenDynaset)
  '''    Else
         Set rstEtiqueta = db.OpenRecordset("SELECT * FROM [Etiquetas - Tempo] Where Funcionario = " & Combo.Text & " ORDER BY Código", dbOpenDynaset)
  '''    End If

     '----[Porta de Saída]----
     If chk_geraArquivo.Value = vbChecked Then
        sCaminhoAux = Replace(LCase(sCaminhoArquivo), "\\tsclient\c\", "C:") & "\ETIQUETAS"
        sCaminhoAux = Replace(LCase(sCaminhoArquivo), "\\tsclient\d\", "D:") & "\ETIQUETAS"
        MsgBox "Este caminho deve existir na sua máquina. " & sCaminhoAux, vbInformation, "Atenção"
        Open sCaminhoArquivo & "\ETIQUETAS\arquivoEtiquetas_HR" & Hour(Now) & "_" & Minute(Now) & "_DIA" & Day(Now) & "_" & Month(Now) & "_" & Year(Now) & ".p1" For Output As #intbuffer  '...Arquivo texto
     Else
        Open cboSaida.Text For Output As #intbuffer  'abre a saída LPT1 pode ser com COM1...Arquivo texto
     End If

     '---[verificando se há etiquetas para imprimir]---
     If rstEtiqueta.EOF And rstEtiqueta.BOF Then
         MsgBox "Nenhuma lista de Etiqueta Formatada para impressão para este usuário selecionado!" & vbCrLf & vbCrLf & "PASSO 1: Vá até a tela de MONTAR LISTA para criar sua lista de etiquetas." & vbCrLf & vbCrLf & "PASSE 2: Vá até a tela de FORMATAR ETIQUETAS para incluir dados de preço, classe, etc e Formate suas etiquetas.", vbInformation, "Quick Store" ' messagem caso não tenha etiquetas
     Else '
 
        rstEtiqueta.MoveFirst
 
        '---[impressão das etiquetas]---
        Do While Not rstEtiqueta.EOF
            '----[Configurações Inicias da Impressora - PPLB ]----
            Print #intbuffer, "O"
            Print #intbuffer, "N"
            Print #intbuffer, "q840"        '"q840"        '105mm * 8 = 840 Comprimento da etiqueta "Rolo" Hor.
            Print #intbuffer, "Q448,16+0"   '"Q448,16+0"   '56 * 8 = 448    Q448,16+0
            Print #intbuffer, "S2"
            Print #intbuffer, "D" & CInt(txtTemperatura.Text)
            Print #intbuffer, "ZT"
            Print #intbuffer, "TTh:m"
            Print #intbuffer, "TDy2.mn.dd"
            
            ' Etiqueta 1
            Print #intbuffer, "A824,352,2,2,1,1,N,""" & ConverteCaracter(NomePart1(rstEtiqueta.Fields("[Sem Preço]"))) & """"
            Print #intbuffer, "A824,328,2,2,1,1,N,""" & ConverteCaracter(NomePart2(rstEtiqueta.Fields("[Sem Preço]"))) & """"
            Print #intbuffer, "A824,304,2,2,1,1,N,""" & ConverteCaracter(NomePart3(rstEtiqueta.Fields("[Sem Preço]"))) & """"
            strAuxi = BuscarCorTam(rstEtiqueta.Fields("Código").Value)
            If Len(strAuxi) > 0 Then Print #intbuffer, "A824,272,2,2,1,1,N,""" & strAuxi & """"
            Print #intbuffer, "B824,240,2,1,2,2,28,B,""" & rstEtiqueta.Fields("Código") & """"
            Print #intbuffer, "B824,120,2,1,2,2,28,B,""" & rstEtiqueta.Fields("Código") & """"
            Print #intbuffer, "A824,40,2,2,1,1,N,""R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00") & """"
    
            rstEtiqueta.MoveNext
            ' Etiqueta 2
            If Not rstEtiqueta.EOF Then
               Print #intbuffer, "A544,352,2,2,1,1,N,""" & ConverteCaracter(NomePart1(rstEtiqueta.Fields("[Sem Preço]"))) & """"
               Print #intbuffer, "A544,328,2,2,1,1,N,""" & ConverteCaracter(NomePart2(rstEtiqueta.Fields("[Sem Preço]"))) & """"
               Print #intbuffer, "A544,304,2,2,1,1,N,""" & ConverteCaracter(NomePart3(rstEtiqueta.Fields("[Sem Preço]"))) & """"
               strAuxi = BuscarCorTam(rstEtiqueta.Fields("Código").Value)
               If Len(strAuxi) > 0 Then Print #intbuffer, "A544,272,2,2,1,1,N,""" & strAuxi & """"
               Print #intbuffer, "B544,240,2,1,2,2,28,B,""" & rstEtiqueta.Fields("Código") & """"
               Print #intbuffer, "B544,120,2,1,2,2,28,B,""" & rstEtiqueta.Fields("Código") & """"
               Print #intbuffer, "A544,40,2,2,1,1,N,""R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00") & """"
    
               rstEtiqueta.MoveNext
            End If
            If Not rstEtiqueta.EOF Then
                 Print #intbuffer, "A264,352,2,2,1,1,N,""" & ConverteCaracter(NomePart1(rstEtiqueta.Fields("[Sem Preço]"))) & """"
                 Print #intbuffer, "A264,328,2,2,1,1,N,""" & ConverteCaracter(NomePart2(rstEtiqueta.Fields("[Sem Preço]"))) & """"
                 Print #intbuffer, "A264,304,2,2,1,1,N,""" & ConverteCaracter(NomePart3(rstEtiqueta.Fields("[Sem Preço]"))) & """"
                 strAuxi = BuscarCorTam(rstEtiqueta.Fields("Código").Value)
                 If Len(strAuxi) > 0 Then Print #intbuffer, "A264,272,2,2,1,1,N,""" & strAuxi & """"
                 Print #intbuffer, "B264,240,2,1,2,2,28,B,""" & rstEtiqueta.Fields("Código") & """"
                 Print #intbuffer, "B264,120,2,1,2,2,28,B,""" & rstEtiqueta.Fields("Código") & """"
                 Print #intbuffer, "A264,40,2,2,1,1,N,""R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00") & """"
    
                 rstEtiqueta.MoveNext
            End If
            Print #intbuffer, "FE"
            Print #intbuffer, "P1"
        Loop
     End If
     
     '---[Fechando a saída]---
     Close #intbuffer
  
     '---[Finalizando Variaveis]---
      rstEtiqueta.Close
     Set rstEtiqueta = Nothing

  ElseIf opt_zpl_epl_bematechLB1000.Value = True Then

      '----[Porta de Saída]----
      ' Abrir a saída USB
      Call openport(cboSaida.Text)

      ' 105 altura da etiqueta
      ' 55 a largura da etiqueta (somadas as 3 que estão na mesma linha)
      Call setup("105", "55", "3.0", "12", "1", "1", "0")
      Call sendcommand("SET TEAR ON")
    
      ' Direction 1 para inverter de 0º para 180º tudo que escrever na etiqueta e, vice-versa
      Call sendcommand("DIRECTION 1")
    
      ' Limpar o buffer toda ver que imprimir uma linha de 3 etiquetas
      Call clearbuffer
    
      '---[Parte de acesso a base de dados]---
      Set rstEtiqueta = db.OpenRecordset("SELECT * FROM [Etiquetas - Tempo] Where Funcionario = " & Combo.Text & " ORDER BY Código", dbOpenDynaset)
    
      ' Ponteiro no início
      rstEtiqueta.MoveFirst
    
      '---[verificando se há etiquetas para imprimir]---
      If rstEtiqueta.EOF Then
          MsgBox "Nenhuma etiqueta a imprimir", vbInformation, "Quick Store"
      Else
        '---[impressão das etiquetas]---
        Do While Not rstEtiqueta.EOF

            ' Etiqueta 1
            Call windowsfont(15, 150, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart1(rstEtiqueta.Fields("[Sem Preço]"))))
            Call windowsfont(15, 170, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart2(rstEtiqueta.Fields("[Sem Preço]"))))
            Call windowsfont(15, 190, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart3(rstEtiqueta.Fields("[Sem Preço]"))))
            strAuxi = BuscarCorTam(rstEtiqueta.Fields("Código").Value)
            If Len(strAuxi) > 0 Then
                Call windowsfont(15, 210, 22, 0, 0, 0, "arial", strAuxi)
            End If
            Call barcode("15", "240", "128", "50", "1", "0", "2", "2", rstEtiqueta.Fields("Código"))
            Call barcode("15", "350", "128", "40", "1", "0", "2", "2", rstEtiqueta.Fields("Código"))
            Call windowsfont(15, 412, 22, 0, 0, 0, "arial", "R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00"))
            
            rstEtiqueta.MoveNext
           
            ' Etiqueta 2
            If Not rstEtiqueta.EOF Then
                Call windowsfont(305, 150, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart1(rstEtiqueta.Fields("[Sem Preço]"))))
                Call windowsfont(305, 170, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart2(rstEtiqueta.Fields("[Sem Preço]"))))
                Call windowsfont(305, 190, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart3(rstEtiqueta.Fields("[Sem Preço]"))))
                strAuxi = BuscarCorTam(rstEtiqueta.Fields("Código").Value)
                If Len(strAuxi) > 0 Then
                    Call windowsfont(315, 210, 22, 0, 0, 0, "arial", strAuxi)
                End If
                Call barcode("305", "240", "128", "50", "1", "0", "2", "2", rstEtiqueta.Fields("Código"))
                Call barcode("305", "350", "128", "40", "1", "0", "2", "2", rstEtiqueta.Fields("Código"))
                Call windowsfont(305, 412, 22, 0, 0, 0, "arial", "R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00"))

                rstEtiqueta.MoveNext
            End If
           
            If Not rstEtiqueta.EOF Then
                ' Etiqueta 3
                Call windowsfont(595, 150, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart1(rstEtiqueta.Fields("[Sem Preço]"))))
                Call windowsfont(595, 170, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart2(rstEtiqueta.Fields("[Sem Preço]"))))
                Call windowsfont(595, 190, 22, 0, 0, 0, "arial", ConverteCaracter(NomePart3(rstEtiqueta.Fields("[Sem Preço]"))))
                strAuxi = BuscarCorTam(rstEtiqueta.Fields("Código").Value)
                If Len(strAuxi) > 0 Then
                    Call windowsfont(615, 210, 22, 0, 0, 0, "arial", strAuxi)
                End If
                Call barcode("590", "240", "128", "50", "1", "0", "2", "2", rstEtiqueta.Fields("Código"))
                Call barcode("590", "350", "128", "40", "1", "0", "2", "2", rstEtiqueta.Fields("Código"))
                Call windowsfont(595, 412, 22, 0, 0, 0, "arial", "R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00"))
                
                rstEtiqueta.MoveNext
            End If

            Call printlabel("1", "1")
            Call clearbuffer
           
       Loop
      End If
      '---[Fechando a saída]---
  
      Call closeport

      '---[Finalizando Variaveis]---
      rstEtiqueta.Close
      Set rstEtiqueta = Nothing
    
      '''''''    ' primeito teste ok SALVOOOOO
      '''''''    Call setup("105", "55", "3.0", "12", "1", "1", "0")
      '''''''    Call sendcommand("SET TEAR ON")
      '''''''    Call sendcommand("DIRECTION 1")            ' EM Pé ou de Ponta Cabeça (0º ou 180º)
      '''''''    Call clearbuffer
      '''''''    Call windowsfont(15, 0, 30, 0, 0, 0, "arial", "OOOOOOOOOO")
      '''''''    Call barcode("25", "5", "EAN13", "70", "1", "0", "2", "2", "555111111111")
      '''''''    Call windowsfont(315, 0, 30, 0, 0, 0, "arial", "UUUUUUUUUUUUU")
      '''''''    Call barcode("325", "5", "EAN13", "70", "1", "0", "2", "2", "777111111112")
      '''''''    Call windowsfont(615, 0, 30, 0, 0, 0, "arial", "XXXXXXXXXXX")
      '''''''    Call barcode("620", "5", "EAN13", "70", "1", "0", "2", "2", "888111111113")
      '''''''    Call printlabel("1", "1")
      '''''''    ' fim
      
      
  ElseIf opt_modelo_02_argox214.Value = True Then
  
     Dim iConta               As Integer
     Dim sNomeSubClasse       As String
     Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
     Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
     Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
     Set rsProdutosParaEtip = db.OpenRecordset("Produtos", , dbReadOnly)
  
  
     'Adicionado Sub Classe na etiqueta
     rsSubclasses.Index = "Código"
     rsProdutosParaEtip.Index = "Código"
     
     intbuffer = FreeFile
  
     If Combo.Text = "" And Combo.Text = "0" Then
         MsgBox "Selecione um usuário dono da lista", vbInformation, "Atenção"
         Exit Sub
     End If
 
  '''    If arrUsuario(0) = "XX" Then
  '''       Set rstEtiqueta = db.OpenRecordset("SELECT * FROM [Etiquetas - Tempo] ORDER BY Código", dbOpenDynaset)
  '''    Else
         Set rstEtiqueta = db.OpenRecordset("SELECT * FROM [Etiquetas - Tempo] Where Funcionario = " & Combo.Text & " ORDER BY Código", dbOpenDynaset)
  '''    End If

     '----[Porta de Saída]----
     If chk_geraArquivo.Value = vbChecked Then
        sCaminhoAux = Replace(LCase(sCaminhoArquivo), "\\tsclient\c\", "C:") & "\ETIQUETAS"
        sCaminhoAux = Replace(LCase(sCaminhoArquivo), "\\tsclient\d\", "D:") & "\ETIQUETAS"
        MsgBox "Este caminho deve existir na sua máquina. " & sCaminhoAux, vbInformation, "Atenção"
        Open sCaminhoArquivo & "\ETIQUETAS\arquivoEtiquetas_HR" & Hour(Now) & "_" & Minute(Now) & "_DIA" & Day(Now) & "_" & Month(Now) & "_" & Year(Now) & ".p1" For Output As #intbuffer  '...Arquivo texto
     Else
        Open cboSaida.Text For Output As #intbuffer  'abre a saída LPT1 pode ser com COM1...Arquivo texto
     End If

     '---[verificando se há etiquetas para imprimir]---
     If rstEtiqueta.EOF And rstEtiqueta.BOF Then
         MsgBox "Nenhuma lista de Etiqueta Formatada para impressão para este usuário selecionado!" & vbCrLf & vbCrLf & "PASSO 1: Vá até a tela de MONTAR LISTA para criar sua lista de etiquetas." & vbCrLf & vbCrLf & "PASSE 2: Vá até a tela de FORMATAR ETIQUETAS para incluir dados de preço, classe, etc e Formate suas etiquetas.", vbInformation, "Quick Store" ' messagem caso não tenha etiquetas
     Else '
 
        rstEtiqueta.MoveFirst
        
        '----[Configurações Inicias da Impressora ]----
        Print #intbuffer, "O"                   'Seleciona Opções
        Print #intbuffer, "N"                   'Limpa moldura buffer
        Print #intbuffer, "q736"                '92mm * 8 = 736              'Comprimento da etiqueta    Hor.
        Print #intbuffer, "Q200,24+0"           '25(compri) * 8 = 200         3 * 8 = 24 (espaço entre etiq))
        Print #intbuffer, "S2"                  'Fixa a velocidade de impressão
        Print #intbuffer, "D" & CInt(txtTemperatura.Text)
        Print #intbuffer, "ZT"                  'Imprime no topo
        Print #intbuffer, "TTh:m"
        Print #intbuffer, "TDy2.mn.dd"
        
 
        '---[impressão das etiquetas]---
        Do While Not rstEtiqueta.EOF
            iConta = 0
            For iConta = 0 To 1
              If iConta = 0 Then
                  ' Etiqueta da direita
                  Print #intbuffer, "B408,016,0,1,2,2,28,B,""" & rstEtiqueta.Fields("Código").Value & """"
                  Print #intbuffer, "A408,088,0,1,1,1,N,""" & ConverteCaracter(NomePart1_tamanhoVariavel(rstEtiqueta.Fields("[Sem Preço]"), 25)) & """"
                  Print #intbuffer, "A408,108,0,1,1,1,N,""" & ConverteCaracter(NomePart2_tamanhoVariavel(rstEtiqueta.Fields("[Sem Preço]"), 25)) & """"
                  
                  rsProdutosParaEtip.Seek "=", rstEtiqueta("Código Produto")
                  If Not rsProdutosParaEtip.NoMatch Then
                      rsSubclasses.Seek "=", rsProdutosParaEtip.Fields("Sub Classe").Value
                      
                      If Not rsSubclasses.NoMatch Then
                        sNomeSubClasse = Trim(rsSubclasses("Nome"))
                      Else
                        sNomeSubClasse = ""
                      End If
                  Else
                    sNomeSubClasse = ""
                  End If
                 
                  Print #intbuffer, "A408,128,0,1,1,1,N,""" & sNomeSubClasse & """"
                  
                  Print #intbuffer, "A408,148,0,1,1,1,N,""" & rstEtiqueta.Fields("Cor").Value & """"
                  Print #intbuffer, "A580,168,0,2,1,1,N,""" & rstEtiqueta.Fields("Tamanho").Value & """"
                  Print #intbuffer, "A408,168,0,2,1,1,N,""R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00") & """"
              
              Else
                  ' Etiqueta da esquerda
                  Print #intbuffer, "B56,016,0,1,2,2,28,B,""" & rstEtiqueta.Fields("Código").Value & """"
                  Print #intbuffer, "A56,088,0,1,1,1,N,""" & ConverteCaracter(NomePart1_tamanhoVariavel(rstEtiqueta.Fields("[Sem Preço]"), 25)) & """"
                  Print #intbuffer, "A56,108,0,1,1,1,N,""" & ConverteCaracter(NomePart2_tamanhoVariavel(rstEtiqueta.Fields("[Sem Preço]"), 25)) & """"
                  
                  
                  rsProdutosParaEtip.Seek "=", rstEtiqueta("Código Produto")
                  If Not rsProdutosParaEtip.NoMatch Then
                      rsSubclasses.Seek "=", rsProdutosParaEtip.Fields("Sub Classe").Value
                      
                      If Not rsSubclasses.NoMatch Then
                        sNomeSubClasse = Trim(rsSubclasses("Nome"))
                      Else
                        sNomeSubClasse = ""
                      End If
                  Else
                    sNomeSubClasse = ""
                  End If
                  
                  Print #intbuffer, "A56,128,0,1,1,1,N,""" & sNomeSubClasse & """"
                  
                  '''Print #intbuffer, "A56,148,0,1,1,1,N,""" & rstEtiqueta.Fields("Cor").Value & """"
                  '''Print #intbuffer, "A224,168,0,2,1,1,N,""" & rstEtiqueta.Fields("Tamanho").Value & """"
                  '''Print #intbuffer, "A408,168,0,2,1,1,N,""R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00") & """"
                  
                  If chk_colunaDaEsquerdaSemPreco.Value = 1 Then
                      'Imprime a cor e o tamanho na mesma linha
                      Print #intbuffer, "A56,148,0,1,1,1,N,""" & rstEtiqueta.Fields("Cor").Value & """"
                      Print #intbuffer, "A264,148,0,2,1,1,N,""" & rstEtiqueta.Fields("Tamanho").Value & """"
                  Else
                      'Imprime a cor
                      Print #intbuffer, "A56,148,0,1,1,1,N,""" & rstEtiqueta.Fields("Cor").Value & """"
                      'Imprime o preço e o tamanho na mesma linha
                      Print #intbuffer, "A56,168,0,2,1,1,N,""R$ " & Format(RetiraNome(rstEtiqueta.Fields("Descrição")), "#,###,##0.00") & """"
                      Print #intbuffer, "A224,168,0,2,1,1,N,""" & rstEtiqueta.Fields("Tamanho").Value & """"
                  End If
              End If
            Next iConta
            
            rstEtiqueta.MoveNext
            
            Print #intbuffer, "P1"
            Print #intbuffer, "FR"
        Loop
     End If
     
     rsSubclasses.Close
     rsTamanhos.Close
     rsCores.Close
     rsProdutosParaEtip.Close
     Set rsSubclasses = Nothing
     Set rsTamanhos = Nothing
     Set rsCores = Nothing
     Set rsProdutosParaEtip = Nothing
     
     '---[Fechando a saída]---
     Close #intbuffer
  
     '---[Finalizando Variaveis]---
      rstEtiqueta.Close
      Set rstEtiqueta = Nothing
  
  End If
  
  MsgBox "Processo concluído", vbInformation, "Atenção"

  Exit Sub
Erro:
  MsgBox "Erro na função de impressão de etiquetas " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Function ConverteCaracter(strString As String) As String
  Dim intPos As Integer
  Dim strChar As String
  Dim strBuffer As String
  
  If strString = "" Then
        ConverteCaracter = ""
        Exit Function
  End If
  
  For intPos = 1 To Len(strString)
    strChar = Mid(strString, intPos, 1)
    
    Select Case strChar
      Case "ç"
        strChar = "c"
      Case "Ç"
        strChar = "C"
        
      Case "á", "à", "ã", "â", "ä"
        strChar = "a"
      Case "Á", "À", "Ã", "Â", "Ä"
        strChar = "A"
        
      Case "é", "ê", "ë", "è"
        strChar = "e"
      Case "É", "Ê", "Ë", "È"
        strChar = "E"
        
      Case "í", "î", "ï", "ì"
        strChar = "i"
      Case "Í", "Î", "Ï", "Ì"
        strChar = "I"
        
      Case "ó", "ô", "ö", "ò", "õ"
        strChar = "o"
      Case "Ó", "Ô", "Ö", "Ò", "Õ"
        strChar = "O"
        
      Case "ú", "û", "ü", "ù"
        strChar = "u"
      Case "Ú", "Û", "Ü", "Ù"
        strChar = "U"
        
      Case "ñ"
        strChar = "n"
      Case "Ñ"
        strChar = "N"
    End Select
    
    strBuffer = strBuffer & strChar
  Next intPos
  
  ConverteCaracter = strBuffer
End Function

'''Private Function BuscarCorTam(ByVal CodProd As String) As String
'''  Dim rstCodigosDaGrade As Recordset
'''  Dim strSQL            As String
'''  Dim blnGrade          As Boolean
'''
'''  On Error GoTo TratarErro
'''
'''  strSQL = ""
'''  strSQL = "SELECT * FROM [Códigos da Grade] WHERE Código = '" & CodProd & "'"
'''
'''  Set rstCodigosDaGrade = db.OpenRecordset(strSQL, dbOpenDynaset)
'''
'''  If rstCodigosDaGrade.RecordCount = 0 Then
'''     blnGrade = False
'''  Else
'''     blnGrade = True
'''  End If
'''
'''  rstCodigosDaGrade.Close
'''  Set rstCodigosDaGrade = Nothing
'''
'''  If Not blnGrade Then 'Não usa grade
'''    BuscarCorTam = ""
'''  Else                 'Usa grade
'''    Dim rstCor         As Recordset
'''    Dim rstTamanhos    As Recordset
'''    Dim strNomeCor     As String
'''    Dim strNomeTamanho As String
'''
'''    'Busca a Cor
'''    strSQL = ""
'''    strSQL = "SELECT Nome FROM Cores WHERE Código = " & CInt(Right(CodProd, 3))
'''
'''    Set rstCor = db.OpenRecordset(strSQL, dbOpenDynaset)
'''
'''    With rstCor
'''      If Not (.BOF And .EOF) Then
'''        .MoveFirst
'''        strNomeCor = .Fields("Nome").Value & ""
'''
'''      End If
'''      .Close
'''    End With
'''
'''    Set rstCor = Nothing
'''
'''    'Busca o Tam
'''    strSQL = ""
'''    strSQL = "SELECT Nome FROM Tamanhos WHERE Código = " & CInt(Left(Right(CodProd, 6), 3))
'''
'''    Set rstTamanhos = db.OpenRecordset(strSQL, dbOpenDynaset)
'''
'''    With rstTamanhos
'''      If Not (.BOF And .EOF) Then
'''        .MoveFirst
'''        strNomeTamanho = .Fields("Nome").Value & ""
'''
'''      End If
'''      .Close
'''    End With
'''
'''    Set rstTamanhos = Nothing
'''
'''    'BuscarCorTam
'''    BuscarCorTam = ConverteCaracter(Left(strNomeCor & String(15, " "), 15) & " " & Left(strNomeTamanho, 3))
'''
'''  End If
'''
'''  Exit Function
'''
'''TratarErro:
'''  MsgBox "Erro ao buscar o nome da cor/tamanho" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
'''  Err.Clear
'''
'''End Function

Private Function NomePart1(sNome As String) As String
    NomePart1 = Mid(sNome, 1, 19)
End Function
Private Function NomePart2(snome2 As String) As String
    NomePart2 = Mid(snome2, 20, 19)
End Function

Private Function NomePart3(snome3 As String) As String
    NomePart3 = Mid(snome3, 39, 19)
End Function

Private Function NomePart1_tamanhoVariavel(sNome As String, tam As Integer) As String
    NomePart1_tamanhoVariavel = Mid(sNome, 1, tam)
End Function
Private Function NomePart2_tamanhoVariavel(snome2 As String, tam As Integer) As String
    NomePart2_tamanhoVariavel = Mid(snome2, 26, tam)
End Function


Private Function RetiraNome(sDescricao As String) As Double
  Dim intX      As Integer
  Dim strParte  As String
  Dim strTodo   As String
  Dim blnPassou As Boolean
  
  blnPassou = False
  
  For intX = 1 To Len(sDescricao)
    strParte = Mid(sDescricao, intX, 1)
    
    If (strParte = " ") Then
      If (Mid(sDescricao, intX - 1, 1) <> "$") Then
        blnPassou = True
      End If
    End If
    
    If ((IsNumeric(strParte)) Or (strParte = ",")) And (Not blnPassou) Then
      strTodo = strTodo & strParte
    End If
  Next
  
  If IsNumeric(strTodo) Then
    RetiraNome = strTodo
  Else
    RetiraNome = 0
  End If

End Function

Private Function BuscarCorTam(ByVal CodProd As String) As String
  Dim rstCodigosDaGrade As Recordset
  Dim strSQL            As String
  Dim blnGrade          As Boolean
  
  On Error GoTo TratarErro
  
  strSQL = ""
  strSQL = "SELECT * FROM [Códigos da Grade] WHERE Código = '" & CodProd & "'"
  
  Set rstCodigosDaGrade = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstCodigosDaGrade.RecordCount = 0 Then
     blnGrade = False
  Else
     blnGrade = True
  End If
  
  rstCodigosDaGrade.Close
  Set rstCodigosDaGrade = Nothing
  
  If Not blnGrade Then 'Não usa grade
    BuscarCorTam = ""
  Else                 'Usa grade
    Dim rstCor         As Recordset
    Dim rstTamanhos    As Recordset
    Dim strNomeCor     As String
    Dim strNomeTamanho As String
    
    'Busca a Cor
    strSQL = ""
    strSQL = "SELECT Nome FROM Cores WHERE Código = " & CInt(Right(CodProd, 3))
  
    Set rstCor = db.OpenRecordset(strSQL, dbOpenDynaset)
  
    With rstCor
      If Not (.BOF And .EOF) Then
        .MoveFirst
        strNomeCor = .Fields("Nome").Value & ""
        
      End If
      .Close
    End With
  
    Set rstCor = Nothing
  
    'Busca o Tam
    strSQL = ""
    strSQL = "SELECT Nome FROM Tamanhos WHERE Código = " & CInt(Left(Right(CodProd, 6), 3))
  
    Set rstTamanhos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
    With rstTamanhos
      If Not (.BOF And .EOF) Then
        .MoveFirst
        strNomeTamanho = .Fields("Nome").Value & ""
        
      End If
      .Close
    End With
  
    Set rstTamanhos = Nothing
  
    'BuscarCorTam
    BuscarCorTam = ConverteCaracter(Left(strNomeCor & String(15, " "), 15) & " " & Left(strNomeTamanho, 3))
  
  End If
  
  Exit Function

TratarErro:
  MsgBox "Erro ao buscar o nome da cor/tamanho" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  Err.Clear
  
End Function

'''Private Sub cmd_limparTabelaDeEtiquetas_Click()
'''On Error GoTo Erro
'''
'''  If MsgBox("Deseja realmente apagar a lista de produtos e quantidades de etiquetas para impressão que você criou?", vbQuestion + vbYesNo) = vbYes Then
'''      Call StatusMsg("Apagando a tabela com produtos e quantidades de etiquetas para impressão...")
'''      MousePointer = vbHourglass
'''
'''      If Len(Nome_func.Caption) > 0 Then
'''          db.Execute "DELETE * FROM Etiquetas WHERE Funcionário = " & CInt(Combo.Text)
'''          MsgBox "Tabela apagada com sucesso", vbInformation, "Sucesso"
'''      Else
'''          If MsgBox("Você esta optou por apagar a lista com produtos para todos os usuários. Deseja continuar e apagar?", vbQuestion + vbYesNo) = vbYes Then
'''              db.Execute "DELETE * FROM Etiquetas"
'''
'''              MsgBox "Tabela apagada com sucesso", vbInformation, "Sucesso"
'''          End If
'''      End If
'''
'''      Call StatusMsg("")
'''      MousePointer = vbDefault
'''  End If
'''
'''  Exit Sub
'''Erro:
'''    MsgBox "Erro ao realizar a limpeza/exclusão da tabela de etiquetas para impressão " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
'''
'''End Sub

Private Sub O_Grande_Click()
  
  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequenaIMPRESSA.bmp")
  
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False

  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_Grande2_Click()
  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False

  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_Grande3_Click()
  
  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequenaIMPRESSA.bmp")
  
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False

  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_Grande4_Click()
  
  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequenaIMPRESSA.bmp")
  
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = True
  cboPosicao.Enabled = True

  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_GrandeProcon_Click()
  
  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6182_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6182_pequenaIMPRESSA.bmp")


  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False
  
  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Value = 1
  Imprime_preço.Enabled = False
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_GrandeProcon2_Click()

  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimacoXXXX_pequenaIMPRESSA.bmp")

  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (True)

End Sub

Private Sub O_Média_Click()
  
  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6180_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6180_pequenaIMPRESSA.bmp")
  
  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False
  
  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_Pequena_Click()

  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6187_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6187_pequenaIMPRESSA.bmp")


  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False

  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub O_Pequena2_Click()
  '30/10/2007 - Celso
  '
  'Solicitante..: Jefferson
  '
  'Finalidade...: Atender solicitação de cliente
  '
  
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False

  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)

End Sub

Private Sub opt_modelo_01_Click()

  chk_colunaDaEsquerdaSemPreco.Visible = False
  cmd_imagem3.Picture = LoadPicture(App.Path & "\Imagens\etiquetaArgoxPPLB_01_etiqueta_pequena.bmp")
  cmd_imagem4.Picture = LoadPicture(App.Path & "\Imagens\etiquetaArgoxPPLB_01_pequena.bmp")
  
  cboSaida.Text = "LPT1"

End Sub

Private Sub opt_modelo_02_argox214_Click()
  If opt_modelo_02_argox214.Value = True Then
      chk_colunaDaEsquerdaSemPreco.Visible = True
  Else
      chk_colunaDaEsquerdaSemPreco.Visible = False
  End If
  
  cmd_imagem3.Picture = LoadPicture(App.Path & "\Imagens\etiquetaArgoxPPLA_02_etiqueta_pequena.bmp")
  cmd_imagem4.Picture = LoadPicture(App.Path & "\Imagens\etiquetaBematechLB1000_01_pequena.bmp")
  
End Sub

Private Sub opt_zpl_epl_bematechLB1000_Click()

  chk_colunaDaEsquerdaSemPreco.Visible = False
  cmd_imagem3.Picture = LoadPicture(App.Path & "\Imagens\etiquetaBematechLB1000_01_etiqueta_pequena.bmp")
  cmd_imagem4.Picture = LoadPicture(App.Path & "\Imagens\etiquetaBematechLB1000_01_pequena.bmp")
  
  cboSaida.Text = "USB"

End Sub

'18/06/2009 - mpdea
Private Sub optMedia6081_GotFocus()
  O_Imprime_Barras.Value = True
  O_Imprime_Pequena.Enabled = False
  O_Não_Imprime.Enabled = False
  
  cmd_imagem1.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6081_pequena.bmp")
  cmd_imagem2.Picture = LoadPicture(App.Path & "\Imagens\etiquetaPimaco6081_pequenaIMPRESSA.bmp")
  
  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False
  
  '18/04/2007 - Anderson
  'Implementação da etiqueta para atender exigencias do procon
  Imprime_preço.Enabled = True
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)
  
End Sub

'18/06/2009 - mpdea
Private Sub optMedia6081_LostFocus()
  O_Imprime_Pequena.Enabled = True
  O_Não_Imprime.Enabled = True
End Sub

Private Sub sldEsquerda_Change()
  Call sldEsquerda_Click
End Sub

Private Sub sldEsquerda_Click()
  Ajusta_Ver = nCheckValues(tmEsquerda)
End Sub

Private Sub sldEsquerda_Scroll()
  Call sldEsquerda_Click
End Sub

Private Sub sldSuperior_Change()
  Call sldSuperior_Click
End Sub

Private Sub sldSuperior_Click()
  Ajusta_Hor = nCheckValues(tmSuperior)
End Sub

Private Sub sldSuperior_Scroll()
  Call sldSuperior_Click
End Sub

Private Sub B_Emite_Click()
  Dim sSql As String
  Dim Linhas As Long
  Dim sCodFunc As Integer
  Dim Prod As String
  Dim Tamanho As Integer
  Dim Tamanho_Str As String
  Dim Cor As Integer
  Dim Cor_Str As String
  Dim Erro As Integer
  Dim Descrição As String
  Dim Sem_Preço As String
  Dim i As Long
  Dim Str1 As String
  Dim Cod_Str As String
  Dim Cod_Cor_Str As String
  Dim Cod_Tam_Str As String
  Dim sObs As String
  '17/02/2005 - Daniel
  Dim strDescricao2 As String
  '20/05/2005 - Daniel
  'Tratamento de novos campos
  Dim blnImprimirUmaEtiq   As Boolean
  Dim blnImprimirPrecoEtiq As Boolean
  '08/06/2005 - Daniel
  'Tratamento para o campo [Etiquetas - Tempo].Funcionario
  Dim intFuncionario As Integer
  '30/10/2007 - Celso
  'Modelo novo de etiqueta com lote e validade
  Dim Lote As String
  Dim DataValidade As String
  
  '19/02/2008 - Celso
  'Implementação PROCON
  Dim nPreco_Vista As Double
  Dim nPreco_Prazo As Double
  Dim sTipo As String
  
  
  '----------------[Fim das declarações de variáveis]----------------
  
  On Error GoTo TratarErro '31/03/2005 - Daniel (Tratamento para erros)
  
  Call StatusMsg("")
  
  If Combo.Text = "" And Combo.Text = "0" Then
      MsgBox "Selecione um usuário dono da lista", vbInformation, "Atenção"
      Exit Sub
  End If
  
  If O_Pequena.Value = True Then
    Margem_Sup = 720  '1.27 cm * 567 twips por cm
    Margem_Inf = Margem_Sup
    Margem_Esq = 794 '1.40 cm * 567 twips por com
    Margem_Dir = Margem_Esq
  End If
  
  '------------------------------------------------------------------
  '30/10/2007 - Celso
  If O_Pequena2.Value = True Then
    Margem_Sup = 720  '1.27 cm * 567 twips por cm
    Margem_Inf = Margem_Sup
    Margem_Esq = 794 '1.40 cm * 567 twips por com
    Margem_Dir = Margem_Esq
  End If
  '------------------------------------------------------------------
  
  If O_Média.Value = True Then
    Margem_Sup = 765  '1.35 cm * 567 twips por cm
    Margem_Inf = 624  '1.10 * 567
    Margem_Esq = 600  '1.10 cm * 567 twips por com
    Margem_Dir = 100  '1.00 * 567
  End If
  
  '17/06/2009 - mpdea
  If optMedia6081.Value Then
    Margem_Sup = 720  '1.27 cm * 567 twips por cm
    Margem_Inf = Margem_Sup
    Margem_Esq = 213  '0.377 cm * 567 twips por cm (231,759)
    Margem_Dir = Margem_Esq
  End If
  
  If O_Grande2.Value = True Then 'grande com descrição
    Margem_Sup = 1304  '2.30 cm * 567 twips por cm
    Margem_Inf = 1021  '1.80 * 567
    Margem_Esq = 1049  '1.85 cm * 567 twips por com
    Margem_Dir = 964   '1.70 * 567
  End If
  
  If O_Grande.Value = True Or O_Grande3.Value = True Then  'roupa
    Margem_Sup = 680   '1.20 cm * 567 twips por cm
    Margem_Inf = 567   '1.00 * 567
    Margem_Esq = 567   '1.00 cm * 567 twips por com
    Margem_Dir = 567   '1.00 * 567
  End If
  
  '24/04/2007 - Anderson
  'Descontinuado
  '18/04/2007 - Anderson
  'Implementação de etiqueta grande com exigências do PROCON
  'If O_GrandeProcon.Value = True Then
    'Margem_Sup = 1202   '2.12 cm * 567 twips por cm
    'Margem_Inf = 1202   '2.12 cm * 567
    'Margem_Esq = 226    '0.40 cm * 567 twips por com
    'Margem_Dir = 226    '0.40 * 567
  'End If
    
  '19/02/2008 - Celso
  'Implementação PROCON
  If O_GrandeProcon2.Value = True Then
    Margem_Sup = 720    '1.27 cm * 567 twips por cm
    Margem_Inf = 720    '1.27 cm * 567
    Margem_Esq = 181    '0.32 cm * 567 twips por com
    Margem_Dir = 181    '0.32 * 567
  End If
    
  'Verifica opções selecionadas
  If Imprime_preço.Value = 1 Then
    If Lista.Text = "" Then
      DisplayMsg "Lista de preços incorreta, verifique."
      ' Lista.SetFocus
      Exit Sub
    End If
  End If
  
  If O_Imprime_Barras.Value = False Then
    If Imprime_preço.Value = False Then
      If Imprime_Nome.Value = False Then
        DisplayMsg "Escolha pelo menos uma opção de impressão (nome, preço, código)."
        Exit Sub
      End If
    End If
  End If
  
  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  If O_Grande4.Value Then
    If Not IsNumeric(cboPosicao.Text) Then
      MsgBox "Qtde. de caracteres para o Cód. de barras inválida, verifique.", vbExclamation, "Atenção"
      cboPosicao.SetFocus
      Exit Sub
    End If
  End If
  '--------------------------------------------------------
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  If O_GrandeProcon2 Then
    If Not IsNumeric(Num_Parcelas.Text) Then
      MsgBox "Número de parcelas deve ser numérico!", vbExclamation, "Atenção"
      Num_Parcelas.SetFocus
      Exit Sub
    End If
    If CInt(Num_Parcelas.Text) < 1 Or CInt(Num_Parcelas.Text) > 36 Then
      MsgBox "Número de parcelas deve ser entre 1 e 36. Verifique!", vbExclamation, "Atenção"
      Num_Parcelas.SetFocus
      Exit Sub
    End If
    If Tab_Vista.Text = "" Then
      DisplayMsg "Tabela de preços incorreta. Verifique!"
      Tab_Vista.SetFocus
      Exit Sub
    End If
    If Tab_Prazo.Text = "" Then
      DisplayMsg "Tabela de preços incorreta. Verifique!"
      Tab_Prazo.SetFocus
      Exit Sub
    End If
  End If
  '--------------------------------------------------
  
  'Apaga o que tiver no arquivo Etiquetas - Tempo
  If Combo.Text <> "" And Combo.Text <> "0" Then
      sSql = "Delete * From [Etiquetas - Tempo] where Funcionario = " & Combo.Text
      db.Execute sSql
  End If
  
  
  Dim iContAux As Integer
  Dim sSqlAux As String
  If CInt(txt_etiquetasEmBranco.Text) > 1 Then
    For iContAux = 1 To CInt(txt_etiquetasEmBranco.Text) - 1
        ws.BeginTrans
        sSqlAux = "Insert into [Etiquetas - Tempo] (Código, [Código Barra], [Código Produto], Descrição, [Sem Preço], Preco, "
        sSqlAux = sSqlAux & "Descricao2, Funcionario, DividirPrecoEtiqueta) "
        sSqlAux = sSqlAux & " values ('00','*00*','00','00','00',0,'00'," & Combo.Text & ",1) "
        
        db.Execute sSqlAux
        ws.CommitTrans
    Next
  End If
    
  
  'Coloca as Etiquetas do Funcionário escolhido
  'no arquivo Etiquetas - Tempo
  Call StatusMsg("Aguarde, gerando etiquetas...")
  B_Emite.Enabled = False
  DoEvents
  
  Set rsEtiquetas_Tempo = db.OpenRecordset("Etiquetas - Tempo")
  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  
  '04/11/2005 - mpdea
  'Adicionado Sub Classe na etiqueta
  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
  
  rsPreços.Index = "Tabela"
  rsProdutos.Index = "Código"
  rsTamanhos.Index = "Código"
  rsCores.Index = "Código"
  rsClasses.Index = "Código"
  
  '04/11/2005 - mpdea
  'Adicionado Sub Classe na etiqueta
  rsSubclasses.Index = "Código"
  
  sCodFunc = CInt(gsHandleNull(Combo.Text))
  Prod = ""
  Erro = False
  
  If Val(sCodFunc) > 0 Then
    Set rsEtiquetas = db.OpenRecordset("SELECT * FROM Etiquetas WHERE Funcionário = " & sCodFunc & " ORDER BY Funcionário, Produto, Tamanho, Cor", dbOpenDynaset)
  Else
    Set rsEtiquetas = db.OpenRecordset("SELECT * FROM Etiquetas ORDER BY Funcionário, Produto, Tamanho, Cor", dbOpenDynaset)
  End If
  
 
  Do While Not rsEtiquetas.EOF
    Prod = rsEtiquetas("Produto")
    Tamanho = rsEtiquetas("Tamanho")
    Cor = rsEtiquetas("Cor")
    Tamanho_Str = ""
    Cor_Str = ""
    If Cor <> 0 Then
      If O_Cor.Value = 1 Then
        rsCores.Seek "=", Cor
        If Not rsCores.NoMatch Then Cor_Str = rsCores("Nome")
      End If
    End If
    If Tamanho <> 0 Then
      If O_Tamanho.Value = 1 Then
        rsTamanhos.Seek "=", Tamanho
        If Not rsTamanhos.NoMatch Then Tamanho_Str = rsTamanhos("Nome")
      End If
    End If
  
    '  If O_Grande = False Then 'não é etiqueta de roupa
    Descrição = ""
    Sem_Preço = ""
    sObs = ""
    
    If Imprime_preço.Value = 1 Then
      rsPreços.Seek "=", Lista.Text, rsEtiquetas("Produto")
      '19/04/2007 - Anderson
      'Incluído clausula para verificar se é etiqueta padrão Procon
      'If Not rsPreços.NoMatch Then
      If Not rsPreços.NoMatch And Not O_GrandeProcon Then
        Descrição = Descrição + "R$ " + Format$(rsPreços("Preço"), "###,###,##0.00") + " "
      End If
    End If
   
    If Imprime_Nome.Value = 1 Then
       rsProdutos.Seek "=", rsEtiquetas("Produto")
       If Not rsProdutos.NoMatch Then
          Descrição = Descrição + rsProdutos("Nome")
          Sem_Preço = rsProdutos("Nome")
          sObs = rsProdutos("Obs") & ""
          If chkAppendClasse.Value = 1 Then
            rsClasses.Seek "=", rsProdutos("Classe")
            If Not rsClasses.NoMatch Then
              Descrição = Trim(Descrição) & " " & Trim(rsClasses("Nome"))
            End If
          End If
          '04/11/2005 - mpdea
          'Adicionado Sub Classe na etiqueta
          If chkAppendSubClasse.Value = 1 Then
            rsSubclasses.Seek "=", rsProdutos("Sub Classe")
            If Not rsSubclasses.NoMatch Then
              Descrição = Trim(Descrição) & " " & Trim(rsSubclasses("Nome"))
            End If
          End If
       Else
          Descrição = rsEtiquetas("Produto")
          Sem_Preço = ""
          sObs = ""
       End If
    End If
    
    If chk_fonteNomeLowerCase.Value = vbChecked Then
        Descrição = LCase(Descrição)
        Descrição = Replace(Descrição, "r$", "R$")
    End If
  
    '17/02/2005 - Daniel
    '
    'Tratamento para o campo [Etiquetas - Tempo].Descricao2
    'Código de barras
    strDescricao2 = rsProdutos("Nome") & ""
    '
    '20/05/2005 - Daniel
    'Tratamento para os campos ImprimirUmaEtiq e ImprimirPrecoEtiq
    blnImprimirUmaEtiq = rsProdutos("ImprimirUmaEtiq").Value
    blnImprimirPrecoEtiq = rsProdutos("ImprimirPrecoEtiq").Value
    '
    '08/06/2005 - Daniel
    'Tratamento para o campo [Etiquetas - Tempo].Funcionario
    If Len(Nome_func.Caption) > 0 Then 'Funcionário preenchido
      intFuncionario = CInt(Combo.Text)
    Else
      intFuncionario = 0
    End If
    
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
    If O_GrandeProcon2.Value Then
      rsPreços.Seek "=", Tab_Vista.Text, rsEtiquetas("Produto")
      If Not rsPreços.NoMatch Then
         nPreco_Vista = CDbl(rsPreços.Fields("Preço"))
      Else
         nPreco_Vista = 0
      End If
      rsPreços.Seek "=", Tab_Prazo.Text, rsEtiquetas("Produto")
      If Not rsPreços.NoMatch Then
         nPreco_Prazo = CDbl(rsPreços.Fields("Preço"))
      Else
         nPreco_Prazo = 0
      End If
    End If
    
    If Parcelamento.Value = True Then
      sTipo = "Tipo ='1'"   'Parcelamento
    Else
      sTipo = "Tipo ='2'"   'Financiamento
    End If

  '----------------------------------------------

    For i = 1 To rsEtiquetas("Qtde")
      Cod_Str = rsEtiquetas("Produto")
      Cod_Str = Trim(Cod_Str)
      
      Cod_Cor_Str = ""
      If rsEtiquetas("Cor") <> 0 Then
        Cod_Cor_Str = "000" + Trim(str(rsEtiquetas("Cor")))
        Cod_Cor_Str = Right$(Cod_Cor_Str, 3)
      End If
      
      Cod_Tam_Str = ""
      If rsEtiquetas("Tamanho") <> 0 Then
        Cod_Tam_Str = "000" + Trim(str(rsEtiquetas("Tamanho")))
        Cod_Tam_Str = Right$(Cod_Tam_Str, 3)
      End If
        
      Cod_Str = Cod_Str + Cod_Tam_Str + Cod_Cor_Str
      
      With rsEtiquetas_Tempo
        .AddNew
        .Fields("Código") = Cod_Str
        .Fields("Código Barra") = "*" + UCase(Cod_Str) + "*"
        .Fields("Código Produto") = rsEtiquetas("Produto")
        .Fields("Sem Preço") = Left(Sem_Preço, .Fields("Sem Preço").Size)
        
        '16/08/2005 - mpdea
        'Corrigido RT-3021
        'Incluído verificação de preço localizado
        If Not rsPreços.NoMatch Then
          .Fields("Preco") = CDbl(rsPreços.Fields("Preço"))
        End If
        
        .Fields("Tamanho") = Left(Tamanho_Str, .Fields("Tamanho").Size)
        .Fields("Cor") = Left(Cor_Str, .Fields("Cor").Size)
        .Fields("Texto Grande") = sObs
        '23/04/2007 - Anderson
        'Incluída geração de código completo quando selecionado etiqueta modelo Procon sem impressão do código de barras
        If O_Não_Imprime And O_GrandeProcon Then
          Descrição = Cod_Str & " - " & Descrição
        End If
        
        If O_Pequena And chk_pequena_fonteNomeMenor.Value = vbChecked Then
            If Len(.Fields("Tamanho")) > 0 And O_Tamanho.Value Then
              Descrição = Descrição & " " & LCase(.Fields("Tamanho"))
            End If
            If Len(.Fields("Cor")) > 0 And O_Cor.Value Then
              Descrição = Descrição & "  " & LCase(.Fields("Cor"))
            End If
        End If
        
        If O_GrandeProcon Then
            If Len(.Fields("Tamanho")) > 0 And O_Tamanho.Value Then
              Descrição = Descrição & " - TAM: " & .Fields("Tamanho")
            End If
            If Len(.Fields("Cor")) > 0 And O_Cor.Value Then
              Descrição = Descrição & " - COR: " & .Fields("Cor")
            End If
        End If
        .Fields("Descrição") = Left(Descrição, .Fields("Descrição").Size)
        '22/05/2007 - Anderson
        'Incluído o código para evitar a repetição do código do produto na descrição
        If O_Não_Imprime And O_GrandeProcon Then
            Descrição = Mid(Descrição, Len(Cod_Str & " - ") + 1)
        End If
        '17/02/2005 - Daniel
        '
        'Tratamento para o campo [Etiquetas - Tempo].Descricao2
        'neste campo será armazenado apenas o Nome do Produto que
        'terá finalidade de atender etiquetas com 09 caracteres no
        'código de barras
        .Fields("Descricao2").Value = Left(strDescricao2 & "", .Fields("Descricao2").Size)
        '20/05/2005 - Daniel
        'Tratamento para os campos ImprimirUmaEtiq e ImprimirPrecoEtiq
        .Fields("ImprimirUmaEtiq").Value = blnImprimirUmaEtiq
        .Fields("ImprimirPrecoEtiq").Value = blnImprimirPrecoEtiq
        '08/06/2005 - Daniel
        'Tratamento para o campo [Etiquetas - Tempo].Funcionario
        .Fields("Funcionario").Value = intFuncionario
        '-------------------------------------------------------------
        
        '19/04/2007 - Anderson
        'Implementação da divisão do preço na etiqueta atendendo as exigências do PROCON
        .Fields("DividirPrecoEtiqueta").Value = rsProdutos("DividirPrecoEtiqueta")
        '30/10/2007 - Celso
        'Modelo novo de etiqueta com lote e data validade
        .Fields("Lote").Value = rsProdutos("Lote") + " "
        .Fields("DataValidade").Value = rsProdutos("DataValidade")
        
        '19/02/2008 - Celso
        'Implementação PROCON
        If O_GrandeProcon2.Value Then
          .Fields("Preco") = nPreco_Vista
          .Fields("PrecoPrazo") = nPreco_Prazo
        End If
        
        .Update
      End With
    Next i
    
    rsEtiquetas.MoveNext
  Loop
  
'''  'Seta Valores e Manda Relatório
'''
'''  '31/03/2005 - Daniel
'''  'Resetando o objeto Rel
'''  'Estava ocorrendo o erro do Crystal "Error opening file"
'''  'File could not be found "Etiquetas - Tempo", at file location "Etiquetas - Tempo"
'''  'e em seguida estava sendo tratado o erro através do TratarErro e era exibida a
'''  'mensagem: Run-time error '20533' Unable to open database
'''  Rel.Reset
'''  '--------------------------------------------------------------------------------
'''
'''  '04/04/2005 - Daniel
'''  'Correção para evitar a mensagem: Nº de tabelas inválidas
'''  If O_Grande4.Value Then
'''    'Nome do BD
'''    Str1 = gsQuickDBFileName
'''    Rel.DataFiles(0) = Str1
'''    '31/03/2005 - Daniel
'''    'Adicionado linha abaixo para evitar bug
'''    Rel.DataFiles(1) = Str1
'''  Else
'''    'Nome do BD
'''    Str1 = gsQuickDBFileName
'''    Rel.DataFiles(0) = Str1
'''  End If
'''
'''  'Saída
'''  If O_Vídeo = True Then Rel.Destination = 0
'''  If O_Impressora = True Then Rel.Destination = 1
'''
'''  'Estado da janela
'''  Rel.WindowState = crptMaximized
'''
'''  '12/05/2005 - Daniel
'''  'Correção para exibição dos botões de Configuração
'''  'de Impressoras e Botão de Pesquisas
'''  Rel.WindowShowPrintSetupBtn = True
'''  Rel.WindowShowSearchBtn = True
'''
'''  'Nome do relatório
'''  If O_Imprime_Barras.Value Then
'''    If O_Pequena.Value Then
'''      Str1 = gsReportPath & "ETIQP1.RPT"
'''    ElseIf O_Pequena2.Value Then ' 30/10/2007 - Celso - Novo modelo etiqueta
'''      Str1 = gsReportPath & "ETIQP4.RPT"
'''    ElseIf O_Média.Value Then
'''      '07/03/2003 - Maikel Cordeiro
'''      'Adicionadas as clausulas abaixo para a etiqueta média com tamanho
'''      '---------------------------------------------------------------
'''      If O_Tamanho.Value = vbChecked Then
'''        Str1 = gsReportPath & "ETIQM1T.RPT"
'''      Else
'''        Str1 = gsReportPath & "ETIQM1.RPT"
'''      End If
'''      '---------------------------------------------------------------
'''    ElseIf optMedia6081.Value Then
'''      '17/06/2009 - mpdea
'''      Str1 = gsReportPath & "Etiq6081.RPT"
'''    ElseIf O_Grande.Value Then
'''      Str1 = gsReportPath & "ETIQROUP.RPT"
'''    ElseIf O_Grande3.Value Then
'''      Str1 = gsReportPath & "ETIQROU2.RPT"
'''    ElseIf O_Grande2.Value Then
'''      Str1 = gsReportPath & "ETIQG1.RPT"
'''    ElseIf O_GrandeProcon.Value Then '19/04/2007 - Anderson - Implementação de etiqueta - Exigência Procon
'''      Str1 = gsReportPath & "ETIQG4.RPT"
'''    End If
'''  ElseIf O_Imprime_Pequena.Value Then
'''    If O_Pequena.Value Then
'''      Str1 = gsReportPath & "ETIQP3.RPT"
'''    ElseIf O_Pequena2.Value Then ' 30/10/2007 - Celso - Novo modelo etiqueta
'''      Str1 = gsReportPath & "ETIQP4.RPT"
'''    ElseIf O_Média.Value Then
'''      '07/03/2003 - Maikel Cordeiro
'''      'Adicionadas as clausulas abaixo para a etiqueta média com tamanho
'''      '---------------------------------------------------------------
'''      If O_Tamanho.Value = vbChecked Then
'''        Str1 = gsReportPath & "ETIQM3T.RPT"
'''      Else
'''        Str1 = gsReportPath & "ETIQM3.RPT"
'''      End If
'''    ElseIf O_Grande.Value Then
'''      Str1 = gsReportPath & "ETIQROU3.RPT"
'''    ElseIf O_Grande3.Value Then
'''      Str1 = gsReportPath & "ETIQROU4.RPT"
'''    ElseIf O_Grande2.Value Then
'''      Str1 = gsReportPath & "ETIQG3.RPT"
'''    ElseIf O_GrandeProcon.Value Then '19/04/2007 - Anderson - Implementação de etiqueta - Exigência Procon
'''      Str1 = gsReportPath & "ETIQG5.RPT"
'''    End If
'''  ElseIf O_Não_Imprime.Value Then
'''    If O_Pequena.Value Then
'''      Str1 = gsReportPath & "ETIQP2.RPT"
'''    ElseIf O_Pequena2.Value Then ' 30/10/2007 - Celso - Novo modelo etiqueta
'''      Str1 = gsReportPath & "ETIQP4.RPT"
'''
'''    ElseIf O_Média.Value Then
'''      '07/03/2003 - Maikel Cordeiro
'''      'Adicionadas as clausulas abaixo para a etiqueta média com tamanho
'''      '---------------------------------------------------------------
'''      If O_Tamanho.Value = vbChecked Then
'''        Str1 = gsReportPath & "ETIQM2T.RPT"
'''      Else
'''        Str1 = gsReportPath & "ETIQM2.RPT"
'''      End If
'''    ElseIf O_Grande.Value Then
'''      Str1 = gsReportPath & "ETIQROUP.RPT"
'''    ElseIf O_Grande3.Value Then
'''      Str1 = gsReportPath & "ETIQROU2.RPT"
'''    ElseIf O_Grande2.Value Then
'''      Str1 = gsReportPath & "ETIQG2.RPT"
'''    ElseIf O_GrandeProcon.Value Then '19/04/2007 - Anderson - Implementação de etiqueta - Exigência Procon
'''      Str1 = gsReportPath & "ETIQG6.RPT"
'''    End If
'''  End If
'''
'''  If O_Grande4.Value Then
'''    '16/02/2005 - Daniel
'''    '
'''    'Solicitante..: Consultora Marineida
'''    '
'''    'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
'''    '
'''    'Tratamento para Qtde. de caracteres a serem impressas no
'''    'Código de barras
'''    If cboPosicao.Text = "12" Then
'''      Str1 = gsReportPath & "EtiqRou5.rpt"
'''    Else
'''      Str1 = gsReportPath & "EtiqRou5B.rpt"
'''    End If
'''  End If
'''
'''  If O_GrandeProcon2.Value Then '20/02/2008 - Celso - Nova Implementação PROCON
'''     Str1 = gsReportPath & "ETIQG7.RPT"
'''  End If
'''
'''  Rel.ReportFileName = Str1
'''
'''  Rel.Formulas(0) = ""
'''  Rel.Formulas(1) = ""
'''  Rel.Formulas(2) = ""
'''
'''  '17/06/2009 - mpdea
'''  If optMedia6081.Value Then
'''    Str1 = "Mensagem = '" + (rsParametros("Mensagem Troca") & "") + "'"
'''    If O_Troca = 0 Then Str1 = "Mensagem = ''"
'''    Rel.Formulas(0) = Str1
'''  End If
'''
'''  '06/10/2004 - Daniel
'''  'Adicionado a linha: Or O_Grande4.Value
'''  'Case: Paulimaq não imprimia nunca a mensagem de troca
'''  If O_Grande.Value = True Or O_Grande3.Value = True Or O_Grande4.Value Then
'''    Str1 = "Mensagem = '" + (rsParametros("Mensagem Troca") & "") + "'"
'''    If O_Troca = 0 Then Str1 = "Mensagem = ''"
'''    Rel.Formulas(0) = Str1
'''    Str1 = "Fone1 = '" + (rsParametros("Mensagem Etiq 1") & "") + "'"
'''    Rel.Formulas(1) = Str1
'''    Str1 = "Fone2 = '" + (rsParametros("Mensagem Etiq 2") & "") + "'"
'''    Rel.Formulas(2) = Str1
'''  End If
'''
'''  '19/02/2008 - Celso
'''  'Implementação PROCON
'''  If O_GrandeProcon2.Value = True Then
'''    Rel.Formulas(0) = sTipo
'''    Rel.Formulas(1) = "Msg1 ='" & Msg_Linha1.Text & "" & "'"
'''    Rel.Formulas(2) = "Msg2 ='" & Msg_Linha2.Text & "" & "'"
'''    Rel.Formulas(3) = "NumParc ='" & Num_Parcelas.Text & "" & "'"
'''  End If
'''
'''  'Margens
'''  Margem_Sup = Margem_Sup + Ajusta_Hor
'''  Margem_Inf = Margem_Inf - Ajusta_Hor
'''
'''  Margem_Esq = Margem_Esq + Ajusta_Ver
'''  Margem_Dir = Margem_Dir - Ajusta_Ver
'''
'''  Rel.MarginTop = Margem_Sup
'''  Rel.MarginBottom = Margem_Inf
'''  Rel.MarginLeft = Margem_Esq
'''  Rel.MarginRight = Margem_Dir
'''
'''  Call StatusMsg("Aguarde, imprimindo...")
'''  MousePointer = vbHourglass
'''
'''  '25/07/2003 - mpdea
'''  'Seta a impressora para relatório
'''  Call SetPrinterName("REL", Rel)
'''
'''
'''  Rel.Action = 1

  rsEtiquetas_Tempo.Close
  rsTamanhos.Close
  rsCores.Close
  rsEtiquetas.Close
  Set rsEtiquetas_Tempo = Nothing
  Set rsTamanhos = Nothing
  Set rsCores = Nothing
  Set rsEtiquetas = Nothing
  
  '04/11/2005 - mpdea
  'Fecha tabelas
  rsClasses.Close
  rsSubclasses.Close
  Set rsClasses = Nothing
  Set rsSubclasses = Nothing
  

  
  Call StatusMsg("")
  B_Emite.Enabled = True
  MousePointer = vbDefault
  
  MsgBox "Lista de Etiquetas salvo com sucesso!", vbInformation, "Sucesso"

  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub CarregarGrid()
On Error GoTo Erro
    
    Dim rsEtiquetas As Recordset

    Set rsEtiquetas = db.OpenRecordset("SELECT [Código Produto], Descrição, Preco FROM [Etiquetas - Tempo] ORDER BY Código", dbOpenDynaset)

    If Not (rsEtiquetas.EOF And rsEtiquetas.BOF) Then
        rsEtiquetas.MoveFirst
    End If
  
    grid1.Rows = 1
  
    While Not rsEtiquetas.EOF
  
        If Not IsNull(rsEtiquetas.Fields(2).Value) Then
            grid1.AddItem rsEtiquetas.Fields(0).Value & vbTab & _
                          rsEtiquetas.Fields(1).Value & vbTab & _
                          FormataValorTexto(rsEtiquetas.Fields(2).Value, 2)
        Else
            grid1.AddItem rsEtiquetas.Fields(0).Value & vbTab & _
                          rsEtiquetas.Fields(1).Value & vbTab & _
                          "0,00"
        
        End If
      
        rsEtiquetas.MoveNext
    Wend
    rsEtiquetas.Close
    Set rsEtiquetas = Nothing

    Exit Sub
Erro:
    MsgBox "Erro na carga da grid " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Combo_CloseUp()
 Combo.Text = Combo.Columns(2).Text
 Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
   Nome_func.Caption = ""
   If IsNull(Combo.Text) Or Combo.Text = "" Then Exit Sub
   If Not IsNumeric(Combo.Text) Then Exit Sub
   If Val(Combo.Text) < 0 Then Exit Sub
   If Val(Combo.Text) > 9999 Then Exit Sub

   rsFuncionarios.Index = "Código"
   rsFuncionarios.Seek "=", Val(Combo.Text)
   If rsFuncionarios.NoMatch Then Exit Sub
   Nome_func.Caption = rsFuncionarios("Nome")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10008)
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

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
  
End Function

Private Sub Form_Load()

  grid1.ColWidth(0) = 2000
  grid1.ColWidth(1) = 7500
  grid1.ColWidth(2) = 1500
  
  grid1.Row = 0
  grid1.TextMatrix(0, 0) = "Código Produto"
  grid1.TextMatrix(0, 1) = "Nome"
  grid1.TextMatrix(0, 2) = "Preço"

  Dim Últ_Tabela As String
  Dim Lugar As Integer
  
  ' 19/02/2008 - Celso
  ' Implementação para atender normas do PROCON - referentes a preços a vista e a prazo
  nTrataDetalhesPROCON (False)
  
  '22/01/2003 - mpdea
  'Modo limitado
  If Not gblnQuickFull Then
    Frame5.Visible = False
    O_Não_Imprime.Value = True
  End If
  O_Tamanho.Visible = gbGrade
  O_Cor.Visible = gbGrade
  
  
  Call CenterForm(Me)
  
  Data1.DatabaseName = gsQuickDBFileName

  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  '05/08/2005 - Daniel
  chkAppendClasse.Enabled = True
  
  '16/02/2005 - Daniel
  '
  'Solicitante..: Consultora Marineida
  '
  'Finalidade...: Atender clientes como a Mozart (Hello Kyt)
  '
  'Tratamento para Qtde. de caracteres a serem impressas no
  'Código de barras
  lblPaulimaq.Enabled = False
  cboPosicao.Enabled = False
  '---------------------------------------------------------
  
  'Pega as tabela usada e joga na lista
  rsPreços.Index = "Só Tabela"
  Lugar = 0
  Últ_Tabela = ""

  Do
    rsPreços.Seek ">", Últ_Tabela
    If Not rsPreços.NoMatch Then
       Últ_Tabela = rsPreços("Tabela")
       Lista.AddItem Últ_Tabela, Lugar
       Lugar = Lugar + 1
    End If
  Loop Until (rsPreços.NoMatch)

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub
  
  If Not IsNull(rsParametros.Fields("BancoPDV").Value) Then
      sCaminhoArquivo = rsParametros.Fields("BancoPDV").Value
  Else
      sCaminhoArquivo = ""
  End If
  
  ' 19/02/2008 - Celso
  ' Nova implementação PROCON
  'Pega as tabelas usadas e joga nos combos da Tabelas de Preços
  rsPreços.Index = "Só Tabela"
  Lugar = 0
  Últ_Tabela = ""

  Do
    rsPreços.Seek ">", Últ_Tabela
    If Not rsPreços.NoMatch Then
       Últ_Tabela = rsPreços("Tabela")
       Tab_Vista.AddItem Últ_Tabela, Lugar
       Tab_Prazo.AddItem Últ_Tabela, Lugar
       Lugar = Lugar + 1
    End If
  Loop Until (rsPreços.NoMatch)
  
  If ParamCodigoUsuario <> 0 Then
      Combo.Text = ParamCodigoUsuario
      Combo_LostFocus
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsFuncionarios.Close
  rsProdutos.Close
  rsPreços.Close
  rsParametros.Close
  Set rsFuncionarios = Nothing
  Set rsProdutos = Nothing
  Set rsPreços = Nothing
  Set rsParametros = Nothing
End Sub

'04/11/2005 - mpdea
'Adicionado Sub Classe
Private Sub Imprime_Nome_Click()
  If Imprime_Nome.Value = 1 Then
    chkAppendClasse.Enabled = True
    chkAppendSubClasse.Enabled = True
    chk_fonteNomeLowerCase.Enabled = True
  Else
    chkAppendClasse.Enabled = False
    chkAppendSubClasse.Enabled = False
    chk_fonteNomeLowerCase.Value = vbUnchecked
    chk_fonteNomeLowerCase.Enabled = False
  End If
End Sub

Private Sub Imprime_Preço_Click()
  If Imprime_preço.Value = 1 Then
      Lista.Enabled = True
  End If
  
  If Imprime_preço.Value = 0 Then
      Lista.Enabled = False
      Lista.ListIndex = -1
  End If
End Sub

Private Sub txt_etiquetasEmBranco_LostFocus()
    If Not IsNumeric(txt_etiquetasEmBranco.Text) Then
        MsgBox "Ditite um número entre 1 a 99", vbInformation, "Atenção"
        txt_etiquetasEmBranco.SetFocus
    End If
End Sub
