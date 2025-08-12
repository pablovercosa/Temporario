VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManContasReceberII 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de Contas a Receber"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   375
   ClientWidth     =   11355
   HelpContextID   =   1300
   Icon            =   "ManContasReceberII.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6330
   ScaleWidth      =   11355
   Begin VB.Frame fraDados 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   11175
      Begin VB.CommandButton B_Calc_Juros 
         Caption         =   "Calcular &Juros"
         Height          =   375
         Left            =   3975
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   7695
         TabIndex        =   40
         Top             =   1200
         Width           =   3375
         Begin VB.CommandButton cmdReceber 
            BackColor       =   &H0000C000&
            Caption         =   "&Receber"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1200
            Width           =   1335
         End
         Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
            Bindings        =   "ManContasReceberII.frx":058A
            DataSource      =   "Data5"
            Height          =   315
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   750
            DataFieldList   =   "Descrição"
            ListAutoPosition=   0   'False
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
            Columns(0).Width=   6879
            Columns(0).Caption=   "Descrição"
            Columns(0).Name =   "Descrição"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Descrição"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1773
            Columns(1).Caption=   "Caixa"
            Columns(1).Name =   "Caixa"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Caixa"
            Columns(1).DataType=   2
            Columns(1).FieldLen=   256
            _ExtentX        =   1323
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.Label Label_Caixa 
            AutoSize        =   -1  'True
            Caption         =   "Caixa"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Nome_Caixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   41
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Impressão"
         Height          =   975
         Left            =   7695
         TabIndex        =   39
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton cmdEmiss 
            Caption         =   "&Imprimir"
            Height          =   400
            Left            =   1800
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optFatura 
            Caption         =   "Fatura"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optRecibo 
            Caption         =   "Recibo"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.TextBox Nota 
         Height          =   315
         Left            =   1095
         MaxLength       =   15
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Descrição 
         Height          =   315
         Left            =   3555
         MaxLength       =   30
         TabIndex        =   19
         Top             =   645
         Width           =   3630
      End
      Begin VB.CommandButton B_Dia 
         Caption         =   "Em &Dia"
         Height          =   375
         Left            =   2535
         TabIndex        =   20
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton B_Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5640
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin Threed.SSPanel sspVariasContas 
         Height          =   855
         Left            =   2535
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "ATENÇÃO: Baixa de várias contas. Digite a data de recebimento. O valor recebido será assumido como o valor da conta."
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin MSMask.MaskEdBox Vencimento 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Pagto 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   2235
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Acréscimo 
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   1455
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Valor_Pago 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   1845
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Data Pagto"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   52
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Index           =   9
         Left            =   2535
         TabIndex        =   51
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Nota"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   50
         Top             =   2670
         Width           =   345
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Valor Pago"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   1890
         Width           =   780
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Acréscimo"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   1170
         Width           =   690
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Nome_Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3555
         TabIndex        =   45
         Top             =   285
         Width           =   3630
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   8
         Left            =   2535
         TabIndex        =   44
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Baixa 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   330
         Width           =   840
      End
   End
   Begin VB.Data datFilial 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   1560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin SSDataWidgets_B.SSDBGrid grdCR 
      Height          =   1575
      Left            =   45
      TabIndex        =   11
      Top             =   1170
      Width           =   11265
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   14
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      RowHeight       =   423
      Columns.Count   =   14
      Columns(0).Width=   873
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Valor"
      Columns(1).Name =   "Valor"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1720
      Columns(2).Caption=   "Vcto"
      Columns(2).Name =   "Vcto"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   10
      Columns(2).Mask =   "##/##/####"
      Columns(2).PromptInclude=   -1  'True
      Columns(2).PromptChar=   32
      Columns(3).Width=   1455
      Columns(3).Caption=   "Desc"
      Columns(3).Name =   "Desc"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1376
      Columns(4).Caption=   "Acresc"
      Columns(4).Name =   "Acresc"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1905
      Columns(5).Caption=   "Val Receb"
      Columns(5).Name =   "Val Receb"
      Columns(5).Alignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1799
      Columns(6).Caption=   "Data Receb"
      Columns(6).Name =   "Data Receb"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   10
      Columns(6).Mask =   "##/##/####"
      Columns(6).PromptInclude=   -1  'True
      Columns(6).PromptChar=   32
      Columns(7).Width=   1640
      Columns(7).Caption=   "Nota"
      Columns(7).Name =   "Nota"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1508
      Columns(8).Caption=   "Cliente"
      Columns(8).Name =   "Cliente"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3016
      Columns(9).Caption=   "Nome"
      Columns(9).Name =   "Nome"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1191
      Columns(10).Caption=   "Seq"
      Columns(10).Name=   "Seq"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   9763
      Columns(11).Caption=   "Descrição"
      Columns(11).Name=   "Descricao"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "Contador"
      Columns(12).Name=   "Contador"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "Tipo Parcelamento"
      Columns(13).Name=   "Tipo Parcelamento"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      _ExtentX        =   19870
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Contas"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vencimento"
      Height          =   660
      Left            =   90
      TabIndex        =   30
      Top             =   45
      Width           =   3765
      Begin MSMask.MaskEdBox Vcto_Final 
         Height          =   315
         Left            =   2430
         TabIndex        =   1
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vcto_Inicial 
         Height          =   315
         Left            =   690
         TabIndex        =   0
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Final:"
         Height          =   255
         Left            =   2025
         TabIndex        =   32
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Inicial:"
         Height          =   255
         Left            =   135
         TabIndex        =   31
         Top             =   255
         Width           =   585
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   8550
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4110
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   8595
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   8595
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton B_Baixa 
      Caption         =   "&Alterar / Baixar"
      Height          =   400
      Left            =   9870
      TabIndex        =   10
      Top             =   660
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   7335
      TabIndex        =   29
      Top             =   60
      Width           =   2415
      Begin VB.OptionButton O_Vencimento 
         Caption         =   "&Vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton O_Cliente 
         Caption         =   "&Cliente"
         Height          =   195
         Left            =   1380
         TabIndex        =   6
         Top             =   255
         Width           =   870
      End
   End
   Begin VB.CommandButton B_Monta 
      Caption         =   "&Pesquisar"
      Height          =   400
      Left            =   9855
      TabIndex        =   9
      Top             =   135
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipos de Contas"
      Height          =   645
      Left            =   3930
      TabIndex        =   28
      Top             =   60
      Width           =   3330
      Begin VB.OptionButton O_Todas 
         Caption         =   "&Todas"
         Height          =   255
         Left            =   2445
         TabIndex        =   4
         Top             =   255
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Recebidas 
         Caption         =   "&Recebidas"
         Height          =   255
         Left            =   1230
         TabIndex        =   3
         Top             =   255
         Width           =   1095
      End
      Begin VB.OptionButton O_Receber 
         Caption         =   "A R&eceber"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   255
         Width           =   1215
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "ManContasReceberII.frx":059E
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   780
      Width           =   1050
      DataFieldList   =   "Nome"
      MaxDropDownItems=   16
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8625
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1376
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1852
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "ManContasReceberII.frx":05B2
      DataSource      =   "datFilial"
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   780
      Width           =   720
      DataFieldList   =   "Nome"
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   6350
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1005
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1270
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Sequência 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2160
      TabIndex        =   57
      Top             =   2880
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Baixa 
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      Height          =   195
      Index           =   7
      Left            =   1320
      TabIndex        =   56
      Top             =   2880
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label L_Tipo_Parc 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   345
      Left            =   840
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label L_Descrição 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   345
      Left            =   480
      TabIndex        =   54
      Top             =   2880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label L_Cliente 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   345
      Left            =   120
      TabIndex        =   53
      Top             =   2880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   840
      Width           =   300
   End
   Begin VB.Label lblFilial 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   35
      Top             =   780
      Width           =   2655
   End
   Begin VB.Label Nome_Fornecedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5985
      TabIndex        =   34
      Top             =   780
      Width           =   3795
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   840
      Width           =   585
   End
End
Attribute VB_Name = "frmManContasReceberII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstParametros As Recordset
Private Const GRID_FULL_HEIGHT As Single = 5100
Private Const GRID_MIN_HEIGHT As Single = 2175

Private Sub B_Baixa_Click()
  Dim Linha As Long
  Dim i As Integer
  Dim Pagas As Integer
  Dim book As Variant
  Dim Valor_Contas As Double
 
  Call StatusMsg("")
  
  If grdCR.SelBookmarks.Count < 1 Then
    DisplayMsg "Selecione ao menos uma linha da tabela antes."
    Exit Sub
  End If
  
  Pagas = False
  Valor_Contas = 0
  For i = 0 To (grdCR.SelBookmarks.Count - 1)
    book = grdCR.SelBookmarks(i)
    If grdCR.Columns("Val Receb").CellValue(book) <> 0 Then Pagas = True
    Valor_Contas = Valor_Contas + grdCR.Columns("Valor").CellValue(book)
  Next i
  If Pagas = True Then
    DisplayMsg "Uma ou mais contas selecionadas já foram recebidas e não podem ser baixadas. Caso deseje use a tela de lançamentos para alterá-las."
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde...")
  
  L_Cliente.Caption = ""
  L_Descrição.Caption = ""
  L_Tipo_Parc.Caption = ""
 
 
  If grdCR.SelBookmarks.Count <> 1 Then
    Valor.Text = Valor_Contas
    Valor.Enabled = False
    Vencimento.Enabled = False
    Desconto.Enabled = False
    Acréscimo.Enabled = False
    Valor_Pago.Text = Valor
    Valor_Pago.Enabled = False
    Nota.Enabled = False
    Nome_Cliente.Enabled = False
    sspVariasContas.Visible = True
    Descrição.Enabled = False
    Nome_Cliente.Caption = ""
    Descrição.Text = ""
    
    '07/08/2003 - mpdea
    'Desabilita botões
    B_Calc_Juros.Enabled = False
    B_Dia.Enabled = False
    
  Else
    Valor.Enabled = True
    Vencimento.Enabled = True
    Desconto.Enabled = True
    Acréscimo.Enabled = True
    Valor_Pago.Enabled = True
    Nota.Enabled = True
    Nome_Cliente.Enabled = True
    sspVariasContas.Visible = False
    Descrição.Enabled = True
    
    '07/08/2003 - mpdea
    'Habilita botões
    B_Calc_Juros.Enabled = True
    B_Dia.Enabled = True
    
  End If
 
  
  If grdCR.SelBookmarks.Count = 1 Then
    book = grdCR.SelBookmarks(0)
  '   Vencimento.Text = Format((grdCR.Columns("Vcto").CellValue(book)), "dd/mm/yyyy")
    Vencimento.Text = gsFormatDate(grdCR.Columns("Vcto").CellValue(book))
    Valor.Text = grdCR.Columns("Valor").CellValue(book)
    Desconto.Text = grdCR.Columns("Desc").CellValue(book)
    Acréscimo.Text = grdCR.Columns("Acresc").CellValue(book)
    Nota.Text = grdCR.Columns("Nota").CellValue(book)
    Sequência.Caption = grdCR.Columns("Seq").CellValue(book)
    Nome_Cliente.Caption = grdCR.Columns("Nome").CellValue(book)
    Descrição.Text = grdCR.Columns("Descricao").CellValue(book)
    L_Cliente.Caption = grdCR.Columns("Cliente").CellValue(book)
    L_Descrição.Caption = grdCR.Columns("Descricao").CellValue(book)
    L_Tipo_Parc.Caption = grdCR.Columns("Tipo Parcelamento").CellValue(book)
  End If
 
  B_Monta.Enabled = False
  B_Baixa.Enabled = False
  grdCR.Enabled = False
 
 
  '04/08/2003 - mpdea
  'Frame agrupando os campos
  fraDados.Visible = True
 
 
  Call StatusMsg("")
  
  grdCR.Height = GRID_MIN_HEIGHT

End Sub

Private Sub B_Calc_Juros_Click()
  Dim Valor_Aux As Double
  Dim Juros As Double
  Dim Erro As Integer
  Dim Dias As Integer
  
  
  Call StatusMsg("")
  
  If Not IsDate(Vencimento.Text) Then
    DisplayMsg "Data de vencimento incorreta, verfique."
    If Vencimento.Enabled = False Then Exit Sub
    Vencimento.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Pagto.Text) Then
    DisplayMsg "Digite a data de recebimento, para que os juros sejam calculados."
    Data_Pagto.SetFocus
    Exit Sub
  End If
  
  
  Valor.Text = gsHandleNull(Valor.Text)
  If Not IsNumeric(Valor.Text) Then
    DisplayMsg "Valor incorreto, verifique."
    Valor.SetFocus
    Exit Sub
  End If
  
  Valor_Aux = CDbl(Valor.Text)
  
  
  Dias = CDate(Data_Pagto.Text) - CDate(Vencimento.Text)
  If Dias = 0 Then
    DisplayMsg "Recebimento em dia, sem juros a calcular."
    Exit Sub
  End If
  
  Juros = Valor_Aux * rstParametros.Fields("Juros").Value / CDbl(30) * CDbl(Dias) / CDbl(100)
  
  Juros = Format(Juros, FORMAT_VALUE)
  
  Acréscimo.Text = Juros
  
  Baixa_DblClick (4)
  
  Call StatusMsg("")
  
End Sub

Private Sub B_Cancela_Click()
  
  '07/08/2003 - mpdea
  'Limpa data de vencimento e caixa
  Vencimento.Mask = ""
  Vencimento.Text = ""
  Vencimento.Mask = "##/##/####"
  
  Combo_Caixa.Text = ""
  Combo_Caixa_LostFocus
  
  
  Valor_Pago.Text = ""
  Data_Pagto.Mask = ""
  Data_Pagto.Text = ""
  Data_Pagto.Mask = "##/##/####"
  
  
  '04/08/2003 - mpdea
  'Frame agrupando os campos
  fraDados.Visible = False


  B_Monta.Enabled = True
  B_Baixa.Enabled = True
  
  grdCR.Enabled = True
  grdCR.Height = GRID_FULL_HEIGHT
  
End Sub

Private Sub B_Dia_Click()
  Valor_Pago.Text = Valor.Text
  Data_Pagto.Text = Vencimento.Text
  Desconto.Text = 0
  Acréscimo.Text = 0
End Sub

Private Sub B_Monta_Click()
 
  Call StatusMsg("")
  
  grdCR.Caption = "Contas"
  
  If Not IsDate(Vcto_Inicial.Text) Then
     DisplayMsg "Vencimento Inicial incorreto."
     Vcto_Inicial.SetFocus
     Exit Sub
  End If
  
  If Not IsDate(Vcto_Final.Text) Then
     DisplayMsg "Vencimento Final incorreto."
     Vcto_Final.SetFocus
     Exit Sub
  End If
    
  If CDate(Vcto_Final.Text) < CDate(Vcto_Inicial.Text) Then
    DisplayMsg "Vencimento inicial deve ser menor ou igual ao vencimento final."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(cboFilial.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
  End If
  
  B_Baixa.Enabled = True
  
  Call StatusMsg("Aguarde, pesquisando...")
  DoEvents
  
  Call LoadGridCR
  
  Call StatusMsg("")
 
End Sub

Private Sub LoadGridCR()
  Dim rsCR As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim Aux_Erro As Integer
  Dim sDescricao As String
  Dim sUnidVenda As String
  Dim sCod As String
  Dim sSql As String
  Dim Data_Ini As String
  Dim Data_Fim As String
  Dim sFilial As String
  
  On Error GoTo ErrHandler
  
  If Not IsDate(Vcto_Inicial.Text) Then
    DisplayMsg "Vencimento inicial incorreto."
    Vcto_Inicial.SetFocus
    Exit Sub
  End If
  Data_Ini = gsGetInvDate(Vcto_Inicial.Text)
  
  If Not IsDate(Vcto_Final.Text) Then
    DisplayMsg "Vencimento final incorreto."
    Vcto_Final.SetFocus
    Exit Sub
  End If
  Data_Fim = gsGetInvDate(Vcto_Final.Text)
  
  'Verifica a filial
  cboFilial_LostFocus
  If lblFilial.Caption = "" Then
    sFilial = "<> 0"
  Else
    sFilial = "= " & cboFilial
  End If
  
  bAllow = grdCR.AllowAddNew
  grdCR.AllowAddNew = True
  grdCR.AllowUpdate = True
  
  sSql = "SELECT Filial, Valor, Vencimento, [Contas a Receber].Desconto as Desconto, Acréscimo as Acrescimo, "
  sSql = sSql & "[Valor Recebido], [Data Recebimento] , Nota, Cliente, Cli_For.Nome, Sequência as Seq, Descrição as Descricao, Contador, [Tipo Parcelamento] FROM [Contas a Receber]"
  sSql = sSql + " INNER JOIN Cli_For ON ([Contas a Receber].Cliente = Cli_For.Código)"
  sSql = sSql + " WHERE Filial " & sFilial & " AND Vencimento >= " + Data_Ini
  sSql = sSql + " And Vencimento <= " + Data_Fim + " AND [Contas a Receber].Tipo = 'R'"
  
  If Nome_Fornecedor.Caption <> "" Then
    sSql = sSql + " And Cliente = " + Combo_Fornecedor.Text
  End If
  
  If O_Receber = True Then sSql = sSql + " AND [Valor Recebido] = 0"
  If O_Recebidas = True Then sSql = sSql + " AND [Valor Recebido] <> 0"
  
  If O_Cliente.Value = True Then sSql = sSql + " ORDER BY Cliente"
  If O_Vencimento.Value = True Then sSql = sSql + " ORDER BY Vencimento"
  
  Set rsCR = db.OpenRecordset(sSql, dbOpenDynaset)

  grdCR.RemoveAll
  grdCR.Redraw = False
  
  grdCR.Columns("Valor").NumberFormat = "##,###,##0.00"
  grdCR.Columns("Val Receb").NumberFormat = "##,###,##0.00"
  
  If Not rsCR.EOF Then
    With rsCR
      .MoveFirst
      Do While Not .EOF
        sRecord = .Fields("Filial") & vbTab & _
          .Fields("Valor") & vbTab & _
          .Fields("Vencimento") & vbTab & _
          .Fields("Desconto") & vbTab & _
          .Fields("Acrescimo") & vbTab & _
          .Fields("Valor Recebido") & vbTab & _
          .Fields("Data Recebimento") & vbTab & _
          .Fields("Nota") & vbTab & _
          .Fields("Cliente") & vbTab & _
          .Fields("Nome") & vbTab & _
          .Fields("Seq") & vbTab & _
          .Fields("Descricao") & vbTab & _
          .Fields("Contador") & vbTab & _
          .Fields("Tipo Parcelamento") '& vbTab & _
'          .Fields("Vendedor")
        grdCR.AddItem sRecord
        .MoveNext
      Loop
      .MoveFirst
    End With
    grdCR.Scroll -99, -99
    grdCR.Redraw = True
  Else
    DisplayMsg "Nenhuma conta encontrada segundo os critérios fornecidos."
    grdCR.Redraw = True
  End If

  grdCR.AllowAddNew = bAllow
  grdCR.AllowUpdate = bAllow

  rsCR.Close
  Set rsCR = Nothing
  Exit Sub
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ler registros do Contas a Receber."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

'12/02/2003 - mpdea
'Comentado código não utilizado
'
'Private Sub WriteGridCR()
'  Dim sSql As String
'  Dim bm As Variant
'  Dim nRow As Long
'  Dim rsCR As Recordset
'
'  On Error GoTo ErrHandler
'
'  grdCR.Update
'
'  Call ws.BeginTrans
'
'  Set rsCR = db.OpenRecordset(sSql, dbOpenDynaset)
'
'  With rsCR
'
'    If Not .EOF Then
'      Do While Not .EOF
'        .Delete
'        .MoveNext
'      Loop
'    End If
'
'    For nRow = 0 To grdCR.Rows - 1
'      bm = grdCR.AddItemBookmark(nRow)
'      If Len(grdCR.Columns("Contato").CellText(bm)) > 0 Then
'        .AddNew
'        '.Fields("Cliente") = cboCodigo.Text
'        .Fields("Seqüência") = nRow + 1
'        .Fields("Contato") = grdCR.Columns("Contato").CellText(bm)
'        .Fields("Cargo") = grdCR.Columns("Cargo").CellText(bm)
'        .Fields("Dia Aniversário") = CInt(gsHandleNull(grdCR.Columns("DiaAniv").CellValue(bm) & ""))
'        .Fields("Mês Aniversário") = grdCR.Columns("MesAniv").CellValue(bm) & ""
'        .Fields("Ramal") = grdCR.Columns("Ramal").CellValue(bm) & ""
'        .Fields("email") = grdCR.Columns("e-mail").CellValue(bm) & ""
'        .Update
'      End If
'    Next nRow
'
'  End With
'
'  rsCR.Close
'  Set rsCR = Nothing
'
'  Call ws.CommitTrans
'  Exit Sub
'
'ErrHandler:
'  gsTitle = LoadResString(201)
'  gsMsg = "Erro ao Atualizar Contatos."
'  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
'  gnStyle = vbOKOnly + vbExclamation
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'  Exit Sub
'
'End Sub

Private Sub cboFilial_CloseUp()
  lblFilial.Caption = cboFilial.Columns("Nome").Text
  cboFilial.Text = cboFilial.Columns("Filial").Text
End Sub

Private Sub cboFilial_KeyPress(KeyAscii As Integer)
  If Not cboFilial.DroppedDown Then
    KeyAscii = gnLimitKeyPress(cboFilial, 2, KeyAscii, True)
  End If
End Sub

Private Sub cboFilial_LostFocus()
  lblFilial.Caption = gsGetNameFilial(Val(cboFilial.Text))
End Sub

Private Sub cmdEmiss_Click()
  Dim nX As Integer
  Dim vBook As Variant
  Dim cValue As Currency
  Dim nCodCliente As Long
  
  If grdCR.SelBookmarks.Count = 0 Then
    Exit Sub
  Else
    If optFatura.Value And grdCR.SelBookmarks.Count > 1 Then
      DisplayMsg "Para a emissão de Fatura selecione apenas uma conta."
      Exit Sub
    ElseIf grdCR.SelBookmarks.Count > 1 Then 'Recibo para várias contas
      'Verifica se as contas selecionadas pertencem ao mesmo cliente
      'e obtém o valor total
      For nX = 0 To (grdCR.SelBookmarks.Count - 1)
        vBook = grdCR.SelBookmarks(nX)
        If nX = 0 Then
          nCodCliente = grdCR.Columns("Cliente").CellValue(vBook)
        Else
          If nCodCliente <> grdCR.Columns("Cliente").CellValue(vBook) Then
            DisplayMsg "Para a emissão de Recibo de várias contas selecione apenas o mesmo cliente."
            Exit Sub
          End If
        End If
        cValue = cValue + grdCR.Columns("Valor").CellValue(vBook) + _
          grdCR.Columns("Acresc").CellValue(vBook) - _
          grdCR.Columns("Desc").CellValue(vBook)
      Next nX
    End If
    
    nReciboVALOR = CDbl(Valor.Text)
    nReciboACRESCIMO = CDbl(Acréscimo.Text)
    nReciboDESCONTO = CDbl(Desconto.Text)
    
    With frmEmiteFatura
      If grdCR.SelBookmarks.Count = 1 Then
        vBook = grdCR.SelBookmarks(0)
        .Transf1.Caption = grdCR.Columns("Vcto").CellValue(vBook)
        .Transf2.Caption = grdCR.Columns("Contador").CellValue(vBook)
        .lblCheckValue.Caption = "True"
      Else
        vBook = grdCR.SelBookmarks(0)
        .Transf1.Caption = grdCR.Columns("Vcto").CellValue(vBook)
        .Transf2.Caption = grdCR.Columns("Contador").CellValue(vBook)
        .lblCheckValue.Caption = "False"
        .L_Valor.Caption = Format(cValue, FORMAT_VALUE)
      End If
      If optFatura.Value Then
        .Caption = "Emissão de Fatura"
        .Tipo.Caption = "F"
      Else
        .Caption = "Emissão de Recibo"
        .Tipo.Caption = "R"
      End If
      .L_Encontrar.Caption = "SIM"
      .optTotalParcela.Enabled = False
      .Show vbModal
    End With
  End If
  
End Sub

Private Sub Baixa_DblClick(Index As Integer)
 If Index = 4 Then
   If IsNull(Valor.Text) Then Exit Sub
   If Valor.Text = "" Then Exit Sub
   If Not IsNumeric(Valor.Text) Then Exit Sub
   If CDbl(Valor.Text) <= 0 Then Exit Sub
 
   If IsNull(Desconto.Text) Then Desconto.Text = 0
   If Desconto.Text = "" Then Desconto.Text = 0
   If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0
   If CDbl(Desconto.Text) < 0 Then Desconto.Text = 0

   If IsNull(Acréscimo.Text) Then Acréscimo.Text = 0
   If Acréscimo.Text = "" Then Acréscimo.Text = 0
   If Not IsNumeric(Acréscimo.Text) Then Acréscimo.Text = 0
   If CDbl(Acréscimo.Text) < 0 Then Acréscimo.Text = 0

   Valor_Pago.Text = CDbl(Valor.Text) - CDbl(Desconto.Text) + CDbl(Acréscimo.Text)
 End If
End Sub

'07/08/2003 - mpdea
'Implementado evento
Private Sub Combo_Caixa_CloseUp()
  Combo_Caixa.Text = Combo_Caixa.Columns("Caixa").Text
  Combo_Caixa_LostFocus
End Sub

'04/08/2003 - mpdea
'Modificado busca do registro

Private Sub Combo_Caixa_LostFocus()
  Dim bytCaixa As Byte
  
  On Error GoTo TratarErro
  
  Call IsDataType(dtByte, Combo_Caixa.Text, bytCaixa)
  
  If bytCaixa <= 0 Then
    Nome_Caixa.Caption = ""
  Else
    '08/06/2005 - Daniel
    'Correção do Runtime error '91'
    'para às linhas comentadas abaixo
    '
    'With Data5.Recordset
    '  .FindFirst "Caixa = " & dtByte    'O erro ocorria exatamente nesta linha colocamos o bytCaixa no lugar do dtByte mas não resolveu...
    '  If .NoMatch Then
    '    Nome_Caixa.Caption = ""
    '  Else
    '    Nome_Caixa.Caption = .Fields("Descrição").Value & ""
    '  End If
    'End With
    Dim rstCaixasEmUso As Recordset
    Dim strSQL         As String

    strSQL = "SELECT Descrição FROM [Caixas em Uso] WHERE Caixa = " & bytCaixa

    Set rstCaixasEmUso = db.OpenRecordset(strSQL, dbOpenDynaset)

    With rstCaixasEmUso
      If Not (.BOF And .EOF) Then
        .MoveFirst
        Nome_Caixa.Caption = .Fields("Descrição").Value & ""
      Else
        Nome_Caixa.Caption = ""
      End If
      .Close
    End With

    Set rstCaixasEmUso = Nothing
    
  End If

  Exit Sub

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns("Código").Text
  Combo_Fornecedor_LostFocus
End Sub

'04/08/2003 - mpdea
'Modificado busca do registro
Private Sub Combo_Fornecedor_LostFocus()
  Dim lngCodCliente As Long
  
  
  Call IsDataType(dtLong, Combo_Fornecedor.Text, lngCodCliente)
  
  If lngCodCliente <= 0 Then
    Nome_Fornecedor.Caption = ""
  Else
    With Data1.Recordset
      .FindFirst "Código = " & lngCodCliente
      If .NoMatch Then
        Nome_Fornecedor.Caption = ""
      Else
        Nome_Fornecedor.Caption = .Fields("Nome").Value & ""
      End If
    End With
  End If
 
End Sub

Private Sub Data_Pagto_LostFocus()
  Data_Pagto.Text = Ajusta_Data(Data_Pagto.Text)
End Sub

Private Sub Data_Pagto_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Pagto.Text = frmCalendario.gsDateCalender(Data_Pagto.Text)
  End Select
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rstParametros = db.OpenRecordset("SELECT * FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName
  Data5.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  grdCR.Height = GRID_FULL_HEIGHT  'Posição inicial
  
  If Not gbCaixas Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
  End If
   
  Call GetSettings
  
  cboFilial.Text = gnCodFilial
  cboFilial_LostFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Call SaveSetting("QuickStore", "CRMan", "Data1", Vcto_Inicial.Text)
  Call SaveSetting("QuickStore", "CRMan", "Data2", Vcto_Final.Text)
  
  rstParametros.Close
  
  Set rstParametros = Nothing

End Sub

Private Sub GetSettings()
  Vcto_Final.Text = GetSetting("QuickStore", "CRMan", "Data2", CDate(Date))
  Vcto_Inicial.Text = GetSetting("QuickStore", "CRMan", "Data1", CDate(Date))
End Sub

Private Sub grdCR_AfterDelete(RtnDispErrMsg As Integer)
  grdCR.Scroll 0, -32767
  grdCR.Scroll 0, 32767
End Sub

Private Sub grdCR_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub

Private Sub grdCR_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  Dim Val_Selec As Double
  Dim i As Integer
  Dim book As Variant
  Dim sConta As String
  Dim sSelec As String
  
  If grdCR.SelBookmarks.Count = 0 Then
    grdCR.Caption = "Contas: nenhuma selecionada"
    Exit Sub
  End If
  
  
  Val_Selec = 0#
  For i = 0 To (grdCR.SelBookmarks.Count - 1)
    book = grdCR.SelBookmarks(i)
    Val_Selec = Val_Selec + grdCR.Columns("Valor").CellValue(book)
  Next i
  
  If grdCR.SelBookmarks.Count = 1 Then
    sConta = "Conta: "
    sSelec = " selecionada"
  Else
    sConta = "Contas: "
    sSelec = " selecionadas"
  End If
  
  grdCR.Caption = sConta & CStr(grdCR.SelBookmarks.Count) & sSelec & ", valor " + Format((CStr(Val_Selec)), "Currency")

End Sub

Private Sub Nota_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Vcto_Final_LostFocus()
  Vcto_Final.Text = Ajusta_Data(Vcto_Final.Text)
End Sub

Private Sub Vcto_Final_GotFocus()
  With Vcto_Final
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub Vcto_Final_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vcto_Final.Text = frmCalendario.gsDateCalender(Vcto_Final.Text)
  End Select
End Sub

Private Sub Vcto_Inicial_LostFocus()
  Vcto_Inicial.Text = Ajusta_Data(Vcto_Inicial.Text)
End Sub

Private Sub Vcto_Inicial_GotFocus()
  With Vcto_Inicial
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub Vcto_Inicial_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vcto_Inicial.Text = frmCalendario.gsDateCalender(Vcto_Inicial.Text)
  End Select
End Sub

Private Sub Vencimento_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Vencimento.Text = frmCalendario.gsDateCalender(Vencimento.Text)
  End Select
End Sub

'16/07/2003 - mpdea
'Recebimento da conta através da tela de Recebimento
Private Sub cmdReceber_Click()
  Dim rstContasReceber As Recordset
  Dim rstCaixa As Recordset
  Dim rstCliFor As Recordset
  Dim rstCartoes As Recordset
  Dim strSQL As String
  Dim blnInTransaction As Boolean
  
  Dim typTotalizadores As tpPaymentType
  Dim typRecebimento As tpPaymentType
  Dim intOrdem As Integer
  Dim varBookmark As Variant
  Dim dblSaldoAnterior As Double
  Dim dblValorPago As Double
  Dim bytCaixa As Byte
  Dim dteDataPgto As Date
  Dim intX As Integer
  Dim lngContador As Long
  Dim lngCodCliente As Long
  Dim intRet As Integer
  
  Dim intBanco As Integer
  Dim strCheque As String
  Dim strData As String
  Dim dblValor As Double
  Dim intCount As Integer
  Dim intParcelas As Integer
  
  Dim strTipoParcelamento As String
  Dim bytNrContaCC As Byte
  
  Dim intCartaoAdministradora As Integer
  Dim bytCartaoQtdeParcelas As Byte
  Dim dblCartaoVlrParcela As Double
  
  Dim dblValorPagar As Double
  
  Dim intRepeatUpdateLocked As Integer
  
  On Error GoTo ErrHandler
  
  
  Call StatusMsg("")
  
  '----------------------------------------------------------------------------
  'Validação inicial
  '----------------------------------------------------------------------------
  With grdCR
    For intX = 0 To (.SelBookmarks.Count - 1)
      varBookmark = .SelBookmarks(intX)
      .Bookmark = varBookmark
      If intX = 0 Then
        lngCodCliente = .Columns("Cliente").CellValue(varBookmark)
      Else
        If lngCodCliente <> .Columns("Cliente").CellValue(varBookmark) Then
          DisplayMsg "Para recebimento de várias contas selecione apenas o mesmo cliente."
          Exit Sub
        End If
      End If
    Next intX
  End With
  
  'Valor a ser pago
  Call IsDataType(dtDouble, Valor_Pago.Text, dblValorPago)
  If dblValorPago <= 0 Then
    DisplayMsg "Valor a Pagar incorreto."
    SelectAllText Valor_Pago, True
    Exit Sub
  End If
  
  '19/08/2003 - mpdea
  'Valor a pagar
  Call IsDataType(dtDouble, Valor.Text, dblValorPagar)
  If dblValorPago < dblValorPagar Then
    DisplayMsg "Valor a Pagar inferior ao ser Pago."
    SelectAllText Valor_Pago, True
    Exit Sub
  End If
  
  'Data do pagamento
  Call IsDataType(dtDate, Data_Pagto.Text, dteDataPgto)
  If dteDataPgto = 0 Then
    DisplayMsg "Data de pagamento incorreta."
    SelectAllText Data_Pagto, True
    Exit Sub
  End If
  
  'Caixa
  If Nome_Caixa.Caption = "" Then
    DisplayMsg "Caixa incorreto."
    SelectAllText Combo_Caixa, True
    Exit Sub
  End If
  Call IsDataType(dtByte, Combo_Caixa.Text, bytCaixa)
  '-----------------------------------------------------------------------------
  
  
  'Busca cliente
  Set rstCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código = " & lngCodCliente, dbOpenSnapshot)
  
  
  '-----------------------------------------------------------------------------
  'Recebimento
  '-----------------------------------------------------------------------------
  With frmRecebimento
    .Limpa_Tela (0)
    .Só_Leitura.Value = 0
    .L_Sequência.Caption = "-1"
    .Receber.Caption = Format(dblValorPago, FORMAT_VALUE)
    .Intervalo_Parc.Caption = rstParametros.Fields("VR Intervalo Parc").Value
    .Combo_Banco.Text = rstCliFor("Conta Cobrança")
    .Conta.Enabled = rstCliFor.Fields("Tem Conta").Value
    .Max_Cheques.Caption = 0
    .Max_Parcelas.Caption = 0
    If Not rstCliFor.Fields("Faturado").Value Then
      .Max_Cheques.Caption = "1"
      .Max_Parcelas.Caption = "1"
    Else
      .Max_Cheques.Caption = "9999"
      .Max_Parcelas.Caption = "9999"
    End If
    
    'Fecha recordset de clientes
    rstCliFor.Close
    Set rstCliFor = Nothing

    .Show vbModal
    If .Retorno.Caption <> "OK" Then
      Unload frmRecebimento
      Exit Sub
    End If
    
    typRecebimento.dblDinheiro = Format(CDbl(frmRecebimento.Dinheiro.Text), FORMAT_VALUE)
    typRecebimento.dblCartao = Format(CDbl(frmRecebimento.Cartão.Text), FORMAT_VALUE)
    typRecebimento.dblVale = Format(CDbl(frmRecebimento.Vale.Text), FORMAT_VALUE)
    typRecebimento.dblCheque = Format(frmRecebimento.Pega_Total_Cheques_Separado(False), FORMAT_VALUE)
    typRecebimento.dblChequePre = Format(frmRecebimento.Pega_Total_Cheques_Separado(True), FORMAT_VALUE)
    typRecebimento.dblParcelamento = Format(frmRecebimento.Pega_Total_Parcelas, FORMAT_VALUE)
    
  
    If .O_Banco.Value Then
      strTipoParcelamento = "B"
      Call IsDataType(dtByte, .Combo_Banco.Text, bytNrContaCC)
    ElseIf .O_Carteira.Value Then
      strTipoParcelamento = "C"
    ElseIf .O_Carnet.Value Then
      strTipoParcelamento = "T"
    End If
    
  End With
  '-----------------------------------------------------------------------------
    
  
  Call WaitSeconds(1, True) 'Aguarda um segundo para o refresh
  Me.Refresh
  
  Call StatusMsg("Aguarde...")
  Screen.MousePointer = vbHourglass
  
  'Inicia transação
  ws.BeginTrans
  blnInTransaction = True
  
  
  'Abre recordset de Contas a Receber
  strSQL = "SELECT * FROM [Contas a Receber] ORDER BY Contador;"
  Set rstContasReceber = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  
  '-----------------------------------------------------------------------------
  'Caixa
  '-----------------------------------------------------------------------------
  'Verifica o início do caixa, abertura do dia e retorna os últimos valores
  If Not gbCheckOpenCaixa(bytCaixa, gnUserCode, dblSaldoAnterior, intOrdem, typTotalizadores) Then
    'Ocorreu erro e a mensagem é exibida pela função
    ws.Rollback
    blnInTransaction = False
    Exit Sub
  End If
  '
  'Atualiza caixa
  Set rstCaixa = db.OpenRecordset("Caixa", dbOpenDynaset)
  With rstCaixa
    intOrdem = intOrdem + 1
    .AddNew
    .Fields("Filial").Value = gnCodFilial
    .Fields("Data").Value = Data_Atual
    .Fields("Caixa").Value = bytCaixa
    .Fields("Funcionário").Value = gnUserCode
    .Fields("Hora").Value = Format(Time, "hh:mm:ss")
    .Fields("Ordem").Value = intOrdem
    .Fields("Descrição").Value = Left("Conta recebida - " & Nome_Cliente.Caption, 30)
    
    .Fields("Dinheiro").Value = typRecebimento.dblDinheiro
    .Fields("Total Dinheiro").Value = typTotalizadores.dblDinheiro + .Fields("Dinheiro").Value
    .Fields("Cheques").Value = typRecebimento.dblCheque
    .Fields("Total Cheques").Value = typTotalizadores.dblCheque + .Fields("Cheques").Value
    .Fields("Cheques Pré").Value = typRecebimento.dblChequePre
    .Fields("Total Cheques Pré").Value = typTotalizadores.dblChequePre + .Fields("Cheques Pré").Value
    .Fields("Cartões").Value = typRecebimento.dblCartao
    .Fields("Total Cartões").Value = typTotalizadores.dblCartao + .Fields("Cartões").Value
    .Fields("Vales").Value = typRecebimento.dblVale
    .Fields("Total Vales").Value = typTotalizadores.dblVale + .Fields("Vales").Value
    .Fields("Parcelamento").Value = typRecebimento.dblParcelamento
    .Fields("Total Parcelamento").Value = typTotalizadores.dblParcelamento + .Fields("Parcelamento").Value
    .Fields("Saldo Anterior").Value = dblSaldoAnterior
    .Fields("Final").Value = Format(.Fields("Saldo Anterior").Value + _
      .Fields("Dinheiro").Value + .Fields("Cheques").Value + _
      .Fields("Cheques Pré").Value + .Fields("Cartões").Value + _
      .Fields("Vales").Value, FORMAT_VALUE)
    .Update
  End With
  '-----------------------------------------------------------------------------
  
  
  '01, 04 e 07/08/2003 - mpdea
  '-----------------------------------------------------------------------------
  'Atualiza contas a receber
  '-----------------------------------------------------------------------------
  With rstContasReceber
    
    '-----------------------------------------------------------------------------
    'Baixa conta(s)
    '-----------------------------------------------------------------------------
    If grdCR.SelBookmarks.Count = 1 Then
      varBookmark = grdCR.SelBookmarks(0)
      grdCR.Bookmark = varBookmark
      lngContador = grdCR.Columns("Contador").CellValue(varBookmark)
      
      .FindFirst "Contador = " & lngContador
      If .NoMatch Then
        Call StatusMsg("")
        Screen.MousePointer = vbDefault
        MsgBox "Erro ao localizar a conta para baixa.", vbCritical, "Erro"
        ws.Rollback
        blnInTransaction = False
        Exit Sub
      Else
        .Edit
        .Fields("Vencimento") = CDate(Vencimento.Text)
'        .Fields("Valor") = CDbl(gsHandleNull(Valor.Text))
        .Fields("Desconto").Value = CDbl(gsHandleNull(Desconto.Text))
        .Fields("Acréscimo").Value = CDbl(gsHandleNull(Acréscimo.Text))
        .Fields("Valor Recebido").Value = dblValorPago
'        If IsDate(Data_Pagto.Text) Then
          .Fields("Data Recebimento").Value = dteDataPgto
'        End If
        If Nota.Text = "" Then
          Nota.Text = "0"
        End If
        .Fields("Nota").Value = Nota.Text & ""
        .Fields("Descrição") = Descrição.Text
        .Fields("Data Alteração") = Data_Atual
        .Update
      End If
    Else
      For intX = 0 To (grdCR.SelBookmarks.Count - 1)
        varBookmark = grdCR.SelBookmarks(intX)
        grdCR.Bookmark = varBookmark
        lngContador = grdCR.Columns("Contador").CellValue(varBookmark)
        
        .FindFirst "Contador = " & lngContador
        If .NoMatch Then
          Call StatusMsg("")
          Screen.MousePointer = vbDefault
          MsgBox "Erro ao localizar a conta para baixa.", vbCritical, "Erro"
          ws.Rollback
          blnInTransaction = False
          Exit Sub
        Else
          .Edit
          .Fields("Valor Recebido").Value = .Fields("Valor").Value
          .Fields("Data Recebimento").Value = dteDataPgto
          .Fields("Data Alteração").Value = Data_Atual
          .Update
        End If
      Next intX
    End If
    '-----------------------------------------------------------------------------
        
        
    '---------------------------------------------------------------------------
    'Cheque
    '---------------------------------------------------------------------------
    For intX = 1 To 50
      intRet = frmRecebimento.Pega_Banco(intX, intBanco, strCheque, strData, dblValor)
      If intRet = 1 Then
        .AddNew
        .Fields("Tipo").Value = "C"
        .Fields("Filial").Value = gnCodFilial
        .Fields("Sequência").Value = 0
        .Fields("Cliente").Value = lngCodCliente
        .Fields("Banco").Value = intBanco
        .Fields("Cheque").Value = strCheque
        .Fields("Vencimento").Value = Format(strData, "dd/mm/yyyy")
        .Fields("Valor").Value = Format(dblValor, FORMAT_VALUE)
        .Fields("Vendedor").Value = 0
        .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
        If CDate(strData) = CDate(Data_Atual) Then
          .Fields("Processado").Value = True
          .Fields("Valor Recebido").Value = .Fields("Valor").Value
          .Fields("Data Recebimento").Value = .Fields("Vencimento").Value
        End If
        .Update
      End If
    Next intX
    '---------------------------------------------------------------------------
      
      
    '---------------------------------------------------------------------------
    'Parcelamento
    '---------------------------------------------------------------------------
    intCount = 0
    For intX = 1 To 50
      intRet = frmRecebimento.Pega_Parcela(intX, strData, dblValor, intParcelas)
      If intRet = 1 Then
        intCount = intCount + 1
        .AddNew
        .Fields("Tipo").Value = "R"
        .Fields("Filial").Value = gnCodFilial
        .Fields("Cliente").Value = lngCodCliente
        .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Parcela").Value = intCount
        .Fields("Descrição").Value = "Parcela " & intCount & "/" & intParcelas
        .Fields("Vencimento").Value = Format(strData, "dd/mm/yyyy")
        .Fields("Valor").Value = Format(dblValor, FORMAT_VALUE)
        .Fields("Sequência").Value = 0
        .Fields("Nota").Value = 0
        .Fields("Vendedor").Value = 0
        .Fields("Tipo Parcelamento").Value = strTipoParcelamento
        .Fields("Conta Boleto").Value = bytNrContaCC
        .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Update
      End If
    Next intX
    '---------------------------------------------------------------------------
  
  
    '---------------------------------------------------------------------------
    'Cartão
    '---------------------------------------------------------------------------
    If typRecebimento.dblCartao > 0 Then
      'Administradora
      Call IsDataType(dtInteger, frmRecebimento.Combo_Empresa.Text, intCartaoAdministradora)
      'Quantidade de parcelas
      Call IsDataType(dtByte, frmRecebimento.Label_Cartão2.Caption, bytCartaoQtdeParcelas)
      'Valor da parcela
      Call IsDataType(dtDouble, frmRecebimento.Label_Cartão4.Caption, dblCartaoVlrParcela)
      
      strSQL = "SELECT * FROM Cartões WHERE Código = " & _
        intCartaoAdministradora
      Set rstCartoes = db.OpenRecordset(strSQL, dbOpenSnapshot)
      If Not (rstCartoes.BOF And rstCartoes.EOF) Then
        For intX = 1 To bytCartaoQtdeParcelas
          .AddNew
          .Fields("Tipo").Value = "O"
          .Fields("Filial").Value = gnCodFilial
          .Fields("Sequência").Value = 0
          .Fields("Cliente").Value = lngCodCliente
          .Fields("Administradora").Value = intCartaoAdministradora
          .Fields("Cartão").Value = frmRecebimento.Num_Cartão.Text
          .Fields("Vencimento").Value = (CDate(Data_Atual) + rstCartoes.Fields("Dias Pagar").Value + ((intX - 1) * 30))
          .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
          
          If bytCartaoQtdeParcelas = 1 Then
            .Fields("Valor Cartão").Value = typRecebimento.dblCartao
            .Fields("Valor").Value = Round(CDbl(typRecebimento.dblCartao * ((1 - rstCartoes.Fields("Taxa").Value / 100))), 2)
          Else
            .Fields("Valor Cartão").Value = dblCartaoVlrParcela
            .Fields("Valor").Value = Round(CDbl(dblCartaoVlrParcela * ((1 - rstCartoes.Fields("Taxa").Value / 100))), 2)
          End If
          
          .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
          .Update
        Next intX
      End If
      rstCartoes.Close
      Set rstCartoes = Nothing
    End If
    '---------------------------------------------------------------------------
    
    
    '---------------------------------------------------------------------------
    'Vendas à vista
    '---------------------------------------------------------------------------
    If typRecebimento.dblDinheiro + typRecebimento.dblVale Then
      If rstParametros.Fields("Gerar Conta Paga").Value Then
        .AddNew
        .Fields("Tipo").Value = "R"
        .Fields("Filial").Value = gnCodFilial
        .Fields("Cliente").Value = lngCodCliente
        .Fields("Sequência").Value = 0
        .Fields("Nota").Value = 0
        .Fields("Vendedor").Value = 0
        .Fields("Descrição").Value = "Pagamento à vista"
        .Fields("Valor").Value = Format(typRecebimento.dblDinheiro + typRecebimento.dblVale, FORMAT_VALUE)
        .Fields("Valor Recebido").Value = .Fields("Valor").Value
        .Fields("Data Recebimento").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Vencimento").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Update
      End If
    End If
    
    .Close
  End With
  Set rstContasReceber = Nothing
  
  
  'Finaliza transação
  ws.CommitTrans
  blnInTransaction = False
  
  'Descarrega a tela de recebimento
  Unload frmRecebimento
  
  'Atualiza a tela
  B_Cancela_Click
  B_Monta_Click
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
ErrHandler:
  'Descarrega a tela de recebimento
  Unload frmRecebimento

  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      End If
    Case Else
      'Cancelamento da transação
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro em Manutenção - Contas a receber: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End Select
End Sub
