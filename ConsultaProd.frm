VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmConsultaProd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Consulta Produtos"
   ClientHeight    =   7275
   ClientLeft      =   285
   ClientTop       =   315
   ClientWidth     =   13155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1100
   Icon            =   "ConsultaProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   877
   Begin VB.TextBox Con_Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   780
      MaxLength       =   20
      TabIndex        =   39
      Top             =   675
      Width           =   3390
   End
   Begin VB.CommandButton cmdPesquisar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "P&esquisar c/ outra tela"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2745
      Width           =   1695
   End
   Begin VB.TextBox Con_Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   3645
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   3525
      Width           =   7755
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Incluir"
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
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3465
      Width           =   1695
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq1"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq2"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Pesq3"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
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
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   135
      Width           =   1695
   End
   Begin VB.Data Data_Preço 
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
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Pesquisa"
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   90
      TabIndex        =   21
      Top             =   15
      Width           =   11280
      Begin VB.OptionButton O_Pesquisa3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Pesquisa 3"
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
         Left            =   4965
         TabIndex        =   15
         Top             =   255
         Width           =   1260
      End
      Begin VB.OptionButton O_Pesquisa2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Pesquisa 2"
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
         Left            =   3045
         TabIndex        =   14
         Top             =   255
         Width           =   1200
      End
      Begin VB.OptionButton O_Pesquisa1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Pesquisa 1"
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
         Left            =   1275
         TabIndex        =   13
         Top             =   255
         Width           =   1230
      End
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
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
         Left            =   7005
         TabIndex        =   12
         Top             =   255
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
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
         Left            =   8730
         TabIndex        =   11
         Top             =   255
         Width           =   1005
      End
   End
   Begin VB.CommandButton B_Anterior 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   750
      Width           =   735
   End
   Begin VB.CommandButton B_Limpa 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Limpar"
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
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2085
      Width           =   1695
   End
   Begin VB.CommandButton B_Próximo 
      BackColor       =   &H00C0FFFF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   750
      Width           =   735
   End
   Begin VB.TextBox Con_Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4980
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   6390
   End
   Begin SSDataWidgets_B.SSDBGrid Grade_Preço 
      Bindings        =   "ConsultaProd.frx":4E95A
      Height          =   2145
      Left            =   6705
      TabIndex        =   10
      Top             =   5085
      Width           =   4710
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowColumnShrinking=   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   16777152
      RowHeight       =   423
      ExtraHeight     =   53
      Columns(0).Width=   3200
      UseDefaults     =   0   'False
      _ExtentX        =   8308
      _ExtentY        =   3784
      _StockProps     =   79
      Caption         =   "Preço"
      BackColor       =   -2147483633
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
   Begin SSDataWidgets_B.SSDBGrid Grade_Estoque 
      Height          =   2685
      Left            =   75
      TabIndex        =   9
      Top             =   4545
      Width           =   6570
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   ":"
      Col.Count       =   4
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   16777152
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   4
      Columns(0).Width=   2223
      Columns(0).Caption=   "Tamanho"
      Columns(0).Name =   "Col1"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2566
      Columns(1).Caption=   "Cor"
      Columns(1).Name =   "Cor"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1244
      Columns(2).Caption=   "Qtde"
      Columns(2).Name =   "Qtde"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4551
      Columns(3).Caption=   "Filial"
      Columns(3).Name =   "Filial"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      UseDefaults     =   0   'False
      _ExtentX        =   11589
      _ExtentY        =   4736
      _StockProps     =   79
      Caption         =   "Estoque"
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
   Begin SSDataWidgets_B.SSDBCombo Con_Pesquisa3 
      Bindings        =   "ConsultaProd.frx":4E973
      DataSource      =   "Data6"
      Height          =   405
      Left            =   4635
      TabIndex        =   5
      Top             =   2145
      Width           =   1575
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
      Columns(0).Width=   8334
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2064
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   714
      _StockProps     =   93
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
   Begin SSDataWidgets_B.SSDBCombo Con_Pesquisa2 
      Bindings        =   "ConsultaProd.frx":4E987
      DataSource      =   "Data5"
      Height          =   420
      Left            =   4635
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
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
      Columns(1).Width=   1429
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   741
      _StockProps     =   93
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
   Begin SSDataWidgets_B.SSDBCombo Con_Pesquisa1 
      Bindings        =   "ConsultaProd.frx":4E99B
      DataSource      =   "Data4"
      Height          =   420
      Left            =   4635
      TabIndex        =   1
      Top             =   1215
      Width           =   1575
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
      Columns(0).Width=   7699
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1535
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2778
      _ExtentY        =   741
      _StockProps     =   93
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
   Begin VB.Label lbl_codigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   40
      Top             =   705
      Width           =   585
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3660
      TabIndex        =   37
      Top             =   2580
      Width           =   555
   End
   Begin VB.Label Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      Height          =   375
      Left            =   3645
      TabIndex        =   36
      Top             =   2850
      Width           =   2295
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Classe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6000
      TabIndex        =   35
      Top             =   2580
      Width           =   945
   End
   Begin VB.Label Sub_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      Height          =   375
      Left            =   6000
      TabIndex        =   34
      Top             =   2850
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3690
      TabIndex        =   33
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localização"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8475
      TabIndex        =   32
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Localização 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      Height          =   375
      Left            =   8475
      TabIndex        =   31
      Top             =   2850
      Width           =   2895
   End
   Begin VB.Image imgFoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2685
      Left            =   90
      Stretch         =   -1  'True
      Top             =   1095
      Width           =   3495
   End
   Begin VB.Label Inativo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "PRODUTO INATIVO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   28
      Top             =   3810
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Ordenação 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   195
      TabIndex        =   27
      Top             =   1155
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label_Pesq1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3645
      TabIndex        =   26
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label_Pesq2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3645
      TabIndex        =   25
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Nome_Pesq1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   420
      Left            =   6240
      TabIndex        =   2
      Top             =   1200
      Width           =   5130
   End
   Begin VB.Label Nome_Pesq2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   420
      Left            =   6240
      TabIndex        =   4
      Top             =   1680
      Width           =   5130
   End
   Begin VB.Label Label_Pesq3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3645
      TabIndex        =   24
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Nome_Pesq3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   420
      Left            =   6240
      TabIndex        =   6
      Top             =   2145
      Width           =   5130
   End
   Begin VB.Label Con_Fabricante 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      Height          =   375
      Left            =   6705
      TabIndex        =   7
      Top             =   4665
      Width           =   4695
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricante"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6705
      TabIndex        =   23
      Top             =   4410
      Width           =   900
   End
   Begin VB.Label Con_IPI 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
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
      Height          =   375
      Left            =   12765
      TabIndex        =   8
      Top             =   4965
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IPI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12375
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbl_nome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4380
      TabIndex        =   20
      Top             =   735
      Width           =   495
   End
End
Attribute VB_Name = "frmConsultaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TB_Produtos    As Recordset
Dim TB_Estoque     As Recordset
Dim TB_Preços      As Recordset
Dim TB_Parâmetros  As Recordset
Dim Rec_Preços     As Recordset
Dim TB_Classes     As Recordset
Dim TB_Sub_Classes As Recordset
Dim TB_Grade       As Recordset
Dim TB_Cores       As Recordset
Dim TB_Tamanhos    As Recordset
Dim TB_Edições     As Recordset
Dim TB_Pesquisa1   As Recordset
Dim TB_Pesquisa2   As Recordset
Dim TB_Pesquisa3   As Recordset
Dim Num_Registro   As Variant

Sub Mostra_Estoque()
 Dim Filial As Integer
 Dim Linha As String
 Dim Estoque As Double
 Dim Prod As String
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim Edição As Long
 Dim Aux As String
 Dim Aux2 As String
 Dim Cod_Completo As String
 Dim Fim As Integer
 Dim Tam_Str As String
 Dim Cor_Str As String

 Filial = 0

 Grade_Estoque.RemoveAll

 TB_Cores.Index = "Código"
 TB_Tamanhos.Index = "Código"

 TB_Grade.Index = "Original"
 TB_Parâmetros.Index = "Filial"
Lp1:
 TB_Parâmetros.Seek ">", Filial
 If TB_Parâmetros.NoMatch Then
    Grade_Estoque.Visible = True
    Exit Sub
 End If
 Filial = TB_Parâmetros("Filial")

 Linha = TB_Parâmetros("Nome")

 Estoque = 0

 If TB_Produtos("Tipo") = "N" Then
   Prod = TB_Produtos("Código")
   Tamanho = 0
   Cor = 0
   
   Estoque = Acha_Estoque(Filial, Prod, Tamanho, Cor, 0, 0)
   
   
   '--------------------------------------------------------------------------------
   '24/10/2002 - mpdea
   'Corrigido exibição do estoque para produtos fracionados
   
'   Estoque = Format(Estoque, "############0.00")
    If TB_Produtos("Fracionado") Then
      Estoque = Format(Estoque, "#0." & String(TB_Produtos("QtdeCasasDecimais"), "0"))
    Else
      Estoque = Format(Estoque, "#0")
    End If
   '--------------------------------------------------------------------------------
   
   
   Linha = "0 : 0:" + str$(Estoque) + ":" + Linha
   Grade_Estoque.Columns(0).Visible = False
   Grade_Estoque.Columns(1).Visible = False
   Grade_Estoque.AddItem Linha
 End If
 
 If TB_Produtos("Tipo") = "G" Then
   Grade_Estoque.Columns(0).Caption = "Tamanho"
   Grade_Estoque.Columns(1).Caption = "Cor"
   Prod = TB_Produtos("Código")
   Cod_Completo = 0
   Fim = False
   Do
     TB_Grade.Seek ">", Prod, Cod_Completo
     If TB_Grade.NoMatch Then Fim = True
     If Fim = False Then If TB_Grade("Código Original") <> Prod Then Fim = True
     If Fim = False Then
        Cod_Completo = TB_Grade("Código")
        Aux = TB_Grade("Código")
        Aux = LTrim(Aux)
        Aux2 = Right(Aux, 6)
        Tamanho = Left(Aux2, 3)
        Cor = Right(Aux, 3)
        
        Tam_Str = LTrim(str$(Tamanho))
        Tam_Str = "    000" + Tam_Str
        Tam_Str = Right$(Tam_Str, 3)
        If Tamanho <> 0 Then
          TB_Tamanhos.Seek "=", Tamanho
          If Not TB_Tamanhos.NoMatch Then
            Tam_Str = Tam_Str + "-" + TB_Tamanhos("Nome")
          End If
        End If
        
        Cor_Str = LTrim(str$(Cor))
        Cor_Str = "         000" + Cor_Str
        Cor_Str = Right$(Cor_Str, 3)
        If Cor <> 0 Then
          TB_Cores.Seek "=", Cor
          If Not TB_Cores.NoMatch Then
             Cor_Str = Cor_Str + "-" + TB_Cores("Nome")
          End If
        End If
          
        Estoque = Acha_Estoque(Filial, Prod, Tamanho, Cor, Edição, 0)
        Linha = Tam_Str + ":" + Cor_Str + ":" + str$(Estoque) + ":" + TB_Parâmetros("Nome")
        Grade_Estoque.AddItem Linha
        Grade_Estoque.Columns(0).Visible = True
        Grade_Estoque.Columns(1).Visible = True
     End If
   Loop While Fim = False
 End If
    
 If TB_Produtos("Tipo") = "E" Then
   Grade_Estoque.Columns(0).Visible = False
   Grade_Estoque.Columns(1).Caption = "Edição"
   Prod = TB_Produtos("Código")
   Cod_Completo = 0
   Fim = False
   TB_Edições.Index = "Produto"
   Do
     TB_Edições.Seek ">", Prod, Cod_Completo
     If TB_Edições.NoMatch Then Fim = True
     If Fim = False Then If TB_Edições("Produto") <> Prod Then Fim = True
     If Fim = False Then
        Cod_Completo = TB_Edições("Código")
        Estoque = Acha_Estoque(Filial, Prod, 0, 0, CLng(Cod_Completo), 0)
        Linha = " :" + str(Cod_Completo) + ":" + str$(Estoque) + ":" + TB_Parâmetros("Nome")
        Grade_Estoque.AddItem Linha
        'Grade_Estoque.Columns(0).Visible = True
        Grade_Estoque.Columns(1).Visible = True
     End If
   Loop While Fim = False
 
 End If
    
 GoTo Lp1

End Sub

Private Sub ShowRecord()
  Dim Cód As String
  Dim sSql As String
  Dim Tab1 As String
  Dim Tab2 As String
  Dim Tab3 As String
  Dim Tab4 As String
  Dim Tab5 As String
  Dim Tab6 As String

  Call StatusMsg("")

' On Error GoTo Processa_Erro

 TB_Parâmetros.Index = "Filial"
 TB_Parâmetros.Seek "=", gnCodFilial
 If TB_Parâmetros.NoMatch Then Exit Sub


  Con_Código.Text = TB_Produtos("Código")
  '19/11/2004 - Daniel
  'Bug na On Site - Os produtos não possuíam "Código Ordenação"
  If Len(TB_Produtos("Código Ordenação")) > 0 Then Ordenação.Caption = TB_Produtos("Código Ordenação")
  Con_Nome.Text = TB_Produtos("Nome") & ""
  Con_IPI.Caption = Format(TB_Produtos("Percentual IPI"), "##0.00")
  Con_Descrição.Text = TB_Produtos("Obs") & ""
  Con_Fabricante.Caption = TB_Produtos("Fabricante") & ""
  
  
  If TB_Produtos("Desativado") = True Then
    Inativo.Visible = True
  Else
    Inativo.Visible = False
  End If
  
  
  Con_Pesquisa1.Text = TB_Produtos("Pesquisa 1") & ""
  Con_Pesquisa1_LostFocus
  Con_Pesquisa2.Text = TB_Produtos("Pesquisa 2") & ""
  Con_Pesquisa2_LostFocus
  Con_Pesquisa3.Text = TB_Produtos("Pesquisa 3") & ""
  Con_Pesquisa3_LostFocus
  
  
  Classe.Caption = ""
  TB_Classes.Index = "Código"
  TB_Classes.Seek "=", TB_Produtos("Classe")
  If Not TB_Classes.NoMatch Then
    Classe.Caption = str(TB_Classes("Código")) + " - " + TB_Classes("Nome")
  End If
  
  Sub_Classe.Caption = ""
  TB_Sub_Classes.Index = "Código"
  TB_Sub_Classes.Seek "=", TB_Produtos("Sub Classe")
  If Not TB_Sub_Classes.NoMatch Then
    Sub_Classe.Caption = str(TB_Sub_Classes("Código")) + " - " + TB_Sub_Classes("Nome")
  End If
  
  Localização.Caption = TB_Produtos("Localização") & ""
  
  
  Tab1 = ""
  Tab2 = ""
  Tab3 = ""
  Tab4 = ""
  Tab5 = ""
  Tab6 = ""
  
  If Not IsNull(TB_Parâmetros("Consulta TAB1")) Then Tab1 = TB_Parâmetros("Consulta Tab1")
  If Not IsNull(TB_Parâmetros("Consulta TAB2")) Then Tab2 = TB_Parâmetros("Consulta Tab2")
  If Not IsNull(TB_Parâmetros("Consulta TAB3")) Then Tab3 = TB_Parâmetros("Consulta Tab3")
  If Not IsNull(TB_Parâmetros("Consulta TAB4")) Then Tab4 = TB_Parâmetros("Consulta Tab4")
  If Not IsNull(TB_Parâmetros("Consulta TAB5")) Then Tab5 = TB_Parâmetros("Consulta Tab5")
  If Not IsNull(TB_Parâmetros("Consulta TAB6")) Then Tab6 = TB_Parâmetros("Consulta Tab6")
  
     'Arruma Grade
  Cód = Con_Código.Text
  
  
  'Buscar tabelas de preços que o usuário logado tem acesso(vinculo)
  'Tabela: AcessoTabelasDePrecosProdutos
  Dim rsAcessosTabPrecoUsu As Recordset
  Dim iTemTabelasPreco As Integer
  
  iTemTabelasPreco = 0
  
  sSql = "Select Tabela From AcessoTabelasDePrecosProdutos Where Usuario=" & gnUserCode
  
  Set rsAcessosTabPrecoUsu = db.OpenRecordset(sSql, dbOpenDynaset)
  
  sSql = "SELECT Tabela, Preço FROM [Preços]"
  sSql = sSql + " WHERE Produto ='" + Cód + "'"
  
  If Not (rsAcessosTabPrecoUsu.EOF And rsAcessosTabPrecoUsu.BOF) Then
      iTemTabelasPreco = 1
      rsAcessosTabPrecoUsu.MoveFirst
      sSql = sSql + " AND ("
  End If
  While Not rsAcessosTabPrecoUsu.EOF
      sSql = sSql + " Tabela ='" & rsAcessosTabPrecoUsu.Fields(0).Value & "'"
  
      rsAcessosTabPrecoUsu.MoveNext
  
      If rsAcessosTabPrecoUsu.EOF Then
          sSql = sSql + " )"
      Else
          sSql = sSql + " OR "
      End If
  Wend
  rsAcessosTabPrecoUsu.Close
  Set rsAcessosTabPrecoUsu = Nothing
 
'''  sSql = "SELECT Tabela, Preço FROM [Preços]"
'''  sSql = sSql + " WHERE Produto ='" + Cód + "'"
'''
'''  sSql = sSql + " AND (Tabela ='" + Tab1 + "'"
'''  sSql = sSql + " OR Tabela ='" + Tab2 + "'"
'''  sSql = sSql + " OR Tabela ='" + Tab3 + "'"
'''  sSql = sSql + " OR Tabela ='" + Tab4 + "'"
'''  sSql = sSql + " OR Tabela ='" + Tab5 + "'"
'''  sSql = sSql + " OR Tabela ='" + Tab6 + "')"
  
 
  Call StatusMsg("Aguarde, montando tabela...")
  
  If iTemTabelasPreco = 1 Then
      Set Rec_Preços = db.OpenRecordset(sSql, dbOpenDynaset)
  Else
      Set Rec_Preços = db.OpenRecordset("SELECT Tabela, Preço FROM [Preços] WHERE Produto ='" + Cód + "' And Tabela='NADAxxxx'", dbOpenDynaset)
  End If
  
  Grade_Preço.DataMode = 1

  Set Data_Preço.Recordset = Rec_Preços

  Grade_Preço.Visible = False
  Grade_Preço.DataMode = 0

  Grade_Preço.ReBind
 
  '15/02/2007 - Anderson
  'Redução da janela para monitores 800X600 - a pedido do Paulo da Ribertec
  'Grade_Preço.Columns(0).Width = 1500 'Tabela
  Grade_Preço.Columns(0).Width = 125 'Tabela
  Grade_Preço.Columns(0).Locked = True
  '15/02/2007 - Anderson
  'Redução da janela para monitores 800X600 - a pedido do Paulo da Ribertec
  'Grade_Preço.Columns(1).Width = 1500 'Preços
  Grade_Preço.Columns(1).Width = 125 'Preços
  '''Grade_Preço.Columns(1).NumberFormat = "###,###,###0.00"
  Grade_Preço.Columns(1).Locked = True
  Grade_Preço.Visible = True
  
  Num_Registro = TB_Produtos.Bookmark

  Call StatusMsg("Aguarde, verificando estoque...")
  
  Mostra_Estoque
  
  Call StatusMsg("")

  imgFoto.Picture = LoadPicture("")
  
  If IsNull(TB_Produtos("Foto")) Then Exit Sub
  If TB_Produtos("Foto") = "" Then Exit Sub
    
  On Error Resume Next
  imgFoto.Picture = LoadPicture(gsConvertImagePath(TB_Produtos("Foto")))
  On Error GoTo 0
  Exit Sub

Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Consulta Produtos - Mostrar Registro")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select

End Sub

Private Sub B_Anterior_Click()
 Dim Atual As Variant
 Dim Atual2 As Variant

 On Error GoTo Processa_Erro

  Call StatusMsg("")
  
 Atual2 = Con_Código.Text
 If IsNull(Atual2) Then Atual2 = ""
 

 If O_Código = True Then
   Atual = Ordenação.Caption
   If IsNull(Atual) Then Atual = Gera_Ordenação("0")
   If Atual = "" Then Atual = Gera_Ordenação("0")

   TB_Produtos.Index = "Código Ordenação"
   
   TB_Produtos.Seek "<", Atual
 End If


 If O_Nome.Value = True Then
   Atual = Con_Nome.Text
   If IsNull(Atual) Then Atual = ""
   
   TB_Produtos.Index = "Nome"
   TB_Produtos.Seek "<", Atual, Gera_Ordenação(CStr(Atual2))
 End If
 
 
 
 If O_Pesquisa1.Value = True Then
   Atual = Con_Pesquisa1.Text
   If IsNull(Atual) Then Atual = 0
   
   TB_Produtos.Index = "Pesquisa 1"
   TB_Produtos.Seek "<", Atual, Atual2
 End If
 
 If O_Pesquisa2.Value = True Then
   Atual = Con_Pesquisa2.Text
   If IsNull(Atual) Then Atual = 0
   
   TB_Produtos.Index = "Pesquisa 2"
   TB_Produtos.Seek "<", Atual, Atual2
 End If
 
 If O_Pesquisa3.Value = True Then
   Atual = Con_Pesquisa3.Text
   If IsNull(Atual) Then Atual = 0
   
   TB_Produtos.Index = "Pesquisa 3"
   TB_Produtos.Seek "<", Atual, Atual2
 End If
 
 
 If TB_Produtos.NoMatch Then
    Beep
    If Not IsNull(Num_Registro) Then
      TB_Produtos.Bookmark = Num_Registro
    End If
    Exit Sub
 End If
 
 Num_Registro = TB_Produtos.Bookmark
 ShowRecord

 Exit Sub
 
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Consulta - Anterior")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select

End Sub

Private Sub B_Limpa_Click()
 Num_Registro = Null
 Con_Código.Text = ""
 Ordenação.Caption = ""
 Con_Nome.Text = ""
 Con_Descrição.Text = ""
 
 Con_Pesquisa1.Text = ""
 Con_Pesquisa1_LostFocus
 Con_Pesquisa2.Text = ""
 Con_Pesquisa2_LostFocus
 Con_Pesquisa3.Text = ""
 Con_Pesquisa3_LostFocus
 
 Inativo.Visible = False
 
 Con_Fabricante.Caption = ""
 Classe.Caption = ""
 Sub_Classe.Caption = ""
 Con_IPI.Caption = ""
 Con_Código.SetFocus
 
 Grade_Estoque.RemoveAll
 Grade_Estoque.Visible = False
 Grade_Preço.Visible = False

 imgFoto.Picture = LoadPicture("")

End Sub


Private Sub B_Próximo_Click()
 Dim Atual As Variant
 Dim Atual2 As Variant

 On Error GoTo Processa_Erro

  Call StatusMsg("")
  
 Atual2 = Con_Código.Text
 If IsNull(Atual2) Then Atual2 = ""
 'If Not IsNumeric(Atual2) Then Atual2 = ""
 'If Val(Atual2) < 0 Then Atual2 = ""

 If O_Código = True Then
   Atual = Ordenação.Caption
   If IsNull(Atual) Then Atual = Gera_Ordenação("0")
   If Atual = "" Then Atual = Gera_Ordenação("0")

   TB_Produtos.Index = "Código Ordenação"
   
   TB_Produtos.Seek ">", Atual
 End If

 If O_Nome.Value = True Then
   Atual = Con_Nome.Text
   If IsNull(Atual) Then Atual = ""
   
   TB_Produtos.Index = "Nome"
   TB_Produtos.Seek ">", Atual, Atual2
 End If
 
 If O_Pesquisa1.Value = True Then
   Atual = Con_Pesquisa1.Text
   If IsNull(Atual) Then Atual = 0
   
   TB_Produtos.Index = "Pesquisa 1"
   TB_Produtos.Seek ">", Atual, Atual2
 End If
 
 If O_Pesquisa2.Value = True Then
   Atual = Con_Pesquisa2.Text
   If IsNull(Atual) Then Atual = 0
   
   TB_Produtos.Index = "Pesquisa 2"
   TB_Produtos.Seek ">", Atual, Atual2
 End If
 
 If O_Pesquisa3.Value = True Then
   Atual = Con_Pesquisa3.Text
   If IsNull(Atual) Then Atual = 0
   
   TB_Produtos.Index = "Pesquisa 3"
   TB_Produtos.Seek ">", Atual, Atual2
 End If
 
 If TB_Produtos.NoMatch Then
    Beep
    If Not IsNull(Num_Registro) Then
      TB_Produtos.Bookmark = Num_Registro
    End If
    Exit Sub
 End If
 
 Num_Registro = TB_Produtos.Bookmark
 ShowRecord

 Exit Sub
 
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Consulta - Próximo")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select

End Sub

'31/08/2006 - Anderson
'Implementação de pesquisa avançada na tela de consulta do produto
Private Sub cmdPesquisar_Click()

  frmPesquisaProduto.strTipoPesquisa = "Pesquisa"
  frmPesquisaProduto.Show

End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

'17/09/2002 - mpdea
'Otimizado inclusão do item para venda
Private Sub Command2_Click()
  Dim intCancel As Integer
  Dim frmX As Form
  
  If Trim(Con_Código.Text) = "" Then Exit Sub
  
  Select Case nChamaConsulta
    Case 1, 2 'Previne outros códigos de chamada
      
      'Form de origem (chamadas comuns)
      If nChamaConsulta = 1 Then
        '18/01/2006 - mpdea
        'Alterado objeto frmVendaRap2 -> g_frmVendaRapida
        Set frmX = g_frmVendaRapida
      Else
        Set frmX = frmSaidas
      End If
      
      With frmX
        'Insere o item
        .Grade1.Columns(0).Text = Con_Código.Text
        .Grade1.Columns(1).Text = "1"
        'Atualiza grid
        .Grade1_BeforeColUpdate 0, "", intCancel
        If intCancel = -1 Then Exit Sub
        'Calcula totais
        .Calcula_Linha
        .Recalcula
        'Move para a próxima linha
        .Grade1.MoveNext
        .Grade1.DoClick
      End With
      
      Set frmX = Nothing
      
    '04/11/2009 - mpdea
    'Tela de Entradas
    Case 3
      With frmEntrada
        'Remove as linhas em branco
        Dim Str_Aux As String
        Dim bm As Variant
        Dim nRow As Long
        For nRow = .grdItens.Rows - 1 To 0 Step -1
          bm = .grdItens.AddItemBookmark(nRow)
          .grdItens.Bookmark = bm
          Str_Aux = gsHandleNull(.grdItens.Columns("Código").CellText(bm))
          If (Str_Aux = "0" Or Str_Aux = "") And Not IsEmpty(bm) Then
            .grdItens.RemoveItem .grdItens.AddItemRowIndex(bm)
          End If
        Next nRow
        .grdItens.Scroll -99, -99
        .grdItens.Update
        
        'Insere o item
        .grdItens.AddItem Con_Código.Text & vbTab & "1"
        'Atualiza grid
        .grdItens.MoveLast
        .grdItens_BeforeColUpdate 0, "", intCancel
        If intCancel = -1 Then Exit Sub
        'Calcula totais
        .Calcula_Linha
        .Recalcula
        .grdItens.Update
      End With
      
  End Select
     
End Sub

Private Sub Con_Código_GotFocus()
  Con_Código.SelStart = 0
  Con_Código.SelLength = Len(Con_Código.Text & "")
End Sub

Private Sub Con_Código_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
  End If
End Sub

Public Sub Con_Código_LostFocus()
'  '26/05/2004 - Daniel
'  'Criado uma validação caso o código do produto digitado não exista
'  'ele não vai mostrar qualquer produto aleatóriamente, irá avisar e sair da rotina
'  Dim rstProdutos    As Recordset
'  Dim blnCairFora    As Boolean
'  Dim blnUsaGrade    As Boolean
'  Dim strCodSemGrade As String
'
'  '13/09/2004 - Daniel
'  'Tratamento para casos em que o produto usa grade
'  '
'  'PPPPP = PRINCIPAL
'  '
'  'OOO = OPCIONAL
'  '
'  'TTT = TAMANHO
'  '
'  'CCC = COR
'  '
'  'PPPPPOOOTTTCCC
'  '
'  '04/11/2004 - Daniel
'  'Tratamento referente ao uso da grade
'  If Len(Con_Código.Text) <= 0 Then Exit Sub
'
'  If UsaGrade(Con_Código.Text) Then blnUsaGrade = True
'
'  '------------------------------------------------------------------------------
'  If Len(Con_Código.Text) > 0 And Not blnUsaGrade Then  'digitou algo
'    Set rstProdutos = db.OpenRecordset("SELECT Código FROM Produtos WHERE Código = '" & CStr(Trim(Con_Código.Text)) & "'", dbOpenDynaset)
'
'    With rstProdutos
'      If .RecordCount = 0 Then blnCairFora = True
'      .Close
'    End With
'
'    Set rstProdutos = Nothing
'
'    If blnCairFora Then
'      MsgBox "Produto com o Código '" & (Con_Código.Text) & "' não está cadastrado.", vbExclamation, "Quick Store"
'      Exit Sub 'Não continua o processo
'    End If
'  End If
'  '------------------------------------------------------------------------------
'
'  If IsNull(Con_Código.Text) Then Exit Sub
'  If Con_Código.Text = "" Then Exit Sub
'  Con_Código.Text = UCase(Con_Código.Text)
'  '26/05/2004 - Daniel
'  'Tratamento para 0 'zero' a esquerda
'  If Not gbZeroEsquerda Then
'    Con_Código.Text = Retira_Zeros(Con_Código.Text)
'  End If
'
'  '13/09/2004 - Daniel
'  'Tratamento para casos em que o produto usa grade
'  If blnUsaGrade Then
'    Call BuscarCodSemGrade(Con_Código.Text, strCodSemGrade)
'
'    TB_Produtos.Index = "Código"
'    TB_Produtos.Seek ">=", strCodSemGrade
'    If Not TB_Produtos.NoMatch Then
'      Call ShowRecord
'    End If
'
'  Else
'    TB_Produtos.Index = "Código"
'    TB_Produtos.Seek ">=", Con_Código.Text
'    If Not TB_Produtos.NoMatch Then
'      Call ShowRecord
'    End If
'  End If
  
  
  
  
  '----------------------------------------------------------------------------
  '09/12/2005 - mpdea
  'Corrigido pesquisa de produtos com grade
  '----------------------------------------------------------------------------
  Dim strCodigo As String
  
  If IsNull(Con_Código.Text) Then Exit Sub
  If Con_Código.Text = "" Then Exit Sub
  Con_Código.Text = UCase(Con_Código.Text)
  If Not gbZeroEsquerda Then Con_Código.Text = Retira_Zeros(Con_Código.Text)
  
  If Not BuscarCodSemGrade(Con_Código.Text, strCodigo) Then
    strCodigo = Con_Código.Text
  End If
  
  TB_Produtos.Index = "Código"
  TB_Produtos.Seek ">=", strCodigo
  If Not TB_Produtos.NoMatch Then
    Call ShowRecord
  End If
  '----------------------------------------------------------------------------
  
End Sub


Private Sub Con_Pesquisa1_CloseUp()
 Con_Pesquisa1.Text = Con_Pesquisa1.Columns(1).Text
 Con_Pesquisa1_LostFocus
End Sub

Private Sub Con_Pesquisa1_LostFocus()

  Nome_Pesq1.Caption = ""
  If IsNull(Con_Pesquisa1.Text) Then Exit Sub
  If Con_Pesquisa1.Text = "" Then Exit Sub
  If Not IsNumeric(Con_Pesquisa1.Text) Then Exit Sub
  If Val(Con_Pesquisa1.Text) > 99999 Then Exit Sub
  If Val(Con_Pesquisa1.Text) < 1 Then Exit Sub
  
  TB_Pesquisa1.Index = "Código"
  TB_Pesquisa1.Seek "=", Con_Pesquisa1.Text
  If TB_Pesquisa1.NoMatch Then Exit Sub
  
  Nome_Pesq1.Caption = TB_Pesquisa1("Nome") & ""
  
End Sub

Private Sub Con_Pesquisa2_CloseUp()
 Con_Pesquisa2.Text = Con_Pesquisa2.Columns(1).Text
 Con_Pesquisa2_LostFocus
End Sub

Private Sub Con_Pesquisa2_LostFocus()

  Nome_Pesq2.Caption = ""
  If IsNull(Con_Pesquisa2.Text) Then Exit Sub
  If Con_Pesquisa2.Text = "" Then Exit Sub
  If Not IsNumeric(Con_Pesquisa2.Text) Then Exit Sub
  If Val(Con_Pesquisa2.Text) > 99999 Then Exit Sub
  If Val(Con_Pesquisa2.Text) < 1 Then Exit Sub
  
  TB_Pesquisa2.Index = "Código"
  TB_Pesquisa2.Seek "=", Con_Pesquisa2.Text
  If TB_Pesquisa2.NoMatch Then Exit Sub
  
  Nome_Pesq2.Caption = TB_Pesquisa2("Nome") & ""
  
End Sub

Private Sub Con_Pesquisa3_CloseUp()
 Con_Pesquisa3.Text = Con_Pesquisa3.Columns(1).Text
 Con_Pesquisa3_LostFocus
End Sub

Private Sub Con_Pesquisa3_LostFocus()

  Nome_Pesq3.Caption = ""
  If IsNull(Con_Pesquisa3.Text) Then Exit Sub
  If Con_Pesquisa3.Text = "" Then Exit Sub
  If Not IsNumeric(Con_Pesquisa3.Text) Then Exit Sub
  If Val(Con_Pesquisa3.Text) > 99999 Then Exit Sub
  If Val(Con_Pesquisa3.Text) < 1 Then Exit Sub
  
  TB_Pesquisa3.Index = "Código"
  TB_Pesquisa3.Seek "=", Con_Pesquisa3.Text
  If TB_Pesquisa3.NoMatch Then Exit Sub
  
  Nome_Pesq3.Caption = TB_Pesquisa3("Nome") & ""
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  
  '24/09/2002 - mpdea
  'Comentado - substituído por Esc (padrão)
'  If KeyCode = vbKeyF3 Then
'    Unload Me
'  End If


If KeyCode = vbKeyF7 And Con_Pesquisa1.Visible = True Then
   Con_Pesquisa1.SetFocus
End If

If KeyCode = vbKeyF8 And Con_Pesquisa2.Visible = True Then
   Con_Pesquisa2.SetFocus
End If

If KeyCode = vbKeyF9 And Con_Pesquisa3.Visible = True Then
   Con_Pesquisa3.SetFocus
End If


If KeyCode = vbKeyF11 Then
   Call B_Anterior_Click
   Con_Código.SetFocus
End If

If KeyCode = vbKeyF12 Then
   Call B_Próximo_Click
End If

If KeyCode = vbKeyEnd Then
   Call B_Limpa_Click
End If

End Sub

Private Sub Form_Load()
 Dim rstFuncionarios As Recordset
 
 If nChamaConsulta = 0 Then Command2.Visible = False
  
 Call CenterForm(Me)
  
 Set TB_Produtos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set TB_Preços = db.OpenRecordset("Preços", , dbReadOnly)
 Set TB_Estoque = db.OpenRecordset("Estoque", , dbReadOnly)
 Set TB_Parâmetros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set TB_Classes = db.OpenRecordset("Classes", , dbReadOnly)
 Set TB_Sub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
 Set TB_Grade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
 Set TB_Cores = db.OpenRecordset("Cores", , dbReadOnly)
 Set TB_Tamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
 Set TB_Edições = db.OpenRecordset("Edições", , dbReadOnly)

 Set TB_Pesquisa1 = db.OpenRecordset("Pesquisa 1", , dbReadOnly)
 Set TB_Pesquisa2 = db.OpenRecordset("Pesquisa 2", , dbReadOnly)
 Set TB_Pesquisa3 = db.OpenRecordset("Pesquisa 3", , dbReadOnly)
 
 Num_Registro = Null
 
 TB_Parâmetros.Index = "Filial"
 TB_Parâmetros.Seek "=", gnCodFilial
 If TB_Parâmetros.NoMatch Then Exit Sub
 
 Data_Preço.DatabaseName = gsQuickDBFileName
 Data4.DatabaseName = gsQuickDBFileName
 Data5.DatabaseName = gsQuickDBFileName
 Data6.DatabaseName = gsQuickDBFileName
 
 If gsPesq1 = "" Then
   Con_Pesquisa1.Visible = False
   Nome_Pesq1.Visible = False
   Label_Pesq1.Visible = False
   O_Pesquisa1.Visible = False
 Else
   Label_Pesq1.Caption = gsPesq1
   O_Pesquisa1.Caption = gsPesq1
 End If
 
 If gsPesq2 = "" Then
   Con_Pesquisa2.Visible = False
   Nome_Pesq2.Visible = False
   Label_Pesq2.Visible = False
   O_Pesquisa2.Visible = False
 Else
   Label_Pesq2.Caption = gsPesq2
   O_Pesquisa2.Caption = gsPesq2
 End If
 
 If gsPesq3 = "" Then
   Con_Pesquisa3.Visible = False
   Nome_Pesq3.Visible = False
   Label_Pesq3.Visible = False
   O_Pesquisa3.Visible = False
 Else
   Label_Pesq3.Caption = gsPesq3
   O_Pesquisa3.Caption = gsPesq3
 End If
 
 Con_Código.Text = gsCodProduto
 Call Con_Código_LostFocus
 
 If Len(Con_Código.Text) > 0 Then Con_Código_LostFocus
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '02/06/2006 - mpdea
  'Reseta código de chamada para consulta do produto
  'Utilizado para inserir produto em Venda Rápida e Saídas
  nChamaConsulta = 0
End Sub

Private Sub Grade_Estoque_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  
  Select Case TB_Produtos.Fields("Tipo").Value
    Case "G"
      With Grade_Estoque.Columns
        Con_Código.Text = TB_Produtos.Fields("Código").Value & _
          Format(Left(.Item(0).Text, 3), "000") & Format(Left(.Item(1).Text, 3), "000")
      End With
      
    Case "E"
      With Grade_Estoque.Columns
        Con_Código.Text = TB_Produtos.Fields("Código").Value & _
          Format(.Item(1).Text, "00000")
      End With
      
  End Select
  
End Sub

'09/12/2005 - mpdea
'Modificado para função retornando se encontrou a grade do produto
Private Function BuscarCodSemGrade(ByVal Cod As String, ByRef CodSem As String) As Boolean
  Dim rstCodigosDaGrade As Recordset
  Dim strSQL            As String
  
  '09/12/2005 - mpdea
  'Corrigido comparação com o código
  strSQL = "SELECT * FROM [Códigos da Grade]"
  strSQL = strSQL & " WHERE Código = '" & Cod & "'"
  
  '09/12/2005 - mpdea
  'Incluído parâmetro read only
  Set rstCodigosDaGrade = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  With rstCodigosDaGrade
    If Not (.BOF And .EOF) Then
      .MoveFirst
      CodSem = .Fields("Código Original").Value & ""
      BuscarCodSemGrade = True
    End If
    .Close
  End With
  
  Set rstCodigosDaGrade = Nothing
  
End Function

'09/12/2005 - mpdea
'Comentado função não utilizada
'Private Function UsaGrade(ByVal CodProduto As String) As Boolean
'  '04/11/2004 - Daniel
'  'Alterada a função para avaliar o uso da grade
'  'pela tabela de Produtos e não mais pela Parâmetros
'  Dim rstProdutos As Recordset
'  Dim strQuery    As String
'
'  strQuery = "SELECT Tipo FROM Produtos "
'  strQuery = strQuery & " WHERE Código = '" & CodProduto & "'"
'
'  Set rstProdutos = db.OpenRecordset(strQuery, dbOpenDynaset)
'
'  With rstProdutos
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'
'      If .Fields("Tipo").Value = "G" Then UsaGrade = True
'
'    End If
'    .Close
'  End With
'
'  Set rstProdutos = Nothing
'
'End Function
