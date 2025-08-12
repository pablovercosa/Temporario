VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmDevolucoes 
   Caption         =   " Devoluções - Vale Créditos"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDevolucoes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   16365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6060
      TabIndex        =   23
      Top             =   6960
      Width           =   4695
      Begin VB.CommandButton cmd_GerarDevolucao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gerar Nova Devolução"
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
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   270
         Width           =   2610
      End
      Begin VB.OptionButton opt_afetaComissao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Afeta comissão"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         TabIndex        =   25
         Top             =   210
         Width           =   1455
      End
      Begin VB.OptionButton opt_naoAfetaComissao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Não afeta comissão"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         TabIndex        =   24
         Top             =   510
         Value           =   -1  'True
         Width           =   1845
      End
   End
   Begin VB.ComboBox cboOrdenacao 
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
      ItemData        =   "frmDevolucoes.frx":4E95A
      Left            =   13020
      List            =   "frmDevolucoes.frx":4E97C
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   548
      Width           =   3300
   End
   Begin VB.TextBox txt_total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14670
      TabIndex        =   19
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   7020
      Width           =   1605
   End
   Begin VB.CommandButton cmd_gravarStatusEmAberto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salvar como 'Em aberto'"
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
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7020
      Width           =   2910
   End
   Begin VB.CommandButton cmd_gravarStatusRecebido 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salvar como 'Recebido'"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7020
      Width           =   2910
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   990
      Width           =   16260
   End
   Begin VB.TextBox txt_sequenciaEntrada 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9780
      MaxLength       =   12
      TabIndex        =   14
      Top             =   135
      Width           =   1705
   End
   Begin VB.TextBox txt_sequenciaVenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9780
      MaxLength       =   12
      TabIndex        =   13
      Top             =   570
      Width           =   1705
   End
   Begin VB.ComboBox cboStatus 
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
      ItemData        =   "frmDevolucoes.frx":4E9F4
      Left            =   5010
      List            =   "frmDevolucoes.frx":4EA01
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   113
      Width           =   2580
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
      Left            =   1650
      Picture         =   "frmDevolucoes.frx":4EA1C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   465
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
      Left            =   3825
      Picture         =   "frmDevolucoes.frx":4F2FE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   465
   End
   Begin VB.Data Data4 
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
      Left            =   5700
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cli_For"
      Top             =   570
      Visible         =   0   'False
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo cboCliente 
      Bindings        =   "frmDevolucoes.frx":4FBE0
      DataSource      =   "Data4"
      Height          =   330
      Left            =   1065
      TabIndex        =   0
      Top             =   570
      Width           =   1005
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      MaxDropDownItems=   16
      BevelType       =   0
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
      BorderStyle     =   0
      BevelColorFace  =   15066597
      CheckBox3D      =   0   'False
      ForeColorEven   =   0
      BackColorEven   =   15066597
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   9075
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   847
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Tipo"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4339
      Columns(3).Caption=   "Cidade"
      Columns(3).Name =   "Cidade"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Cidade"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1191
      Columns(4).Caption=   "Estado"
      Columns(4).Name =   "Estado"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Estado"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   0
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox Data_Fim 
      Height          =   315
      Left            =   2505
      TabIndex        =   5
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   150
      Width           =   1290
      _ExtentX        =   2275
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
      Left            =   330
      TabIndex        =   6
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   150
      Width           =   1275
      _ExtentX        =   2249
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
   Begin SSDataWidgets_B.SSDBGrid grdDevolucoes 
      Height          =   5445
      Left            =   60
      TabIndex        =   16
      Top             =   1500
      Width           =   16260
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
      Col.Count       =   16
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      ExtraHeight     =   212
      Columns.Count   =   16
      Columns(0).Width=   556
      Columns(0).Name =   "Marca"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   1799
      Columns(1).Caption=   "Devolução"
      Columns(1).Name =   "DtDevolucao"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1799
      Columns(2).Caption=   "Status"
      Columns(2).Name =   "Status"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   873
      Columns(3).Caption=   "Oper"
      Columns(3).Name =   "CodigoOperacaoEntrada"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3863
      Columns(4).Caption=   "Produto"
      Columns(4).Name =   "CodProduto"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   7170
      Columns(5).Caption=   "Nome"
      Columns(5).Name =   "NomeProduto"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1429
      Columns(6).Caption=   "Qtde"
      Columns(6).Name =   "Quantidade"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1905
      Columns(7).Caption=   "Vlr Unitário"
      Columns(7).Name =   "valorUnitario"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1905
      Columns(8).Caption=   "Vlr Total"
      Columns(8).Name =   "ValorTotal"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1958
      Columns(9).Caption=   "Cliente"
      Columns(9).Name =   "CodigoCliente"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   6271
      Columns(10).Caption=   "Nome"
      Columns(10).Name=   "NomeCliente"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1429
      Columns(11).Caption=   "Digitador"
      Columns(11).Name=   "CodigoDigitador"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   4233
      Columns(12).Caption=   "Nome"
      Columns(12).Name=   "NomeDigitador"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   2249
      Columns(13).Caption=   "Seq.Devolução"
      Columns(13).Name=   "SequenciaDevolucao"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   2249
      Columns(14).Caption=   "Seq.Venda"
      Columns(14).Name=   "SequenciaVenda"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   13785
      Columns(15).Caption=   "Observações"
      Columns(15).Name=   "Observacoes"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      _ExtentX        =   28681
      _ExtentY        =   9604
      _StockProps     =   79
      ForeColor       =   0
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Ordenar por"
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
      Left            =   12000
      TabIndex        =   22
      Top             =   615
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total R$ "
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
      Left            =   13800
      TabIndex        =   20
      Top             =   7080
      Width           =   780
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Sequência Entrada"
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
      Left            =   8220
      TabIndex        =   12
      Top             =   180
      Width           =   1530
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Sequência Venda"
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
      Left            =   8310
      TabIndex        =   11
      Top             =   615
      Width           =   1425
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   180
      Width           =   525
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
      Left            =   2220
      TabIndex        =   8
      Top             =   180
      Width           =   270
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
      Left            =   60
      TabIndex        =   7
      Top             =   180
      Width           =   225
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2085
      TabIndex        =   2
      Top             =   570
      Width           =   5505
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cliente/Forn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00666666&
      Height          =   210
      Left            =   60
      TabIndex        =   1
      Top             =   615
      Width           =   990
   End
End
Attribute VB_Name = "frmDevolucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCliFor As Recordset

Private Sub cboCliente_Click()
  cboCliente.Text = cboCliente.Columns(1).Text
End Sub

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(1).Text
  cboCliente_LostFocus
End Sub

Public Sub cboCliente_LostFocus() '16/04/2004 - Daniel - Mudado para Public
  Dim Aux As Variant
  
  Nome_Cliente.Caption = ""
  
  Aux = cboCliente.Text
  If IsNull(Aux) Then Exit Sub
  If Aux = "" Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) < 1 Then Exit Sub
  If Val(Aux) > 99999999 Then Exit Sub
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(Aux)
  If rsCliFor.NoMatch Then
    'Somente exibe o aviso se não estiver em navegação dos registros
    DisplayMsg "Cliente incorreto."
    cboCliente.SetFocus
    Exit Sub
  End If
  
  Nome_Cliente.Caption = rsCliFor("Nome") & ""
  
End Sub


Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub cmd_GerarDevolucao_Click()
    Dim objEntrada As frmEntrada
    Set objEntrada = New frmEntrada
    
    If opt_afetaComissao.Value = True Then
        objEntrada.bTelaChamadoraDevolucao_ValeCredito = True
        objEntrada.bGerarDevolucaoAfentandoComissao = True
        objEntrada.Show
    Else
        objEntrada.bTelaChamadoraDevolucao_ValeCredito = True
        objEntrada.bGerarDevolucaoAfentandoComissao = False
        objEntrada.Show
    End If
End Sub

Private Sub cmd_gravarStatusEmAberto_Click()
  On Error GoTo ErrHandler

  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim sCodProduto As String
  Dim sQtde As String
  Dim sValorTotal As String
  Dim sSql As String

  With grdDevolucoes
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)
    
          If CBool(.Columns("Marca").CellValue(bm)) = True Then
              lngSequencia = CLng(gsHandleNull(.Columns("SequenciaDevolucao").CellValue(bm)))
              sCodProduto = .Columns("CodProduto").CellValue(bm)
              sQtde = .Columns("Quantidade").CellValue(bm)
              sValorTotal = .Columns("ValorTotal").CellValue(bm)
              
              If .Columns("Status").CellValue(bm) = "Recebido" Then
                  sQtde = Replace(sQtde, ",", ".")
                  sSql = "Update [Entradas - Produtos] set ConsignacaoFechada = 0 "
                  sSql = sSql & " Where Filial = " & gnCodFilial
                  sSql = sSql & " And Sequência = " & lngSequencia
                  sSql = sSql & " And Código = '" & sCodProduto & "' "
                  sSql = sSql & " And Qtde = " & sQtde
                  
                  db.Execute sSql, dbFailOnError
                  
                  'LOG *****************
                  sSql = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
                  sSql = sSql & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & lngSequencia & " Prd:" & sCodProduto & " Qtde:" & sQtde & " Vlr:" & sValorTotal & " EM ABERTO", 80) & "', 'DEVOLUCAO_VALE ABERT')"
                  db.Execute sSql, dbFailOnError
                  'fim *******************
              End If
          End If
          
      Next nRow
  End With
  
  MsgBox "Gravado com sucesso", vbInformation, "Sucesso"
  
  Exit Sub

ErrHandler:

  MsgBox "Erro ao gravar status " & Err.Number & " (" & Err.Description & "). Sequência: " & lngSequencia & ".", vbCritical, "Erro"

End Sub

Private Sub cmd_gravarStatusRecebido_Click()
  On Error GoTo ErrHandler

  Dim bm As Variant
  Dim nRow As Long
  Dim lngSequencia As Long
  Dim sCodProduto As String
  Dim sQtde As String
  Dim sValorTotal As String
  Dim sSql As String

  With grdDevolucoes
      For nRow = 0 To .Rows - 1
          bm = .AddItemBookmark(nRow)
    
          If CBool(.Columns("Marca").CellValue(bm)) = True Then
              lngSequencia = CLng(gsHandleNull(.Columns("SequenciaDevolucao").CellValue(bm)))
              sCodProduto = .Columns("CodProduto").CellValue(bm)
              sQtde = .Columns("Quantidade").CellValue(bm)
              sValorTotal = .Columns("ValorTotal").CellValue(bm)
              
              If .Columns("Status").CellValue(bm) = "Em Aberto" Then
                  sQtde = Replace(sQtde, ",", ".")
                  sSql = "Update [Entradas - Produtos] set ConsignacaoFechada = -1 "
                  sSql = sSql & " Where Filial = " & gnCodFilial
                  sSql = sSql & " And Sequência = " & lngSequencia
                  sSql = sSql & " And Código = '" & sCodProduto & "' "
                  sSql = sSql & " And Qtde = " & sQtde
                  
                  db.Execute sSql, dbFailOnError
                  
                  'LOG *****************
                  sSql = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
                  sSql = sSql & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & lngSequencia & " Prd:" & sCodProduto & " Qtde:" & sQtde & " Vlr:" & sValorTotal & " RECEBIDO", 80) & "', 'DEVOLUCAO_VALE RECEB')"
                  db.Execute sSql, dbFailOnError
                  'fim *******************
              End If
          End If
          
      Next nRow
  End With
  
  MsgBox "Gravado com sucesso", vbInformation, "Sucesso"
  
  Exit Sub

ErrHandler:

  MsgBox "Erro ao gravar status " & Err.Number & " (" & Err.Description & "). Sequência: " & lngSequencia & ".", vbCritical, "Erro"
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo Erro

  Dim rsEntradas As Recordset
  Dim strSQL As String
  Dim strStatus As String
  Dim bUsaCondicao As Boolean
  Dim dTotalGrade As Double
  Dim sValorTotalGrade As String

  With grdDevolucoes
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With

  bUsaCondicao = True
  dTotalGrade = 0
   
  strSQL = "SELECT E.Data, E.Operação, E.Sequência, E.ChaveReferenciada, EP.Código, P.Nome, EP.Qtde, EP.Preço, EP.[Preço Final], E.Fornecedor, C.Nome, E.Digitador, U.Nome, EP.ConsignacaoFechada, E.Observações "
  strSQL = strSQL & " FROM Entradas E, [Entradas - produtos] EP, Produtos P, Cli_For C, Funcionários U"
  strSQL = strSQL & " WHERE E.Filial = " & gnCodFilial
  strSQL = strSQL & " AND (E.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
  strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
  
  If Trim(txt_sequenciaEntrada.Text) <> "" Then
      If IsNumeric(txt_sequenciaEntrada.Text) Then
          strSQL = strSQL & " AND E.Sequência = " & Trim(txt_sequenciaEntrada.Text)
          bUsaCondicao = False
      End If
  End If
  
  If Trim(txt_sequenciaVenda.Text) <> "" And bUsaCondicao = True Then
      If IsNumeric(txt_sequenciaVenda.Text) Then
          strSQL = strSQL & " AND E.ChaveReferenciada = '" & Trim(txt_sequenciaVenda.Text) & "' "
          bUsaCondicao = False
      End If
  End If
  
  If cboStatus.Text <> "" And bUsaCondicao = True Then
      ' Em Aberto    ou    Recebido
      If cboStatus.Text = "Recebido" Then
          strSQL = strSQL & " AND EP.ConsignacaoFechada = -1 "
      Else
          strSQL = strSQL & " AND EP.ConsignacaoFechada = 0 "
          '''strSQL = strSQL & " AND (E.Obs_Obs1 = null or E.Obs_Obs1 = '' or E.Obs_Obs1 = 'Em Aberto') "
      End If
  End If
  
  strSQL = strSQL & " AND E.Filial = EP.Filial AND E.Sequência = EP.Sequência "
  strSQL = strSQL & " AND E.Operação in (-1,-2) "
  strSQL = strSQL & " AND E.Digitador = U.Código "
  
  If Trim(cboCliente.Text) <> "" And bUsaCondicao = True Then
      If IsNumeric(cboCliente.Text) Then
          strSQL = strSQL & " AND E.Fornecedor = " & Trim(cboCliente.Text)
      End If
  End If
  
  strSQL = strSQL & " AND E.Fornecedor = C.Código "
  strSQL = strSQL & " AND EP.Código = P.Código "

  ' Ordenar por:
  '    Data Crescente
  '    Data Decrescente
  '    Sequência
  '    Cliente
  '    Digitador
  '    Quantidade
  '    Valor Unitário
  '    Valor Total
  '    Status
  '''  E.Data, E.Operação, E.Sequência, E.ChaveReferenciada, EP.Código, P.Nome, EP.Qtde, EP.Preço, EP.[Preço Final], E.Fornecedor, C.Nome, E.Digitador, U.Nome, E.Obs_Obs1, E.Observações
  If cboOrdenacao.Text = "Data Crescente" Then
      strSQL = strSQL & " ORDER BY E.Data "
  ElseIf cboOrdenacao.Text = "Data Decrescente" Then
      strSQL = strSQL & " ORDER BY E.Data DESC "
  ElseIf cboOrdenacao.Text = "Sequência" Then
      strSQL = strSQL & " ORDER BY E.Sequência "
  ElseIf cboOrdenacao.Text = "Cliente" Then
      strSQL = strSQL & " ORDER BY E.Fornecedor "
  ElseIf cboOrdenacao.Text = "Digitador" Then
      strSQL = strSQL & " ORDER BY E.Digitador "
  ElseIf cboOrdenacao.Text = "Quantidade" Then
      strSQL = strSQL & " ORDER BY EP.Qtde "
  ElseIf cboOrdenacao.Text = "Valor Unitário" Then
      strSQL = strSQL & " ORDER BY EP.Preço "
  ElseIf cboOrdenacao.Text = "Valor Total" Then
      strSQL = strSQL & " ORDER BY EP.[Preço Final] "
  ElseIf cboOrdenacao.Text = "Status" Then
      strSQL = strSQL & " ORDER BY EP.ConsignacaoFechada DESC "
  End If
  
  
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsEntradas
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        '''If IsNull(rsEntradas.Fields("Obs_Obs1").Value) Or rsEntradas.Fields("Obs_Obs1").Value = "" Then
        If IsNull(rsEntradas.Fields("ConsignacaoFechada").Value) Or rsEntradas.Fields("ConsignacaoFechada").Value = 0 Then
          strStatus = "Em Aberto"
        Else
          strStatus = "Recebido"
        End If
  
        'E.Data,            E.Operação,   E.Sequência,  E.ChaveReferenciada,
        'EP.Código,         P.Nome,       EP.Qtde,      EP.Preço,
        'EP.[Preço Final],  E.Fornecedor, C.Nome,       E.Digitador,
        'U.Nome,            E.Obs_Obs1,   E.Observações
        

        'Adiciona registro
        grdDevolucoes.AddItem "0" & vbTab & _
                              .Fields(0).Value & vbTab & _
                              strStatus & vbTab & _
                              .Fields(1).Value & vbTab & _
                              .Fields(4).Value & vbTab & _
                              .Fields(5).Value & vbTab & _
                              .Fields(6).Value & vbTab & _
                              Format(.Fields(7).Value, FORMAT_VALUE) & vbTab & _
                              Format(.Fields(8).Value, FORMAT_VALUE) & vbTab & _
                              .Fields(9).Value & vbTab & _
                              .Fields(10).Value & vbTab & _
                              .Fields(11).Value & vbTab & _
                              .Fields(12).Value & vbTab & _
                              .Fields(2).Value & vbTab & _
                              .Fields(3).Value & vbTab & _
                              .Fields(14).Value
        
        sValorTotalGrade = Format(.Fields(8).Value, FORMAT_VALUE)
        dTotalGrade = dTotalGrade + CDbl(sValorTotalGrade)
        
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsEntradas = Nothing
  
  ' ---------------------------------------------------------------
  ' Buscar devoluções de produto com grade (se houver)
  Dim rsTamanho As Recordset
  Dim rsCor As Recordset
  Dim sProdutoGradeAux As String
  Dim sTamanho As String
  Dim sCor As String
  
  Set rsTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsCor = db.OpenRecordset("Cores", , dbReadOnly)
  
  strSQL = "SELECT E.Data, E.Operação, E.Sequência, E.ChaveReferenciada, EP.Código, P.Nome, EP.Qtde, EP.Preço, EP.[Preço Final], E.Fornecedor, C.Nome, E.Digitador, U.Nome, EP.ConsignacaoFechada, E.Observações "
  strSQL = strSQL & " FROM Entradas E, [Entradas - produtos] EP, Produtos P, Cli_For C, Funcionários U, [Códigos da Grade] PG "
  strSQL = strSQL & " WHERE E.Filial = " & gnCodFilial
  strSQL = strSQL & " AND (E.Data BETWEEN #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "# "
  strSQL = strSQL & " AND #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#) "
  
  If Trim(txt_sequenciaEntrada.Text) <> "" Then
      If IsNumeric(txt_sequenciaEntrada.Text) Then
          strSQL = strSQL & " AND E.Sequência = " & Trim(txt_sequenciaEntrada.Text)
          bUsaCondicao = False
      End If
  End If
  
  If Trim(txt_sequenciaVenda.Text) <> "" And bUsaCondicao = True Then
      If IsNumeric(txt_sequenciaVenda.Text) Then
          strSQL = strSQL & " AND E.ChaveReferenciada = '" & Trim(txt_sequenciaVenda.Text) & "' "
          bUsaCondicao = False
      End If
  End If
  
  If cboStatus.Text <> "" And bUsaCondicao = True Then
      ' Em Aberto    ou    Recebido
      If cboStatus.Text = "Recebido" Then
          strSQL = strSQL & " AND EP.ConsignacaoFechada = -1 "
      Else
          strSQL = strSQL & " AND EP.ConsignacaoFechada = 0 "
          '''strSQL = strSQL & " AND (E.Obs_Obs1 = null or E.Obs_Obs1 = '' or E.Obs_Obs1 = 'Em Aberto') "
      End If
  End If
  
  strSQL = strSQL & " AND E.Filial = EP.Filial AND E.Sequência = EP.Sequência "
  strSQL = strSQL & " AND E.Operação in (-1,-2) "
  strSQL = strSQL & " AND E.Digitador = U.Código "
  
  If Trim(cboCliente.Text) <> "" And bUsaCondicao = True Then
      If IsNumeric(cboCliente.Text) Then
          strSQL = strSQL & " AND E.Fornecedor = " & Trim(cboCliente.Text)
      End If
  End If
  
  strSQL = strSQL & " AND E.Fornecedor = C.Código "
  strSQL = strSQL & " and EP.Código = PG.Código"
  strSQL = strSQL & " and PG.[Código Original] = P.Código"

  ' Ordenar por:
  '    Data Crescente
  '    Data Decrescente
  '    Sequência
  '    Cliente
  '    Digitador
  '    Quantidade
  '    Valor Unitário
  '    Valor Total
  '    Status
  '''  E.Data, E.Operação, E.Sequência, E.ChaveReferenciada, EP.Código, P.Nome, EP.Qtde, EP.Preço, EP.[Preço Final], E.Fornecedor, C.Nome, E.Digitador, U.Nome, E.Obs_Obs1, E.Observações
  If cboOrdenacao.Text = "Data Crescente" Then
      strSQL = strSQL & " ORDER BY E.Data "
  ElseIf cboOrdenacao.Text = "Data Decrescente" Then
      strSQL = strSQL & " ORDER BY E.Data DESC "
  ElseIf cboOrdenacao.Text = "Sequência" Then
      strSQL = strSQL & " ORDER BY E.Sequência "
  ElseIf cboOrdenacao.Text = "Cliente" Then
      strSQL = strSQL & " ORDER BY E.Fornecedor "
  ElseIf cboOrdenacao.Text = "Digitador" Then
      strSQL = strSQL & " ORDER BY E.Digitador "
  ElseIf cboOrdenacao.Text = "Quantidade" Then
      strSQL = strSQL & " ORDER BY EP.Qtde "
  ElseIf cboOrdenacao.Text = "Valor Unitário" Then
      strSQL = strSQL & " ORDER BY EP.Preço "
  ElseIf cboOrdenacao.Text = "Valor Total" Then
      strSQL = strSQL & " ORDER BY EP.[Preço Final] "
  ElseIf cboOrdenacao.Text = "Status" Then
      strSQL = strSQL & " ORDER BY EP.ConsignacaoFechada DESC "
  End If
  
  
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsEntradas
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        '''If IsNull(rsEntradas.Fields("Obs_Obs1").Value) Or rsEntradas.Fields("Obs_Obs1").Value = "" Then
        If IsNull(rsEntradas.Fields("ConsignacaoFechada").Value) Or rsEntradas.Fields("ConsignacaoFechada").Value = 0 Then
          strStatus = "Em Aberto"
        Else
          strStatus = "Recebido"
        End If
  
        rsTamanho.Index = "Código"
        rsTamanho.Seek "=", Mid(.Fields(4).Value, Len(.Fields(4).Value) - 5, 3)
        If Not rsTamanho.NoMatch Then
            sTamanho = rsTamanho.Fields("Nome").Value
        Else
            sTamanho = ""
        End If
        
        rsCor.Index = "Código"
        rsCor.Seek "=", Mid(.Fields(4).Value, Len(.Fields(4).Value) - 2, 3)
        If Not rsCor.NoMatch Then
            sCor = rsCor.Fields("Nome").Value
        Else
            sCor = ""
        End If
        sProdutoGradeAux = .Fields(5).Value & " " & sTamanho & " " & sCor
  
  
        'E.Data,            E.Operação,   E.Sequência,  E.ChaveReferenciada,
        'EP.Código,         P.Nome,       EP.Qtde,      EP.Preço,
        'EP.[Preço Final],  E.Fornecedor, C.Nome,       E.Digitador,
        'U.Nome,            E.Obs_Obs1,   E.Observações
        

        'Adiciona registro
        grdDevolucoes.AddItem "0" & vbTab & _
                              .Fields(0).Value & vbTab & _
                              strStatus & vbTab & _
                              .Fields(1).Value & vbTab & _
                              .Fields(4).Value & vbTab & _
                              sProdutoGradeAux & vbTab & _
                              .Fields(6).Value & vbTab & _
                              Format(.Fields(7).Value, FORMAT_VALUE) & vbTab & _
                              Format(.Fields(8).Value, FORMAT_VALUE) & vbTab & _
                              .Fields(9).Value & vbTab & _
                              .Fields(10).Value & vbTab & _
                              .Fields(11).Value & vbTab & _
                              .Fields(12).Value & vbTab & _
                              .Fields(2).Value & vbTab & _
                              .Fields(3).Value & vbTab & _
                              .Fields(14).Value
        
        sValorTotalGrade = Format(.Fields(8).Value, FORMAT_VALUE)
        dTotalGrade = dTotalGrade + CDbl(sValorTotalGrade)
        
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsEntradas = Nothing
  rsTamanho.Close
  rsCor.Close
  Set rsTamanho = Nothing
  Set rsCor = Nothing
  ' ---------------------------------------------------------------
   
  
  txt_total.Text = Format(dTotalGrade, FORMAT_VALUE)
    
  Exit Sub
Erro:
  MsgBox "Erro ao pesquisar registros de Devolução " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_Load()
  On Error GoTo Erro
  
  Data4.DatabaseName = gsQuickDBFileName
  Data_Ini.Text = Format(Data_Atual - 30, "dd/MM/yyyy")
  Data_Fim.Text = Format(Data_Atual, "dd/MM/yyyy")
  
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)

  Exit Sub
Erro:
  MsgBox "Erro ao carregar tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsCliFor.Close
  Set rsCliFor = Nothing
End Sub

Private Sub grdDevolucoes_Change()
  grdDevolucoes.Update
End Sub

Private Sub grdDevolucoes_RowLoaded(ByVal Bookmark As Variant)
  If IsEmpty(Bookmark) Then Exit Sub
End Sub
