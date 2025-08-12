VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B94AD83-0849-11D2-B2EB-444553540000}#1.0#0"; "TEXTGRID.OCX"
Begin VB.Form frmLayoutGen 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6315
   ClientLeft      =   465
   ClientTop       =   540
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LayoutGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11100
   Begin TEXTGRIDLib.TextGrid TextGrid1 
      Height          =   3015
      Left            =   60
      TabIndex        =   6
      Top             =   1605
      Width           =   10980
      _Version        =   65536
      _ExtentX        =   19368
      _ExtentY        =   5318
      _StockProps     =   109
      BackColor       =   -2147483637
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      MaxLine         =   180
      MaxCharacter    =   180
      LineHeight      =   16
      CharacterWidth  =   8
   End
   Begin VB.ComboBox Combo_Classe 
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
      ItemData        =   "LayoutGen.frx":4E95A
      Left            =   1890
      List            =   "LayoutGen.frx":4E96D
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton B_Grava 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Gravar"
      Height          =   400
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton B_Le 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ler"
      Height          =   400
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "Campos"
      Height          =   1605
      Left            =   1725
      TabIndex        =   34
      Top             =   4650
      Width           =   1575
      Begin VB.CommandButton B_Apaga 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Apagar atual..."
         Height          =   400
         Left            =   120
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   690
         Width           =   1335
      End
      Begin VB.CommandButton B_Limpa 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Limpar &Tudo..."
         Height          =   400
         Left            =   120
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   1395
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Duplo-clique para transferir seleção"
      Top             =   90
      Width           =   5880
   End
   Begin VB.Frame Frame4 
      Caption         =   "Linhas "
      Height          =   1605
      Left            =   60
      TabIndex        =   32
      Top             =   4650
      Width           =   1575
      Begin VB.CommandButton B_Insere_Linha 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Inserir"
         Height          =   400
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton B_Remove 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Remover"
         Height          =   400
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   690
         Width           =   1335
      End
      Begin VB.CommandButton B_Duplica 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Duplicar"
         Height          =   400
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1155
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tamanho"
      Height          =   1605
      Left            =   3390
      TabIndex        =   29
      Top             =   4650
      Width           =   2310
      Begin VB.TextBox Colunas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   870
         MaxLength       =   3
         TabIndex        =   13
         Top             =   645
         Width           =   540
      End
      Begin VB.TextBox Linhas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   870
         MaxLength       =   3
         TabIndex        =   12
         Top             =   270
         Width           =   540
      End
      Begin VB.CommandButton B_Redimensiona 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Redimensionar"
         Height          =   400
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1125
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Colunas"
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Linhas"
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressão"
      Height          =   1605
      Left            =   5790
      TabIndex        =   28
      Top             =   4650
      Width           =   3600
      Begin VB.TextBox txtNumPol 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Text            =   "4"
         Top             =   1245
         Width           =   420
      End
      Begin VB.OptionButton optComprPag 
         Appearance      =   0  'Flat
         Caption         =   "&Polegadas:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1290
         Width           =   1140
      End
      Begin VB.OptionButton optComprPag 
         Appearance      =   0  'Flat
         Caption         =   "&Qtde linhas úteis no lay-out"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1035
         Width           =   2310
      End
      Begin VB.OptionButton optComprPag 
         Appearance      =   0  'Flat
         Caption         =   "&Normal via driver"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   795
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CheckBox O_Oitavo 
         Appearance      =   0  'Flat
         Caption         =   "Impressão em 1/&8 """
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1440
         TabIndex        =   16
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox O_Comprimida 
         Appearance      =   0  'Flat
         Caption         =   "&Comprimida"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Comprimento página:"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   555
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zoom"
      Height          =   1605
      Left            =   9480
      TabIndex        =   27
      Top             =   4650
      Width           =   1575
      Begin VB.OptionButton O_Muito 
         Appearance      =   0  'Flat
         Caption         =   "&Muito Pequeno"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   900
         Width           =   1380
      End
      Begin VB.OptionButton O_Pequeno 
         Appearance      =   0  'Flat
         Caption         =   "&Pequeno"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   585
         Width           =   1065
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         Caption         =   "&Normal"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.CommandButton B_OK_Fixo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Transferir texto"
      Height          =   400
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Tamanho 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   10410
      MaxLength       =   3
      TabIndex        =   25
      Top             =   1230
      Width           =   615
   End
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      MaxLength       =   132
      TabIndex        =   3
      Top             =   900
      Width           =   5025
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   8640
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblClasse 
      Caption         =   "Campos disponíveis"
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
      Left            =   120
      TabIndex        =   35
      Top             =   60
      Width           =   1590
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   9120
      Picture         =   "LayoutGen.frx":4E9A9
      Top             =   30
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Texto Fixo"
      Height          =   195
      Left            =   5985
      TabIndex        =   26
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Campo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      ToolTipText     =   "Área de transferência"
      Top             =   1230
      Width           =   4365
   End
End
Attribute VB_Name = "frmLayoutGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Tabela
  Nome As String
  Tamanho As Integer
  Classe As String
End Type

Private Campos(200, 3) As Tabela

Private Type Tabela2
  Campo As String
  Coluna As Integer
  Tamanho As Integer
  Tipo As String
End Type

Private Tipo As String
Private Máx As Long
Private Tam_Lista As Long

Private Const Laranja = &H80FF&
Private Const Amarelo = &HFFFF&
Private Const Verde = &HFF00&
Private Const Azul = &HFFFF00
Private Const Branco = &HFFFFFF
Private Const Rosa = &HC0C0FF

Private Cor As String

Public gnTypeDoc As Integer
Public gsTypeDoc As String

Private gsFileExtension As String
Private gsTitleCaption As String
Private gsFileHeader As String
Private gsDialogTextFile As String

Private gbDirty As Boolean

'23/10/2009 - mpdea
'Adequado array de campos com variável de controle de registro (Substituído número fixo por variável A e incremento)
Private Sub Inicia_Lista()
  Dim A As Integer
  
  Select Case gnTypeDoc
  
    Case 0  'BOLETOS
    
      gsFileExtension = "*.CBB"
      gsTitleCaption = "Layout de Boletos Bancários"
      gsFileHeader = "*** Configurações Boleto:"
      gsDialogTextFile = "Configuração de Boletos Bancários " & "(" & gsFileExtension & ") | "
      
      Campos(A, gnTypeDoc).Nome = "Código Cliente"
      Campos(A, gnTypeDoc).Tamanho = 8
      A = A + 1

      Campos(A, gnTypeDoc).Nome = "Nome Cliente"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Fantasia"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Endereço"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      '23/10/2009 - mpdea
      'Número do endereço
      Campos(A, gnTypeDoc).Nome = "Número Endereço"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Complemento"
      Campos(A, gnTypeDoc).Tamanho = 15
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Bairro"
      Campos(A, gnTypeDoc).Tamanho = 20
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "CEP"
      Campos(A, gnTypeDoc).Tamanho = 9
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Cidade"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Estado"
      Campos(A, gnTypeDoc).Tamanho = 2
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "CGC"
      Campos(A, gnTypeDoc).Tamanho = 18
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Inscrição Estadual"
      Campos(A, gnTypeDoc).Tamanho = 18
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Seqüência"
      Campos(A, gnTypeDoc).Tamanho = 8
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Número Nota"
      Campos(A, gnTypeDoc).Tamanho = 8
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data Emissão Conta"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data Saída"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Fatura Receber"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Descrição"
      Campos(A, gnTypeDoc).Tamanho = 40
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Desconto"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Acréscimo"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Vencimento"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Mensagem Cliente"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Obs1"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Obs2"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Obs3"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_60"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso61_120"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso121_180"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_45"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso46_90"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso91_135"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso136_180"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_30"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso31_60"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso61_90"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso91_120"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso121_150"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso151_180"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "FINAL DE CONFIGURAÇÃO"
      Campos(A, gnTypeDoc).Tamanho = 21
      A = A + 1
      
      '16/08/2002 - mpdea
      'Incluído o campo de personalização LINHA EM NEGRITO
      Campos(A, gnTypeDoc).Nome = "LINHA_EM_NEGRITO"
      Campos(A, gnTypeDoc).Tamanho = 16
      A = A + 1
      
      '15/01/2004 - Daniel
      'Incluído o campo Valor Recebido proveniente da
      'tabela de [Contas a Receber]
      Campos(A, gnTypeDoc).Nome = "Valor Recebido"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      '15/01/2004 - Daniel
      'Incluído o campo valor total com finalidade
      'de somatório de ticket - Case: F. Linhares
      Campos(A, gnTypeDoc).Nome = "Valor Total do Ticket"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
  
  
    Case 1  'CARNÊ
      
      gsFileExtension = "*.CCA"
      gsTitleCaption = "Layout de Carnês"
      gsFileHeader = "*** Configurações Carnê :"
      gsDialogTextFile = "Configuração de Carnês " & "(" & gsFileExtension & ") | "
      
      Campos(A, gnTypeDoc).Nome = "Código Cliente"
      Campos(A, gnTypeDoc).Tamanho = 8
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Cliente"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Fantasia"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Endereço"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      '23/10/2009 - mpdea
      'Número do endereço
      Campos(A, gnTypeDoc).Nome = "Número Endereço"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Complemento"
      Campos(A, gnTypeDoc).Tamanho = 15
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Bairro"
      Campos(A, gnTypeDoc).Tamanho = 20
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "CEP"
      Campos(A, gnTypeDoc).Tamanho = 9
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Cidade"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Estado"
      Campos(A, gnTypeDoc).Tamanho = 2
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "CGC"
      Campos(A, gnTypeDoc).Tamanho = 18
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Inscrição Estadual"
      Campos(A, gnTypeDoc).Tamanho = 18
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Seqüência"
      Campos(A, gnTypeDoc).Tamanho = 8
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Número Nota"
      Campos(A, gnTypeDoc).Tamanho = 8
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data Emissão Conta"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data Saída"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Fatura Receber"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Descrição"
      Campos(A, gnTypeDoc).Tamanho = 40
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Desconto"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Acréscimo"
      Campos(A, gnTypeDoc).Tamanho = 12
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Vencimento"
      Campos(A, gnTypeDoc).Tamanho = 10
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Mensagem Cliente"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Obs1"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Obs2"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Obs3"
      Campos(A, gnTypeDoc).Tamanho = 50
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_60"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso61_120"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso121_180"
      Campos(A, gnTypeDoc).Tamanho = 60
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_45"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso46_90"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso91_135"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso136_180"
      Campos(A, gnTypeDoc).Tamanho = 45
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_30"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso31_60"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso61_90"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso91_120"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso121_150"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso151_180"
      Campos(A, gnTypeDoc).Tamanho = 30
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "FINAL DE CONFIGURAÇÃO"
      Campos(A, gnTypeDoc).Tamanho = 21
      A = A + 1
      
      '16/08/2002 - mpdea
      'Incluído o campo de personalização LINHA EM NEGRITO
      Campos(A, gnTypeDoc).Nome = "LINHA_EM_NEGRITO"
      Campos(A, gnTypeDoc).Tamanho = 16
      A = A + 1
    
    
    Case 2    'TICKET OU NOTA FISCAL

      If gsTypeDoc = "NOTA" Then
        gsFileExtension = "*.CNF"
        gsFileHeader = "*** Configurações Nota:  "
        gsDialogTextFile = "Configuração de Notas Fiscais " & "(" & gsFileExtension & ") | "
        gsTitleCaption = "Layout de Nota Fiscal"
      Else
        gsFileExtension = "*.CTI"
        gsFileHeader = "*** Configurações Ticket:"
        gsDialogTextFile = "Configuração de Tickets " & "(" & gsFileExtension & ") | "
        gsTitleCaption = "Layout de Ticket"
      End If

      Campos(A, gnTypeDoc).Nome = "Número Nota"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
    
      Campos(A, gnTypeDoc).Nome = "Código Filial"
      Campos(A, gnTypeDoc).Tamanho = 2
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Filial"
      Campos(A, gnTypeDoc).Tamanho = 25
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '31/05/2007 - Anderson
      'Incluído o Prometido Para
      Campos(A, gnTypeDoc).Nome = "Prometido para"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '31/05/2007 - Anderson
      'Incluído o Prometido Para
      Campos(A, gnTypeDoc).Nome = "Orçamento Aprovado por"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data"
      Campos(A, gnTypeDoc).Tamanho = "10"
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data Saída"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Hora Saída"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Operação"
      Campos(A, gnTypeDoc).Tamanho = 4
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Operação"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Fiscal"
      Campos(A, gnTypeDoc).Tamanho = 14
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      If gsTypeDoc = "NOTA" Then
        '06/05/2007 - Anderson
        'Implementação da impressão do código CFOP por serviço
        Campos(A, gnTypeDoc).Nome = "Código Fiscal Completo (Operação + Itens)"
        Campos(A, gnTypeDoc).Tamanho = 24
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Código Fiscal Item 1"
        Campos(A, gnTypeDoc).Tamanho = 4
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Código Fiscal Item 2"
        Campos(A, gnTypeDoc).Tamanho = 4
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Código Fiscal Item 3"
        Campos(A, gnTypeDoc).Tamanho = 4
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Código Fiscal Item 4"
        Campos(A, gnTypeDoc).Tamanho = 4
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Código Fiscal Item 5"
        Campos(A, gnTypeDoc).Tamanho = 4
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        '24/04/2008 - mpdea
        'Descrição e total por CFOP relacionado a movimentação
        Campos(A, gnTypeDoc).Nome = "Nome Operação - Código Fiscal 1"
        Campos(A, gnTypeDoc).Tamanho = 30
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Nome Operação - Código Fiscal 2"
        Campos(A, gnTypeDoc).Tamanho = 30
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Nome Operação - Código Fiscal 3"
        Campos(A, gnTypeDoc).Tamanho = 30
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Nome Operação - Código Fiscal 4"
        Campos(A, gnTypeDoc).Tamanho = 30
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Nome Operação - Código Fiscal 5"
        Campos(A, gnTypeDoc).Tamanho = 30
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
        
        Campos(A, gnTypeDoc).Nome = "Valor Total - Código Fiscal 1"
        Campos(A, gnTypeDoc).Tamanho = 12
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
      
        Campos(A, gnTypeDoc).Nome = "Valor Total - Código Fiscal 2"
        Campos(A, gnTypeDoc).Tamanho = 12
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
      
        Campos(A, gnTypeDoc).Nome = "Valor Total - Código Fiscal 3"
        Campos(A, gnTypeDoc).Tamanho = 12
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
      
        Campos(A, gnTypeDoc).Nome = "Valor Total - Código Fiscal 4"
        Campos(A, gnTypeDoc).Tamanho = 12
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
      
        Campos(A, gnTypeDoc).Nome = "Valor Total - Código Fiscal 5"
        Campos(A, gnTypeDoc).Tamanho = 12
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
      End If
          
      Campos(A, gnTypeDoc).Nome = "Seqüência"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
          
      Campos(A, gnTypeDoc).Nome = "Código Vendedor"
      Campos(A, gnTypeDoc).Tamanho = 4
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Vendedor"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '20/11/2006 - Anderson
      'Incluído o campo Apelido
      Campos(A, gnTypeDoc).Nome = "Apelido"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '--------------------------------------------------------------------------
      '07/08/2003 - mpdea
      'Incluído Código e Nome do Técnico
      Campos(A, gnTypeDoc).Nome = "Código Técnico"
      Campos(A, gnTypeDoc).Tamanho = 4
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Técnico"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      '--------------------------------------------------------------------------
      
      
      Campos(A, gnTypeDoc).Nome = "Código Cliente"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Cliente"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Fantasia"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Endereço"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
            
      '23/10/2009 - mpdea
      'Número do endereço
      Campos(A, gnTypeDoc).Nome = "Número Endereço"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Complemento"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Bairro"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
          
      Campos(A, gnTypeDoc).Nome = "CEP"
      Campos(A, gnTypeDoc).Tamanho = 9
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
  
      Campos(A, gnTypeDoc).Nome = "Cidade"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Estado"
      Campos(A, gnTypeDoc).Tamanho = 2
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "FONE1"
      Campos(A, gnTypeDoc).Tamanho = 25
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "FONE2"
      Campos(A, gnTypeDoc).Tamanho = 25
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "CGC"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Inscrição Estadual"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Caixa"
      Campos(A, gnTypeDoc).Tamanho = 2
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
  
      Campos(A, gnTypeDoc).Nome = "Nome Caixa"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Tabela Preço"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Referência Interna"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Observações"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs1"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs2"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs3"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs4"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs5"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs6"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs7"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Obs8"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      
      '------------------------------------------------------------------------
      '31/01/2006 - mpdea
      'Adicionado novos campos para impressão de Mensagens para Nota Fiscal
      If gsTypeDoc = "NOTA" Then
        Campos(A, gnTypeDoc).Nome = "MensagemNotaFiscal1"
        Campos(A, gnTypeDoc).Tamanho = 80
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
         
        Campos(A, gnTypeDoc).Nome = "MensagemNotaFiscal2"
        Campos(A, gnTypeDoc).Tamanho = 80
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
         
        Campos(A, gnTypeDoc).Nome = "MensagemNotaFiscal3"
        Campos(A, gnTypeDoc).Tamanho = 80
        Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
        A = A + 1
      End If
      '------------------------------------------------------------------------
      
       
      Campos(A, gnTypeDoc).Nome = "Nome Transportadora"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
        
      Campos(A, gnTypeDoc).Nome = "CNPJ Transportadora"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
        
      Campos(A, gnTypeDoc).Nome = "IE Transportadora"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
        
      Campos(A, gnTypeDoc).Nome = "Ender Transportadora"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
        
      Campos(A, gnTypeDoc).Nome = "Municipio Transportadora"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
        
      Campos(A, gnTypeDoc).Nome = "UF Transportadora"
      Campos(A, gnTypeDoc).Tamanho = 2
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
        
      Campos(A, gnTypeDoc).Nome = "Placa"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "UF Placa"
      Campos(A, gnTypeDoc).Tamanho = 2
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Qtde Trans"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Espécie Trans"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Marca Trans"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Peso Bruto"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Peso Líquido"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
       
      Campos(A, gnTypeDoc).Nome = "Frete Conta"
      Campos(A, gnTypeDoc).Tamanho = 1
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      
      '15/08/2002 - mpdea
      'Incluído o campo de informações sobre o orçamento (número do orçamento e terminal)
      Campos(A, gnTypeDoc).Nome = "Número do Orçamento"
      Campos(A, gnTypeDoc).Tamanho = 24
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      
      '--------------------------------------------
      '08/01/2004 - Daniel
      'Inclusão dos campos Valor Recebido e Troco
      'da tabela de Saídas
      Campos(A, gnTypeDoc).Nome = "Valor Recebido Venda"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Troco"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      '--------------------------------------------
      
      '--------------------------------------------
      '09/01/2004 - Daniel
      'Inclusão do campo SOMA DA QTDE DE ITENS
      'proveniente de variável
      Campos(A, gnTypeDoc).Nome = "Soma da Qtde de Itens"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      '--------------------------------------------
      
      '--------------------------------------------
      '30/01/2004 - Daniel
      'Inclusão dos campos Percentual CSLL,
      'Percentual COFINS, Percentual PIS,
      'Percentual IRFF da tabela Parâmetros Filial
      'e Totais: Total CSLL, Total COFINS
      'Total PIS, Total IRRF
      Campos(A, gnTypeDoc).Nome = "Percentual CSLL"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Percentual COFINS"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Percentual PIS"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Percentual IRRF"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '--------------------------------------------
      Campos(A, gnTypeDoc).Nome = "Total CSLL"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Total COFINS"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Total PIS"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Total IRRF"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      '--------------------------------------------
      
      '----------------------------------------------------------------------
      '13/04/2004 - Daniel
      'Inclusão dos Campos:
      'Saídas.[Num Autorização], Saídas.MesX, Cli_For.[Endereço Cob],
      'Cli_For.[Complemento Cob], Cli_For.[Bairro Cob], Cli_For.[Cidade Cob],
      'Cli_For.[Estado Cob] e Cli_For.[CEP Cob]
      Campos(A, gnTypeDoc).Nome = "Número da Autorização"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Mês X"
      Campos(A, gnTypeDoc).Tamanho = 1
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Endereço Cob"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Complemento Cob"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Bairro Cob"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Cidade Cob"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Estado Cob"
      Campos(A, gnTypeDoc).Tamanho = 2
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "CEP Cob"
      Campos(A, gnTypeDoc).Tamanho = 9
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '-----------------------------------------------------------------------
      
      '-----------------------------------------------------------------------
      '06/05/2004 - Daniel
      'Adição do campo ObsIsentoIPI da tabela Cli_For
      Campos(A, gnTypeDoc).Nome = "Obs Isenção IPI"
      Campos(A, gnTypeDoc).Tamanho = 100
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '17/05/2004 - Daniel
      'Adição do campo ObsDiferimento da tabela Diferimento
      Campos(A, gnTypeDoc).Nome = "Obs Diferimento"
      Campos(A, gnTypeDoc).Tamanho = 70
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      '-----------------------------------------------------------------------
      
      '-----------------------------------------------------------------------
      '27/04/2005 - Daniel
      'Adição do campo Seguro da tabela de Saídas
      Campos(A, gnTypeDoc).Nome = "Seguro"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      '-----------------------------------------------------------------------
      
      '19/12/2007 - Anderson
      'Implementação do NSU
      Campos(A, gnTypeDoc).Nome = "NSU"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '19/12/2007 - Anderson
      'Implementação do NSU
      Campos(A, gnTypeDoc).Nome = "NSU (Data Geração)"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '19/12/2007 - Anderson
      'Implementação do NSU
      Campos(A, gnTypeDoc).Nome = "NSU (Hora Geração)"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Rem PRODUTOS
      Rem PRODUTOS
      Rem PRODUTOS
      
      Campos(A, gnTypeDoc).Nome = "Código Produto"
      Campos(A, gnTypeDoc).Tamanho = "20"
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Produto Fornecedor"
      Campos(A, gnTypeDoc).Tamanho = 20
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Classificação Fiscal"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      '11/11/2004 - Daniel
      'Adicionado a Descrição da Classificação Fiscal
      Campos(A, gnTypeDoc).Nome = "Descrição da Class. Fiscal"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
            
      '29/04/2008 - mpdea
      'CFOP do produto
      Campos(A, gnTypeDoc).Nome = "CFOP do Produto"
      Campos(A, gnTypeDoc).Tamanho = 14
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Situação Tributária"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
            
      '05/05/2011 - mpdea
      'NBM/NCM do produto
      Campos(A, gnTypeDoc).Nome = "Código NBM/NCM"
      Campos(A, gnTypeDoc).Tamanho = 8
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      '27/04/2005 - Daniel
      'Adicionado o campo Fabricante da tabela Produtos
      Campos(A, gnTypeDoc).Nome = "Fabricante"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
    '---------------------------------------------------------------------------------
      '27/11/2002 - mpdea
      'Alterado o tamanho do campo para nome do produto (50 -> 80)
      
      Campos(A, gnTypeDoc).Nome = "Nome Produto"
      Campos(A, gnTypeDoc).Tamanho = "80"
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      '---------------------------------------------------------------------------------
      '04/09/2002 - mpdea
      'Incluído os campos para impressão específica do nome do produto como
      'está no campo Nome do cadastro ou o campo Nome para nota (Fixo)
      
      Campos(A, gnTypeDoc).Nome = "Nome Produto (Cadastro)"
      Campos(A, gnTypeDoc).Tamanho = "80"
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Produto (Nota)"
      Campos(A, gnTypeDoc).Tamanho = "80"
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
      
      Campos(A, gnTypeDoc).Nome = "Qtde"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Unidade Venda"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Preço Unitário"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Preço Produto Total"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Perc Desconto Produto"
      Campos(A, gnTypeDoc).Tamanho = 6
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Perc ICM Produto"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor ICM Produto"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Perc IPI Produto"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor IPI Produto"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Preço Final Produto"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Pesquisa 1"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Pesquisa 2"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Código Pesquisa 3"
      Campos(A, gnTypeDoc).Tamanho = 5
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Pesquisa 1"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Pesquisa 2"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Pesquisa 3"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Qtde Itens Produto"
      Campos(A, gnTypeDoc).Tamanho = 6
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Local"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Descrição Adicional"
      Campos(A, gnTypeDoc).Tamanho = 50
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Cor"
      Campos(A, gnTypeDoc).Tamanho = 3
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Cor"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Tamanho"
      Campos(A, gnTypeDoc).Tamanho = 3
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Nome Tamanho"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      '27/09/2004 - mpdea
      'Incluído a Volumagem por Quantidade
      Campos(A, gnTypeDoc).Nome = "Volumagem por Qtde"
      Campos(A, gnTypeDoc).Tamanho = 9
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      '29/11/2004 - Daniel
      'Incluído os campos Lote e Data de Validade
      Campos(A, gnTypeDoc).Nome = "Lote"
      Campos(A, gnTypeDoc).Tamanho = 15
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Data de Validade"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "PRODUTO"
      A = A + 1
      
            
      Rem SERVIÇOS
      Rem SERVIÇOS
      Rem SERVIÇOS
      Campos(A, gnTypeDoc).Nome = "Código Serviço"
      Campos(A, gnTypeDoc).Tamanho = 4
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
                  
      '29/04/2008 - mpdea
      'CFOP do serviço
      Campos(A, gnTypeDoc).Nome = "CFOP do Serviço"
      Campos(A, gnTypeDoc).Tamanho = 14
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1

      '12/02/2003 - mpdea
      'Corrigido a limitação de caracteres (50 -> 60)
      Campos(A, gnTypeDoc).Nome = "Nome Serviço"
      Campos(A, gnTypeDoc).Tamanho = 60
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Qtde Serviço"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Preço Unitário Serviço"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Preço Total Serviço"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Serviço Concluído"
      Campos(A, gnTypeDoc).Tamanho = 3
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Qtde Itens Serviço"
      Campos(A, gnTypeDoc).Tamanho = 6
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor Total Serviço"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      '27/07/2005 - Daniel
      'Campo: CST (Código de Situação Tributária)
      'Finalidade: Atender a realidade da empresa W.V. Hidroanálise Ltda (J.R. Hidroquímica)
      Campos(A, gnTypeDoc).Nome = "CST"
      Campos(A, gnTypeDoc).Tamanho = 1
      Campos(A, gnTypeDoc).Classe = "SERVIÇO"
      A = A + 1
      
      
      Rem RODAPÉ
      Rem RODAPÉ
      Rem RODAPÉ
      Campos(A, gnTypeDoc).Nome = "Valor Total Serviços"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor ISS"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Base Cálculo ICM"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor ICM"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Base Cálculo ICM Subs"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor ICM Subs"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor Total dos Produtos"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '19/08/2003 - mpdea
      'Incluído o campo para exibição do Total de Produtos com Desconto no SubTotal
      Campos(A, gnTypeDoc).Nome = "Valor Total dos Produtos com Desconto no SubTotal"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '26/08/2003 - mpdea
      'Incluído o campo para exibição do Total de Produtos menos Total de Descontos
      Campos(A, gnTypeDoc).Nome = "Valor Total dos Produtos menos Total de Descontos"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      '06/09/2002 - mpdea
      'Incluído o campo para exibição do Desconto no SubTotal
      Campos(A, gnTypeDoc).Nome = "Desconto no SubTotal"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor Total de Descontos"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
                    
      Campos(A, gnTypeDoc).Nome = "Valor Frete"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor Total IPI"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor Total da Nota"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "CABEÇALHO"
      A = A + 1
      
        
      Rem OUTROS
      Rem OUTROS
      Rem OUTROS
      
      Campos(A, gnTypeDoc).Nome = "FINAL DE CONFIGURAÇÃO"
      Campos(A, gnTypeDoc).Tamanho = 21
      Campos(A, gnTypeDoc).Classe = "OUTROS"
      A = A + 1
     
      Campos(A, gnTypeDoc).Nome = "PROXIMO_PRODUTO"
      Campos(A, gnTypeDoc).Tamanho = 1
      Campos(A, gnTypeDoc).Classe = "OUTROS"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "PROXIMO_SERVIÇO"
      Campos(A, gnTypeDoc).Tamanho = 1
      Campos(A, gnTypeDoc).Classe = "OUTROS"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "PROXIMA_FATURA"
      Campos(A, gnTypeDoc).Tamanho = 1
      Campos(A, gnTypeDoc).Classe = "OUTROS"
      A = A + 1
      
      
      '16/08/2002 - mpdea
      'Incluído o campo de personalização LINHA EM NEGRITO
      Campos(A, gnTypeDoc).Nome = "LINHA_EM_NEGRITO"
      Campos(A, gnTypeDoc).Tamanho = 16
      Campos(A, gnTypeDoc).Classe = "OUTROS"
      A = A + 1
      
           
     
      Rem Fatura
      Campos(A, gnTypeDoc).Nome = "Data Fatura"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Fatura"
      Campos(A, gnTypeDoc).Tamanho = 10
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Valor Fatura"
      Campos(A, gnTypeDoc).Tamanho = 12
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_60"
      Campos(A, gnTypeDoc).Tamanho = 60
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso61_120"
      Campos(A, gnTypeDoc).Tamanho = 60
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso121_180"
      Campos(A, gnTypeDoc).Tamanho = 60
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_45"
      Campos(A, gnTypeDoc).Tamanho = 45
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso46_90"
      Campos(A, gnTypeDoc).Tamanho = 45
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso91_135"
      Campos(A, gnTypeDoc).Tamanho = 45
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso136_180"
      Campos(A, gnTypeDoc).Tamanho = 45
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso1_30"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso31_60"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso61_90"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso91_120"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso121_150"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
      
      Campos(A, gnTypeDoc).Nome = "Extenso151_180"
      Campos(A, gnTypeDoc).Tamanho = 30
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
     
      Campos(A, gnTypeDoc).Nome = "Qtde Parcelas Fatura"
      Campos(A, gnTypeDoc).Tamanho = 6
      Campos(A, gnTypeDoc).Classe = "FATURA"
      A = A + 1
    
    '  If Tipo_Tela.Caption = "TICKET" Then
      Campos(A, gnTypeDoc).Nome = "RESUMO DO PAGAMENTO"
      Campos(A, gnTypeDoc).Tamanho = 19
      Campos(A, gnTypeDoc).Classe = "OUTROS"
      A = A + 1
     ' End If
            
      
  End Select
  
  Tam_Lista = UBound(Campos, 1) + 1
  
End Sub

Private Sub Mostra_Campo_Atual(i As Long)
 Dim sTexto As String
 Dim nTamanho As Long
 
  sTexto = TextGrid1.ItemText(i)
  sTexto = sTexto & "  - início "
  sTexto = sTexto & str(TextGrid1.ItemStartCharacter(i))
  sTexto = sTexto & "  - tamanho "
  nTamanho = TextGrid1.ItemEndCharacter(i) - TextGrid1.ItemStartCharacter(i)
  sTexto = sTexto & str(nTamanho)
  sTexto = sTexto & "   Max " + str(TextGrid1.ItemMaxLen(i))
 
  Call StatusMsg(sTexto)
  
End Sub

Private Function Pega_Tipo_Campo(ByVal gnTypeDoc As Integer, Texto As String) As String

  Dim i As Integer
  
  For i = 0 To Tam_Lista - 1
   
    If Campos(i, gnTypeDoc).Nome = Texto Then
       Pega_Tipo_Campo = Campos(i, gnTypeDoc).Classe
       Exit Function
    End If
  
  Next i
  
  Pega_Tipo_Campo = "OUTROS"

End Function

Private Function Pega_Tamanho(Nome_Campo As String) As Long
 
  Dim i As Long
  
  For i = 0 To Tam_Lista
   If Campos(i, gnTypeDoc).Nome = Nome_Campo Then
     Pega_Tamanho = Campos(i, gnTypeDoc).Tamanho
     Exit Function
   End If
  Next i
  
  Pega_Tamanho = 0
  
End Function

Private Sub Remonta_Lista()
  Dim i As Integer
  
  If gnTypeDoc = 2 Then
    If Combo_Classe.Text = "Cabeçalho / Rodapé" Then Texto = "CABEÇALHO"
    If Combo_Classe.Text = "Produtos" Then Texto = "PRODUTO"
    If Combo_Classe.Text = "Outros" Then Texto = "OUTROS"
    If Combo_Classe.Text = "Serviços" Then Texto = "SERVIÇO"
    If Combo_Classe.Text = "Fatura" Then Texto = "FATURA"
  End If
   
  Lista.Clear
  
  If gnTypeDoc <> 2 Then
    For i = 0 To Tam_Lista - 1
      If Len(Campos(i, gnTypeDoc).Nome & "") > 0 Then
        Lista.AddItem Campos(i, gnTypeDoc).Nome
      End If
    Next i
  Else
    For i = 0 To Tam_Lista - 1
     If Campos(i, gnTypeDoc).Classe = Texto Then
       Lista.AddItem Campos(i, gnTypeDoc).Nome
     End If
    Next i
  End If
  
End Sub

Private Function Retorna_Linha(Linha As Integer) As String
  Dim Campos As Long
  Dim Campos_Linha As Integer
  Dim i As Integer
  Dim ws As String
  Dim Menor As Integer
  Dim Menor1 As Integer
  Dim Fim As Integer
  Dim Coluna_Atual As Integer
  Dim Brancos As Integer
  Dim Texto As String
  Dim Tabe(100) As Tabela2
 
  ws = Chr(160)

  Campos = TextGrid1.ItemCount
  For i = 0 To Campos - 1
    If TextGrid1.ItemLine(i) = Linha Then
      Campos_Linha = Campos_Linha + 1
      Tabe(Campos_Linha).Campo = TextGrid1.ItemText(i)
      Tabe(Campos_Linha).Coluna = TextGrid1.ItemStartCharacter(i)
      Tabe(Campos_Linha).Tamanho = (TextGrid1.ItemEndCharacter(i) - TextGrid1.ItemStartCharacter(i))
      If Right(TextGrid1.ItemText(i), 1) = ws Then
        Tabe(Campos_Linha).Tipo = "CAMPO"
        Menor = Len(Tabe(Campos_Linha).Campo)
        Texto = Left(Tabe(Campos_Linha).Campo, (Menor - 1))
        Tabe(Campos_Linha).Campo = Texto
      Else
        Tabe(Campos_Linha).Tipo = "TEXTO"
      End If
    End If
  Next i
 
  If Campos_Linha = 0 Then
    Retorna_Linha = ""
    Exit Function
  End If

  Fim = False
  Coluna_Atual = 1
  Texto = ""
  Do
    Menor1 = 0
    Menor = 999
    For i = 1 To Campos_Linha
      If Tabe(i).Coluna < Menor And Tabe(i).Coluna > 0 Then
        Menor = Tabe(i).Coluna
        Menor1 = i
      End If
    Next i
  
    If Menor1 = 0 Then Fim = True
    If Fim = False Then
      Brancos = Tabe(Menor1).Coluna - Coluna_Atual
      If Brancos > 0 Then
        Texto = Texto + "{"
        For i = 1 To Brancos
         Texto = Texto + " "
        Next i
        Texto = Texto + "}"
        Coluna_Atual = Coluna_Atual + Brancos
      End If
      If Tabe(Menor1).Tipo = "TEXTO" Then
        Texto = Texto + "{" + Tabe(Menor1).Campo + "}"
      Else
        Texto = Texto + "[" + Tabe(Menor1).Campo + "," + Trim(str(Tabe(Menor1).Tamanho)) + "]"
      End If
      Coluna_Atual = Coluna_Atual + Tabe(Menor1).Tamanho
      Tabe(Menor1).Coluna = -1 'elimina este do loop
   End If
  Loop Until Fim = True

  Retorna_Linha = Texto
End Function

Private Sub B_Apaga_Click()
  
  If TextGrid1.ActiveItem < 0 Then Exit Sub
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja apagar o campo """ & TextGrid1.ItemText(TextGrid1.ActiveItem) & """?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Exit Sub
  End If
  
  TextGrid1.RemoveItem (TextGrid1.ActiveItem)
  gbDirty = True
  
End Sub

Private Sub B_Duplica_Click()
  Dim Resp As String
  Dim Erro As Integer
  
  Call StatusMsg("")

  Resp = InputBox("Deseja duplicar qual linha ?", Me.Caption)
  
  Erro = False
  If IsNull(Resp) Then Erro = True
  If Erro = False Then If Resp = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Resp) Then Erro = True
  If Erro = False Then If Val(Resp) < 1 Then Erro = True
  If Erro = False Then If Val(Resp) > TextGrid1.MaxLine Then Erro = True
  
  If Erro = True Then
    DisplayMsg "Linha incorreta."
    Exit Sub
  End If
  
  TextGrid1.DuplicateLine Val(Resp), (Val(Resp) + 1)
  gbDirty = True

End Sub

Private Sub B_Grava_Click()
  Dim Linhas As Integer
  Dim i As Integer
  Dim Texto As String
  Dim nFileNum As Integer
  Dim nLastLine As Integer
  
  nLastLine = Verifica_Final
  If nLastLine = -1 Then
    DisplayMsg "Não é possível gravar sem o campo FINAL DE CONFIGURAÇÃO"
    Exit Sub
  End If
  
  Call StatusMsg("")
  Dialog1.CancelError = True
  Dialog1.InitDir = gsConfigPath
  Dialog1.Filter = gsDialogTextFile + gsFileExtension
  On Error GoTo Cancelou
  Dialog1.Action = 2
  If Dialog1.FileName = gsFileExtension Then Exit Sub
  On Error GoTo 0
  
  
  On Error GoTo Erro_Abrir
  nFileNum = FreeFile
  Open Dialog1.FileName For Output As #nFileNum
  On Error GoTo 0
  
  Linhas = TextGrid1.MaxLine
  Texto = ""
  
  
  Texto = gsFileHeader & " COMPRIMIDO : "
  
  If O_Comprimida.Value = 1 Then
    Texto = Texto + "SIM"
  Else
    Texto = Texto + "NÃO"
  End If
  
  Texto = Texto + "   OITAVO : "
  If O_Oitavo.Value = 1 Then
    Texto = Texto + "SIM"
  Else
    Texto = Texto + "NÃO"
  End If
  
  Texto = Texto + "   COMPRIMENTO : "
  If optComprPag(0).Value = True Then
    Texto = Texto + "NÃO"
  Else
    If optComprPag(1).Value = True Then
      Texto = Texto + "LIN"
    Else
      Texto = Texto + Trim(txtNumPol.Text)
    End If
  End If
  
  Rem Linha final fica
  Rem *** Configurações: COMPRIMIDO : XXX   OITAVO : XXX   COMPRIMENTO : XXX
  
  Write #nFileNum, Texto
  
  For i = 1 To Linhas
    Texto = Retorna_Linha(i)
    
    
    If InStr(Texto, "FINAL DE CONFIGURAÇÃO") Then
      Write #nFileNum, "*** Fim de arquivo ***"
      Close #nFileNum
      Me.Caption = gsTitleCaption & " - [" & Dialog1.FileName & "]"
      gbDirty = False
      Exit Sub
    End If
      
      
    If Texto = "" Then
      Write #nFileNum, "[LINHA_BRANCO,1]"
    Else
      Write #nFileNum, Texto
    End If
    
    
  Next i
  
  Close #nFileNum
  
  Me.Caption = gsTitleCaption & " - [" & Dialog1.FileName & "]"
  gbDirty = False
  Exit Sub

Cancelou:
  DisplayMsg "Configuração não gravada."
  Exit Sub
  
Erro_Abrir:
  DisplayMsg "Houve erro ao criar o arquivo, configuração não gravada."
  Exit Sub
  
End Sub

Private Function Verifica_Final() As Integer
  Dim i As Long
  Dim Texto As String
  
  For i = 0 To TextGrid1.ItemCount
    Texto = TextGrid1.ItemText(i)
    If Left(Texto, 21) = "FINAL DE CONFIGURAÇÃO" Then
      Verifica_Final = TextGrid1.ItemLine(i) + 1
      Exit Function
    End If
  Next i
  
  Verifica_Final = -1
  
End Function

Private Sub B_Insere_Linha_Click()
  Dim Resp As String
  Dim Erro As Integer
  
  Call StatusMsg("")

  Resp = InputBox("Deseja inserir antes de qual linha ?", Me.Caption)
  
  Erro = False
  If IsNull(Resp) Then Erro = True
  If Erro = False Then If Resp = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Resp) Then Erro = True
  If Erro = False Then If Val(Resp) < 1 Then Erro = True
  If Erro = False Then If Val(Resp) > TextGrid1.MaxLine Then Erro = True
  
  If Erro = True Then
    DisplayMsg "Linha incorreta."
    Exit Sub
  End If
  
  TextGrid1.InsertLine (Val(Resp))
  gbDirty = True
End Sub

Private Sub B_Le_Click()
  Dim Linha As Integer
  Dim Coluna As Integer
  Dim Ajusta_Grade As Integer
  Dim Termina As Integer
  Dim Aux_Str As String
  Dim Num_Campos As Integer
  Dim i As Integer
  Dim Texto1 As String
  Dim Tipo As String
  Dim Tamanho As Integer
  Dim Final As Integer
  Dim Aux As String
  Dim Aux2 As String
  Dim Pos As Integer
  Dim Ult_Pos As Integer
  Dim Nome_Gravar As String
  Dim Cor_Campo As Long
  Dim Tipo_Campo As String
  Dim Pos2 As Long
  Dim nFileNum As Integer
  
  For i = TextGrid1.ItemCount To 0 Step -1
    TextGrid1.RemoveItem (i)
  Next i
'
  Call StatusMsg("")
  
  On Error GoTo Cancelou
  
  Dialog1.InitDir = gsConfigPath
  Dialog1.Filter = gsDialogTextFile + gsFileExtension
  Dialog1.Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly + cdlOFNLongNames
  Dialog1.ShowOpen
  If Dialog1.FileName = gsFileExtension Then Exit Sub
  On Error GoTo 0
 
  Aux = Dialog1.FileName
  If Len(Aux) < 5 Then
    DisplayMsg "Nome incorreto."
    Exit Sub
  End If
 
  Aux2 = Right(Aux, 4)
  Aux2 = UCase(Aux2)
  If Aux2 <> Right(gsFileExtension, 4) Then
   DisplayMsg "Extensão incorreta."
   Exit Sub
  End If
  
  Me.Caption = gsTitleCaption & " - [" & Aux & "]"
  Me.Refresh
  
  Nome_Gravar = ""
  Nome_Gravar = Aux
 
 
 For i = 0 To TextGrid1.MaxLine - 1
    TextGrid1.RemoveLine (i)
  Next i
  
    
  nFileNum = FreeFile
  Open Nome_Gravar For Input As #nFileNum
  
  
  Ajusta_Grade = False
    
  Linha = 1
  Coluna = 1
  
  
Lp1:
    Input #nFileNum, Aux_Str
    If EOF(1) Then GoTo Fim
    
    If Left(Aux_Str, Len(gsFileHeader)) = gsFileHeader Then GoTo Configura
    
    Aux_Str = Apaga_Aspas(Aux_Str)
    
    If Aux_Str = "*** Fim de arquivo ***" Then GoTo Fim
    If Left(Aux_Str, 13) = "[LINHA_BRANCO" Then GoTo Pula_Linha
     
    Num_Campos = Conta_Campos(Aux_Str)
    If Num_Campos = 0 Then GoTo Lp1
    
    For i = 1 To Num_Campos
     Texto1 = Separa_Campos(Aux_Str, i, Tipo)
     
     
    If Tipo = "CAMPO" Then Tamanho = Separa_Tamanho(Texto1)
    
    If Texto1 <> "" And Trim(Texto1) = "" Then
        Coluna = Coluna + Len(Texto1)
        GoTo Fim_Next
    End If
    If Texto1 = "ESPACO_BRANCO" Then
        Coluna = Coluna + Tamanho
        GoTo Fim_Next
    End If
    
    If Tipo = "TEXTO" Then
       Tamanho = Len(Texto1)
       Tipo_Campo = "Texto"
    End If
    
    If Tipo = "CAMPO" Then
      If gnTypeDoc = 2 Then
        Tipo_Campo = Pega_Tipo_Campo(gnTypeDoc, Texto1)
      End If
      Texto1 = Texto1 + Chr(160)
    End If
    
         
    Final = Coluna + Tamanho
    TextGrid1.AddItem Linha, Coluna, Final, Tamanho, Texto1
    
    
    If gnTypeDoc <> 2 Then
      Cor_Campo = Branco
    Else
      On Error Resume Next
      If Tipo_Campo = "CABEÇALHO" Then Cor_Campo = Laranja
      If Tipo_Campo = "PRODUTO" Then Cor_Campo = Amarelo
      If Tipo_Campo = "SERVIÇO" Then Cor_Campo = Verde
      If Tipo_Campo = "OUTROS" Then Cor_Campo = Branco
      If Tipo_Campo = "FATURA" Then Cor_Campo = Azul
      If Tipo_Campo = "Texto" Then Cor_Campo = Rosa
      On Error GoTo 0
    End If
    
    
    Pos2 = TextGrid1.ItemCount - 1
    
    On Error Resume Next
    TextGrid1.ItemBackColor(Pos2) = Cor_Campo
    On Error GoTo 0
   
    Coluna = Coluna + Tamanho
Fim_Next:
    Next i
    
    If Texto1 <> "LINHA_BRANCO" Then
      Linha = Linha + 1
      Coluna = 1
    End If
   
  GoTo Lp1
  
  
Configura:
  If Mid(Aux_Str, 40, 3) = "SIM" Then
    O_Comprimida.Value = 1
  Else
    O_Comprimida.Value = 0
  End If
  
  If Mid(Aux_Str, 55, 3) = "SIM" Then
    O_Oitavo.Value = 1
  Else
    O_Oitavo.Value = 0
  End If
  
  If Mid(Aux_Str, 75, 3) = "NÃO" Or Trim(Mid(Aux_Str, 75, 3)) = "" Then
    optComprPag(0).Value = True
  Else
    If IsNumeric(Mid(Aux_Str, 75, 3)) Then
      optComprPag(2).Value = True
      txtNumPol.Text = Mid(Aux_Str, 75, 3)
    Else
      optComprPag(1).Value = True
    End If
  End If
  
  GoTo Lp1
  
Pula_Linha:
  Linha = Linha + Val(Mid(Aux_Str, 15, 1))
  Coluna = 1
  GoTo Lp1
  
Fim:
  TextGrid1.AddItem Linha, Coluna, (Coluna + 21), -1, "FINAL DE CONFIGURAÇÃO"
     
  Close #nFileNum
  gbDirty = False
  Exit Sub
  
Cancelou:
  DisplayMsg "Leitura cancelada."
  
End Sub

Private Sub B_Limpa_Click()
  Dim i As Integer
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja perder as alterações e iniciar uma nova configuração?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Exit Sub
  End If
  
  For i = TextGrid1.ItemCount To 0 Step -1
    TextGrid1.RemoveItem (i)
  Next i
  
  Call StatusMsg("")
  gbDirty = True
  
End Sub

Private Sub B_OK_Fixo_Click()
  If IsNull(Texto.Text) Then Exit Sub
  If Texto.Text = "" Then Exit Sub
    
  Campo.Caption = Texto.Text
  Tamanho.Text = Len(Texto.Text)
  Tipo = "TEXTO"
  
  Cor = "Rosa"
  
End Sub

Private Sub B_Redimensiona_Click()
 Dim Erro As Integer
 
 Erro = False
 
 If IsNull(Linhas.Text) Then Erro = True
 If Erro = False Then If Linhas.Text = "" Then Erro = True
 If Erro = False Then If Not IsNumeric(Linhas.Text) Then Erro = True
 If Erro = False Then If Val(Linhas.Text) < 2 Then Erro = True
 
 If Erro = True Then
   DisplayMsg "Número de linhas deve ser entre 2 e 999."
   Linhas.SetFocus
   Exit Sub
 End If
 
 
 Erro = False
 
 If IsNull(Colunas.Text) Then Erro = True
 If Erro = False Then If Colunas.Text = "" Then Erro = True
 If Erro = False Then If Not IsNumeric(Colunas.Text) Then Erro = True
 If Erro = False Then If Val(Colunas.Text) < 2 Then Erro = True
 
 If Erro = True Then
   DisplayMsg "Número de colunas deve ser entre 2 e 999."
   Colunas.SetFocus
   Exit Sub
 End If
 
 
 Erro = MsgBox("Ao redimensionar o tamanho da grade, caso existam campos que fiquem total ou parcialmente fora da nova área escolhida, estes campos serão apagados. Deseja redimensionar ?", vbOKCancel, "Atenção")
 If Erro = vbCancel Then
    Linhas.Text = TextGrid1.MaxLine
    Colunas.Text = TextGrid1.MaxCharacter
    Exit Sub
 End If
 
 TextGrid1.MaxCharacter = Colunas.Text
 TextGrid1.MaxLine = Linhas.Text
 
End Sub

Private Sub B_Remove_Click()
  Dim Resp As String
  Dim Erro As Integer
  
  Call StatusMsg("")

  Resp = InputBox("Deseja remover qual linha ?", Me.Caption)
  
  Erro = False
  If IsNull(Resp) Then Erro = True
  If Erro = False Then If Resp = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Resp) Then Erro = True
  If Erro = False Then If Val(Resp) < 1 Then Erro = True
  If Erro = False Then If Val(Resp) > TextGrid1.MaxLine Then Erro = True
  
  If Erro = True Then
    DisplayMsg "Linha incorreta."
    Exit Sub
  End If
  
  TextGrid1.RemoveLine (Val(Resp))
  gbDirty = True
End Sub

Private Sub Colunas_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Combo_Classe_Click()
  Remonta_Lista
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  If gnTypeDoc = 2 Then
    lblClasse.Visible = True
    Combo_Classe.Visible = True
    Lista.Top = 345
    Lista.Height = 1230
  Else
    lblClasse.Visible = False
    Combo_Classe.Visible = False
    Lista.Top = 90
    Lista.Height = 1425
  End If
  Call Inicia_Lista
  Call Remonta_Lista
  gbDirty = False
  Me.Caption = gsTitleCaption
  Me.Refresh

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If gbDirty = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Arquivo modificado. Deseja sair sem gravá-lo?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbYes Then
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    Cancel = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call StatusMsg("")
End Sub

Private Sub Linhas_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Lista_DblClick()
  Campo.Caption = Lista.List(Lista.ListIndex)
  Tipo = "CAMPO"
  Máx = Pega_Tamanho(Campo.Caption)
  Tamanho.Text = Máx
  Campo.Caption = Campo.Caption + Chr(160)
  
  If gnTypeDoc <> 2 Then
    Cor = "Branco"
  Else
    If Combo_Classe.Text = "Cabeçalho / Rodapé" Then Cor = "Laranja"
    If Combo_Classe.Text = "Produtos" Then Cor = "Amarelo"
    If Combo_Classe.Text = "Serviços" Then Cor = "Verde"
    If Combo_Classe.Text = "Fatura" Then Cor = "Azul"
    If Combo_Classe.Text = "Outros" Then Cor = "Branco"
  End If

End Sub

Private Sub O_Comprimida_Click()
  gbDirty = True
End Sub

Private Sub O_Muito_Click()
  TextGrid1.CharacterWidth = 5
  TextGrid1.LineHeight = 10
  TextGrid1.Font = "Arial"
  TextGrid1.Font.Size = 6
End Sub

Private Sub O_Normal_Click()
  TextGrid1.CharacterWidth = 8
  TextGrid1.LineHeight = 16
  TextGrid1.Font = "FIXEDSYS"
  TextGrid1.Font.Size = 9
End Sub

Private Sub O_Oitavo_Click()
  gbDirty = True
End Sub

Private Sub O_Pequeno_Click()
  TextGrid1.CharacterWidth = 6
  TextGrid1.LineHeight = 12
  TextGrid1.Font = "Arial"
  TextGrid1.Font.Size = 6
End Sub

Private Sub optComprPag_Click(Index As Integer)
  txtNumPol.Enabled = optComprPag(2).Value = True
  gbDirty = True
End Sub

Private Sub Tamanho_Change()
  If Tipo = "CAMPO" Then Tamanho.Enabled = True
  If Tipo = "TEXTO" Then Tamanho.Enabled = False
End Sub

Private Sub Tamanho_LostFocus()
  Dim i As Integer
  
  If Val(Tamanho.Text) > Máx Then
    i = MsgBox("Tamanho máximo é " + str(Máx), vbOKOnly, "Atenção")
    Tamanho.SetFocus
    Exit Sub
  End If
  
End Sub

Private Sub TextGrid1_Debug(ByVal DebugString As String)
  Dim i As Integer
  Debug.Print DebugString
  i = 1
End Sub

Private Sub TextGrid1_DragDrop(Source As Control, X As Single, Y As Single)
  Dim i As Boolean
  Dim J As Long
  Dim Cor_Item As Long
  
  i = TextGrid1.DragDropItem(Campo.Caption, Val(Tamanho.Text), X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY, Val(Tamanho.Text))
  
  J = TextGrid1.ItemCount - 1
  
  If Cor = "Laranja" Then Cor_Item = Laranja
  If Cor = "Amarelo" Then Cor_Item = Amarelo
  If Cor = "Verde" Then Cor_Item = Verde
  If Cor = "Branco" Then Cor_Item = Branco
  If Cor = "Azul" Then Cor_Item = Azul
  If Cor = "Rosa" Then Cor_Item = Rosa
  
  'TextGrid1.ItemBackColor(J) = Cor_Item
  
  Call StatusMsg(i)
  
  Mostra_Campo_Atual (TextGrid1.ActiveItem)
  gbDirty = True

End Sub

Private Sub TextGrid1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  If State = 1 Then
    TextGrid1.DragItemEnd
    Source.DragIcon = LoadPicture()
  Else
    If State = 2 Then
      Source.DragIcon = Image1.Picture
    End If
    TextGrid1.DragItem Campo.Caption, Val(Tamanho.Text), X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY
  End If
  
  Mostra_Campo_Atual (TextGrid1.ActiveItem)
End Sub

Private Sub TextGrid1_ItemClick(ByVal Index As Long)
  Mostra_Campo_Atual (Index)
  gbDirty = True
End Sub

Private Sub TextGrid1_ItemResize(ByVal Index As Long)
  Mostra_Campo_Atual (Index)
End Sub
