VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Grade de Produtos - TAMANHO e COR"
   ClientHeight    =   7485
   ClientLeft      =   765
   ClientTop       =   1980
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1380
   Icon            =   "Grade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7485
   ScaleWidth      =   9210
   Begin VB.TextBox txt_alvoWebApi 
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
      Left            =   7590
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_valorProduto 
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
      Left            =   7590
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1020
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_unidadeProduto 
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
      Left            =   7830
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton B_Grava 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Gravar"
      Height          =   400
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Grava Grade"
      Top             =   5280
      Width           =   1905
   End
   Begin VB.CommandButton B_Limpa 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Limpar"
      Height          =   400
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Limpa Grade"
      Top             =   6990
      Width           =   1905
   End
   Begin VB.CommandButton B_Cria 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Criar"
      Height          =   400
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cria Grade"
      Top             =   2700
      Width           =   1905
   End
   Begin VB.Data Data2 
      Caption         =   "Tamanho"
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
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Tamanho"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Cor"
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
      Left            =   4020
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cor"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBGrid Grade_Velha 
      Height          =   2175
      Left            =   60
      TabIndex        =   10
      Top             =   480
      Width           =   7095
      _Version        =   196617
      DataMode        =   1
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
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   4286
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2858
      Columns(1).Caption=   "Tamanho"
      Columns(1).Name =   "Tamanho"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3519
      Columns(2).Caption=   "Cor"
      Columns(2).Name =   "Cor"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      UseDefaults     =   0   'False
      _ExtentX        =   12515
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Grade &Existente"
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown2 
      Bindings        =   "Grade.frx":4E95A
      Height          =   855
      Left            =   1140
      TabIndex        =   6
      Top             =   4170
      Width           =   1815
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
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
      Columns(0).Width=   3200
      _ExtentX        =   3201
      _ExtentY        =   1508
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown1 
      Bindings        =   "Grade.frx":4E96E
      Height          =   855
      Left            =   4860
      TabIndex        =   5
      Top             =   4170
      Width           =   1935
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
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
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   1508
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid Grade_Nova 
      Height          =   2115
      Left            =   60
      TabIndex        =   4
      Top             =   5280
      Width           =   7095
      _Version        =   196617
      DataMode        =   1
      Rows            =   900
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowDelete     =   -1  'True
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   2831
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2752
      Columns(1).Caption=   "Tamanho"
      Columns(1).Name =   "Tamanho"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2540
      Columns(2).Caption=   "Cor"
      Columns(2).Name =   "Cor"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1482
      Columns(3).Caption=   "Estoque"
      Columns(3).Name =   "Estoque"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1561
      Columns(4).Caption=   "Etiqueta"
      Columns(4).Name =   "Etiqueta"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   11
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      UseDefaults     =   0   'False
      _ExtentX        =   12515
      _ExtentY        =   3731
      _StockProps     =   79
      Caption         =   "Grade &Nova"
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
   Begin SSDataWidgets_B.SSDBGrid Grade_Tam 
      Height          =   2535
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   3435
      _Version        =   196617
      DataMode        =   1
      Rows            =   30
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   1323
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Tamanho"
      Columns(1).Name =   "Tamanho"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      UseDefaults     =   0   'False
      _ExtentX        =   6059
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "&Tamanhos"
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
   Begin SSDataWidgets_B.SSDBGrid Grade_Cor 
      Height          =   2535
      Left            =   3720
      TabIndex        =   2
      Top             =   2700
      Width           =   3435
      _Version        =   196617
      DataMode        =   1
      Rows            =   50
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   1244
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3360
      Columns(1).Caption=   "Cor"
      Columns(1).Name =   "Cor"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      UseDefaults     =   0   'False
      _ExtentX        =   6059
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "C&ores"
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
   Begin VB.Label Classe 
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
      Height          =   255
      Left            =   7500
      TabIndex        =   12
      Top             =   1650
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Sub_Classe 
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
      Height          =   255
      Left            =   7860
      TabIndex        =   11
      Top             =   1650
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Nome_Prod 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2220
      TabIndex        =   1
      Top             =   90
      Width           =   6915
   End
   Begin VB.Label Cod_Prod 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2055
   End
End
Attribute VB_Name = "frmGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCores As Recordset
Dim rsTamanhos As Recordset
Dim rsGrade As Recordset
Dim Rec_Grade As Recordset
Dim rsEstoque As Recordset
'Dim rsProduto As Recordset
Dim rsResumo As Recordset
Dim rsFuncionarios As Recordset
Dim rsEtiquetas As Recordset

Private Type Tabela
  Código As Integer
  Nome As String
End Type

Dim Tabe_Cor(50) As Tabela
Dim Tabe_Tam(30) As Tabela

Private Type Tabela2
  Código As String
  Tamanho As String
  Cor As String
  Estoque As Long
  Etiqueta As Integer
End Type

Private Type Tabela3
  Código As String
  Tamanho As String
  Cor As String
End Type

Dim Tabe_Nova(900) As Tabela2
Dim Tabe_Velha(900) As Tabela3

' Variaveis para tratamento do Componente WebApi
Dim rsCNPJ As Recordset
Dim http As MSXML2.XMLHTTP
Dim sRETORNO_ERRO_WEBAPI As String
Dim alvoEnderecoWEBAPI As String
'Dim tentativasWEBAPI As Integer
Dim statusProdutoStatusWise As String
Dim sNomeProdutoFormatado As String
Dim sCodigoProdutoFormatado As String

Private Sub Inicia_Estoque(ByVal sProduto As String, ByVal nTamanho As Integer, ByVal nCor As Integer, ByVal nEstoque As Long)

'  Dim Mes_Atual As Integer
'  Dim Ano_Atual As Integer
'  Dim Aux_Str As String
'
'  Aux_Str = Format$(Data_Atual, "dd/mm/yyyy")
'  Ano_Atual = Val(Right(Aux_Str, 4))
'  Mes_Atual = Val(Mid(Aux_Str, 4, 2))
    
  With rsEstoque
    .Index = "Data"
    .Seek ">", gnCodFilial, sProduto, nTamanho, nCor, 0
    If Not .NoMatch Then
      If !Produto = sProduto And !Tamanho = nTamanho And !Cor = nCor Then
        'Estoque deste produto já foi inicializado.
        Exit Sub
      End If
    End If
    'Adiciona registro
    .AddNew
    !Filial = gnCodFilial
    !Data = Data_Atual
    !Produto = sProduto
    !Tamanho = nTamanho
    !Cor = nCor
    !Classe = Classe.Caption
    ![Sub Classe] = Sub_Classe.Caption
    ![Ajuste Entra] = nEstoque
    ![Estoque Final] = nEstoque
    .Update
  End With
  'Grava as informações na Tabela de Estoque Final
  Call Grava_Estoque_Final(gnCodFilial, sProduto, nTamanho, nCor, 0, CSng(nEstoque), CDate(Data_Atual))

End Sub

Private Sub B_Cria_Click()
  Dim Conta As Integer
  Dim i As Integer
  Dim I_Cor As Integer
  Dim I_Tam As Integer
  Dim I_Geral As Integer
  Dim Aux1 As String
  Dim Aux2 As String
  Dim Aux3 As String
 
  Conta = 0
  For i = 0 To 49
    If Tabe_Cor(i).Código <> 0 Then Conta = Conta + 1
  Next i
  
  If Conta = 0 Then
    DisplayMsg "Impossível criar grade, não existe nenhuma cor."
    Exit Sub
  End If
 
  Conta = 0
  For i = 0 To 29
    If Tabe_Tam(i).Código <> 0 Then Conta = Conta + 1
  Next i
  
  If Conta = 0 Then
    DisplayMsg "Impossível criar grade, não existe nenhum tamanho."
    Exit Sub
  End If
 
  I_Geral = 0
  For I_Tam = 0 To 29
    For I_Cor = 0 To 49
      If Tabe_Cor(I_Cor).Código <> 0 And Tabe_Tam(I_Tam).Código <> 0 Then
        Aux1 = "000" + LTrim(str$(Tabe_Cor(I_Cor).Código))
        Aux1 = Right$(Aux1, 3)
        Aux2 = "000" + LTrim(str$(Tabe_Tam(I_Tam).Código))
        Aux2 = Right$(Aux2, 3)
        Aux3 = Cod_Prod.Caption
        Aux3 = Aux3 + Aux2 + Aux1
        Tabe_Nova(I_Geral).Código = Trim(Aux3)
        
        Aux1 = "00000" + Trim(str(Tabe_Tam(I_Tam).Código))
        Aux1 = Right(Aux1, 3)
        Tabe_Nova(I_Geral).Tamanho = Aux1 + "-" + Tabe_Tam(I_Tam).Nome
        
        Aux1 = "00000" + Trim(str(Tabe_Cor(I_Cor).Código))
        Aux1 = Right(Aux1, 3)
        Tabe_Nova(I_Geral).Cor = Aux1 + "-" + Tabe_Cor(I_Cor).Nome
        
        I_Geral = I_Geral + 1
        sNomeProdutoFormatado = Nome_Prod + " " + Tabe_Cor(I_Cor).Nome + " " + Tabe_Tam(I_Tam).Nome
      End If
     Next I_Cor
  Next I_Tam
 
  Grade_Nova.MoveLast
  Grade_Nova.MoveFirst
 
End Sub

'Private Sub ChamadaWebApiWise()
'  On Error GoTo ErrChamadaWebApiWise
'
'    Dim cod_produto As String
'    Dim Nome_Produto As String
'    Dim valor_produto As String
'    Dim unidade_produto As String
'    Dim nI As Integer
'    Dim retWebApi As Boolean
'    Dim cnpj As String
'    Dim msgRetWebApi As String
'
'    cod_produto = sCodigoProdutoFormatado
'    Nome_Produto = sNomeProdutoFormatado
'    unidade_produto = txt_unidadeProduto.Text
'    valor_produto = txt_valorProduto.Text
'
'    alvoEnderecoWEBAPI = txt_alvoWebApi.Text
'    If alvoEnderecoWEBAPI = "INTEGRACAO_WISE=NAO" Then
'      bolINTEGRACAO_WISE = False
'    Else
'      bolINTEGRACAO_WISE = True
'    End If
'
'    If bolINTEGRACAO_WISE = True Then
'
'      rsCNPJ.MoveFirst
'
'      Do While Not rsCNPJ.EOF
'          cnpj = rsCNPJ.Fields(0)
'          cnpj = Trim(cnpj)
'          cnpj = Replace(cnpj, ".", "")
'          cnpj = Replace(cnpj, "/", "")
'          cnpj = Replace(cnpj, "-", "")
'
'          If cnpj <> "" Then
'            tentativasWEBAPI = 0
'
'            ' Criar produto no wise
'            retWebApi = WebRequest(alvoEnderecoWEBAPI + "api/product/InsertProduct?merchantId=" + cnpj + "&name=" + Nome_Produto + "&code=" + cod_produto + "&value=" + valor_produto + "&description=&ean=&unit=" + unidade_produto)
'
'            If retWebApi = False Then
'                MsgBox "Atenção: Erro de integração deste produto junto ao Sistema Wise! Erro: " + sRETORNO_ERRO_WEBAPI, vbInformation
'            End If
'          End If
'
'          rsCNPJ.MoveNext
'        Loop
'    End If
'
'    Exit Sub
'
'ErrChamadaWebApiWise:
'  'msgChamadaWebApiWise = Err.Description
'  MsgBox "Atenção: Erro de integração deste produto junto ao Sistema Wise! Erro: (Rotina ChamadaWebApiWise) " + Err.Description, vbInformation
'End Sub
'
'Public Function WebRequest(url As String) As Boolean
'On Error GoTo trata_WebApiErro
'    ' Função que chama o componente WebApi
'    Dim retWebApiAux As String
'
'    http.Open "GET", url, False
'    http.Send
'
'    retWebApiAux = http.statusText
'
'    If retWebApiAux = "OK" Then
'        sRETORNO_ERRO_WEBAPI = "Cadastrado com sucesso!"
'        WebRequest = True
'    Else
'        sRETORNO_ERRO_WEBAPI = http.responseText
'        WebRequest = False
'    End If
'
'    Exit Function
'
'trata_WebApiErro:
'    Dim iInStrRet As Integer
'    iInStrRet = InStr(1, Err.Description, "timed out")
'    If iInStrRet > 0 And tentativasWEBAPI < 4 Then
'        tentativasWEBAPI = tentativasWEBAPI + 1
'        WebRequest url
'    End If
'    sRETORNO_ERRO_WEBAPI = Err.Description
'End Function

Private Sub B_Grava_Click()
  Dim Conta As Integer
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Estoque As Long
  Dim Aux_Str As String
  Dim Cód As String
  Dim Etiquetas As Long
  Dim nRow As Long
  Dim bm As Variant
  Dim sValor As String
  
  On Error GoTo Processa_Erro

'------------------------------------------------------------------------------
'  For nRow = 0 To Grade_Nova.Rows
'    bm = Grade_Nova.GetBookmark(nRow)
'    sValor = Grade_Nova.Columns(0).CellText(bm)
'    If gsHandleNull(sValor) <> "0" Then
'      Exit For
'    End If
'  Next nRow
'  If nRow > Grade_Nova.Rows Then
'    DisplayMsg "Crie a grade antes de gravar."
'    Exit Sub
'  End If
  
  '14/11/2002 - mpdea
  'Modificado esquema para verificar se a grade foi criada
  If Tabe_Nova(0).Código = "" Then
    DisplayMsg "Crie a grade antes de gravar."
    Exit Sub
  End If
'------------------------------------------------------------------------------
  
  Call StatusMsg("Gravando grade, e inicializando o estoque.")
  
'  Set rsProduto = db.OpenRecordset("Produtos", , dbReadOnly)
'  Set rsResumo = db.OpenRecordset("Resumo Produtos")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  
  rsGrade.Index = "Código"
  rsEtiquetas.Index = "Funcionário"
  For Conta = 0 To 899
    If Tabe_Nova(Conta).Código <> "" And Tabe_Nova(Conta).Código <> "0" Then
      rsGrade.Seek "=", Tabe_Nova(Conta).Código
      If rsGrade.NoMatch Then
        rsGrade.AddNew
      Else
        rsGrade.Edit
      End If
      rsGrade("Código") = Tabe_Nova(Conta).Código
      rsGrade("Código Original") = Cod_Prod.Caption
      rsGrade.Update
      
      Aux_Str = Tabe_Nova(Conta).Código
      Aux_Str = Trim(Aux_Str)
      Cód = Cod_Prod.Caption
      sCodigoProdutoFormatado = Aux_Str
      
      Tamanho = Val(Left(Tabe_Nova(Conta).Tamanho, 3))
      Cor = Val(Left(Tabe_Nova(Conta).Cor, 3))
      Estoque = Tabe_Nova(Conta).Estoque
      'If Estoque > 0 Then
          Call Inicia_Estoque(Cód, Tamanho, Cor, Estoque)
      'End If
      'Grava Etiquetas
      If Tabe_Nova(Conta).Etiqueta = True Then
        rsEtiquetas.Seek "=", 1, Cód, Tamanho, Cor
        If rsEtiquetas.NoMatch Then
          rsEtiquetas.AddNew
          rsEtiquetas("Funcionário") = gnUserCode
          rsEtiquetas("Produto") = Cód
          rsEtiquetas("Tamanho") = Tamanho
          rsEtiquetas("Cor") = Cor
          Etiquetas = 0
        Else
          rsEtiquetas.Edit
          Etiquetas = rsEtiquetas("Qtde")
        End If
        rsEtiquetas("Qtde") = Etiquetas + Tabe_Nova(Conta).Estoque
        rsEtiquetas.Update
      End If
    End If
  Next Conta
  
  
  '02/07/2004 - mpdea
  'Atualiza a data de alteração do produto
  db.Execute "UPDATE Produtos SET [Data Alteração] = '" & _
             Format(Date, "DD/MM/YYYY") & _
             "' WHERE Código = '" & Cód & "'"
  
 
  Call StatusMsg("")
  DisplayMsg "Grade gravada."
  
  Exit Sub

Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar Grade."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub B_Limpa_Click()
 Dim i As Integer
 
' For i = 0 To 899
'   Tabe_Nova(i).Código = 0
'   Tabe_Nova(i).Tamanho = ""
'   Tabe_Nova(i).Cor = ""
'   Tabe_Nova(i).Estoque = 0
'   Tabe_Nova(i).Etiqueta = False
' Next i
 
 Erase Tabe_Nova
 
 Grade_Nova.MoveLast
 Grade_Nova.MoveFirst
End Sub

Private Sub DropDown1_CloseUp()
  Grade_Cor.Columns(0).Text = DropDown1.Columns(1).Text
End Sub

Private Sub DropDown2_CloseUp()
  Grade_Tam.Columns(0).Text = DropDown2.Columns(1).Text
End Sub

Private Sub Form_Activate()
  Dim Conta As Integer
  Dim sSql As String
  Dim Prod, Prod2 As String
  Dim Cor As Integer
  Dim Tamanho As Integer
  Dim Pos As Long
  Dim Aux As String
 
  Prod = Cod_Prod.Caption
 
  Call StatusMsg("Aguarde, iniciando tabelas...")
  
  Erase Tabe_Cor
  Erase Tabe_Tam
  Erase Tabe_Nova
  Erase Tabe_Velha
 
  Grade_Cor.MoveLast
  Grade_Cor.MoveFirst
  Grade_Tam.MoveLast
  Grade_Tam.MoveFirst
  Grade_Nova.MoveLast
  Grade_Nova.MoveFirst
  Grade_Velha.Rows = 900 '?
 
  Call StatusMsg("")
 
  'Monta grade existente
  rsGrade.Index = "Original"
  rsTamanhos.Index = "Código"
  rsCores.Index = "Código"
  Prod2 = ""
  Pos = 0
 
Lp1:
  rsGrade.Seek ">", Prod, Prod2
  If rsGrade.NoMatch Then
    GoTo Fim
  ElseIf rsGrade("Código Original") <> Prod Then
    GoTo Fim
  End If
  
  Prod2 = rsGrade("Código")
  
  Aux = Right(Prod2, 6)
  Tamanho = Left(Aux, 3)
  Cor = Right(Aux, 3)
   
  Tabe_Velha(Pos).Código = Prod2
   
  rsTamanhos.Seek "=", Tamanho
  If Not rsTamanhos.NoMatch Then
    Tabe_Velha(Pos).Tamanho = rsTamanhos("Nome")
  End If
  
  rsCores.Seek "=", Cor
  If Not rsCores.NoMatch Then
    Tabe_Velha(Pos).Cor = rsCores("Nome")
  End If
  
  Pos = Pos + 1
  
  GoTo Lp1
 
Fim:
  
  Grade_Velha.MoveLast
  Grade_Velha.MoveFirst
  Grade_Velha.Refresh
  
  Grade_Cor.SetFocus
  SendKeys "{Tab}"
  
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade")
  Set rsEstoque = db.OpenRecordset("Estoque")
  
  '******************************* INICIO:     Tratamento para o WEBAPI
  'Recuperar todos os CNPJs
  Set rsCNPJ = db.OpenRecordset("Select CGC From [Parâmetros Filial]")
  
  Set http = CreateObject("MSXML2.ServerXMLHTTP")
  '******************************* FIM:        Tratamento para o WEBAPI
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  If Not rsCNPJ Is Nothing Then
    rsCNPJ.Close
  End If
  Set rsCNPJ = Nothing
  Set http = Nothing
End Sub

Private Sub Grade_Cor_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 
 Aux = Grade_Cor.Columns(ColIndex).Text
 
 If IsNull(Aux) Then Aux = ""
 If Aux = "" Then Aux = 0
 
 Call StatusMsg("")
 
 If ColIndex = 0 Then
   If Not IsNumeric(Aux) Then
     DisplayMsg "Digite um número."
     Cancel = True
     Exit Sub
   End If
   If Val(Aux) = 0 Then
     Grade_Cor.Columns(1).Text = ""
     Exit Sub
   End If
   rsCores.Index = "Código"
   rsCores.Seek "=", Val(Aux)
   If rsCores.NoMatch Then
     DisplayMsg "Cor não encontrada."
     Cancel = True
     Exit Sub
   End If
   Grade_Cor.Columns(1).Text = rsCores("Nome")
    
 End If
  
End Sub

Private Sub Grade_Cor_InitColumnProps()
  Grade_Cor.Columns(0).DropDownHwnd = DropDown1.hwnd
End Sub

Private Sub Grade_Cor_LostFocus()
 Grade_Cor.MoveNext
 Grade_Cor.MovePrevious
End Sub

Private Sub Grade_Cor_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)
  Dim Linha As Integer
  
  Linha = Grade_Cor.Row
  
  Tabe_Cor(Linha).Código = Grade_Cor.Columns(0).Text
  Tabe_Cor(Linha).Nome = Grade_Cor.Columns(1).Text

End Sub

Private Sub Grade_Cor_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade_Cor.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p
End Sub

Private Sub Grade_Cor_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim r, i, J, p As Integer
  
  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      p = Grade_Cor.Rows
    Else
      p = 0
    End If
   Else
    p = StartLocation
    If ReadPriorRows Then
      p = p - 1
    Else
      p = p + 1
    End If
  End If
  
  For i = 0 To RowBuf.RowCount - 1
    If p < 0 Or p >= Grade_Cor.Rows Then Exit For
       RowBuf.Value(i, 0) = Tabe_Cor(p).Código
       RowBuf.Value(i, 1) = Tabe_Cor(p).Nome
       
     RowBuf.Bookmark(i) = p
     If ReadPriorRows Then
       p = p - 1
     Else
       p = p + 1
     End If
     
     r = r + 1
   Next i
   
   RowBuf.RowCount = r
   
End Sub

Private Sub Grade_Cor_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
 Dim Linha As Integer
 
 Linha = WriteLocation

 If IsNull(Grade_Cor.Columns(0).Text) Then Exit Sub
 If Grade_Cor.Columns(0).Text = "" Then Exit Sub
 If Not IsNumeric(Grade_Cor.Columns(0).Text) Then Exit Sub
 If IsNull(Grade_Cor.Columns(1).Text) Then Exit Sub
 If Grade_Cor.Columns(1).Text = "" Then Exit Sub
 

Tabe_Cor(Linha).Código = Grade_Cor.Columns(0).Text
Tabe_Cor(Linha).Nome = Grade_Cor.Columns(1).Text


End Sub

Private Sub Grade_Nova_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  If MsgBox("Deseja retirar este produto da grade?", vbQuestion + vbYesNo, "Apagar") = vbNo Then
    Cancel = 1
  Else
    DispPromptMsg = False
  End If
End Sub

Private Sub Grade_Nova_LostFocus()
  If Grade_Nova.RowChanged Then
    Grade_Nova.Update
  End If
End Sub

Private Sub Grade_Nova_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)
  Dim Linha As Integer
  
  Linha = Grade_Nova.Row
  With Tabe_Nova(Linha)
    .Código = Grade_Nova.Columns(0).Text
    .Tamanho = Grade_Nova.Columns(1).Text
    .Cor = Grade_Nova.Columns(2).Text
    .Estoque = Grade_Nova.Columns(3).Text
    .Etiqueta = Grade_Nova.Columns(4).Value
  End With
End Sub

Private Sub Grade_Nova_UnboundDeleteRow(Bookmark As Variant)
 Dim Conta, Ini As Integer
 Dim Aux As Variant
 
 
 Aux = Bookmark
 
 Ini = Val(Aux)
 
 For Conta = Ini To 898
   Tabe_Nova(Conta).Código = Tabe_Nova(Conta + 1).Código
   Tabe_Nova(Conta).Tamanho = Tabe_Nova(Conta + 1).Tamanho
   Tabe_Nova(Conta).Cor = Tabe_Nova(Conta + 1).Cor
   Tabe_Nova(Conta).Estoque = Tabe_Nova(Conta + 1).Estoque
   Tabe_Nova(Conta).Etiqueta = Tabe_Nova(Conta + 1).Etiqueta
Next Conta

Tabe_Nova(899).Código = 0
Tabe_Nova(899).Tamanho = ""
Tabe_Nova(899).Cor = ""
Tabe_Nova(899).Estoque = 0
Tabe_Nova(899).Etiqueta = False

Grade_Nova.MoveLast
Grade_Nova.MoveFirst
 
End Sub

Private Sub Grade_Nova_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade_Nova.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub

Private Sub Grade_Nova_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim r, i, J, p As Integer

If IsNull(StartLocation) Then
  If ReadPriorRows Then
    p = Grade_Nova.Rows
  Else
    p = 0
  End If
 Else
  p = StartLocation
  If ReadPriorRows Then
    p = p - 1
  Else
    p = p + 1
  End If
End If

For i = 0 To RowBuf.RowCount - 1
  If p < 0 Or p >= Grade_Nova.Rows Then Exit For
     RowBuf.Value(i, 0) = Tabe_Nova(p).Código
     RowBuf.Value(i, 1) = Tabe_Nova(p).Tamanho
     RowBuf.Value(i, 2) = Tabe_Nova(p).Cor
     RowBuf.Value(i, 3) = Tabe_Nova(p).Estoque
     RowBuf.Value(i, 4) = Tabe_Nova(p).Etiqueta
     
   RowBuf.Bookmark(i) = p
   If ReadPriorRows Then
     p = p - 1
   Else
     p = p + 1
   End If
   
   r = r + 1
 Next i
 
 RowBuf.RowCount = r

End Sub

Private Sub Grade_Nova_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
  Dim Linha As Integer
  
  Linha = WriteLocation
  
  If IsNull(Grade_Nova.Columns(3).Text) Then Grade_Nova.Columns(3).Text = 0
  If Not IsNumeric(Grade_Nova.Columns(3).Text) Then Grade_Nova.Columns(3).Text = 0
  
  Tabe_Nova(Linha).Código = Grade_Nova.Columns(0).Text
  Tabe_Nova(Linha).Tamanho = Grade_Nova.Columns(1).Text
  Tabe_Nova(Linha).Cor = Grade_Nova.Columns(2).Text
  Tabe_Nova(Linha).Estoque = Grade_Nova.Columns(3).Text
  Tabe_Nova(Linha).Etiqueta = Grade_Nova.Columns(4).Value

End Sub

Private Sub Grade_Tam_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  
  Aux = Grade_Tam.Columns(ColIndex).Text
  
  If IsNull(Aux) Then Aux = ""
  If Aux = "" Then Aux = 0
  
  Call StatusMsg("")
  
  If ColIndex = 0 Then
    If Not IsNumeric(Aux) Then
      DisplayMsg "Digite um número."
      Cancel = True
      Exit Sub
    End If
    If Val(Aux) = 0 Then
      Grade_Tam.Columns(1).Text = ""
      Exit Sub
    End If
    rsTamanhos.Index = "Código"
    rsTamanhos.Seek "=", Val(Aux)
    If rsTamanhos.NoMatch Then
      DisplayMsg "Tamanho não encontrado."
      Cancel = True
      Exit Sub
    End If
    Grade_Tam.Columns(1).Text = rsTamanhos("Nome")
  End If
End Sub

Private Sub Grade_Tam_InitColumnProps()
  Grade_Tam.Columns(0).DropDownHwnd = DropDown2.hwnd
End Sub

Private Sub Grade_Tam_LostFocus()
 Grade_Tam.MoveNext
 Grade_Tam.MovePrevious
End Sub

Private Sub Grade_Tam_UnboundAddData(ByVal RowBuf As ssRowBuffer, NewRowBookmark As Variant)
  Dim nLinha As Integer
  
  nLinha = Grade_Tam.Row
  
  Tabe_Tam(nLinha).Código = Grade_Tam.Columns(0).Text
  Tabe_Tam(nLinha).Nome = Grade_Tam.Columns(1).Text
End Sub

Private Sub Grade_Tam_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade_Tam.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub

Private Sub Grade_Tam_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim r As Integer
  Dim i As Integer
  Dim J As Integer
  Dim p As Integer

  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      p = Grade_Tam.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
    If ReadPriorRows Then
      p = p - 1
    Else
      p = p + 1
    End If
  End If

  For i = 0 To RowBuf.RowCount - 1
    If p < 0 Or p >= Grade_Tam.Rows Then
      Exit For
    Else
      RowBuf.Value(i, 0) = Tabe_Tam(p).Código
      RowBuf.Value(i, 1) = Tabe_Tam(p).Nome
      RowBuf.Bookmark(i) = p
      If ReadPriorRows Then
        p = p - 1
      Else
        p = p + 1
      End If
      r = r + 1
    End If
  Next i
  RowBuf.RowCount = r
   
End Sub

Private Sub Grade_Tam_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
  Dim nLinha As Integer
  
  nLinha = WriteLocation
  If IsNull(Grade_Tam.Columns(0).Text) Then
    Exit Sub
  ElseIf Grade_Tam.Columns(0).Text = "" Then
    Exit Sub
  ElseIf Not IsNumeric(Grade_Tam.Columns(0).Text) Then
    Exit Sub
  ElseIf IsNull(Grade_Tam.Columns(1).Text) Then
    Exit Sub
  ElseIf Grade_Tam.Columns(1).Text = "" Then
    Exit Sub
  Else
    Tabe_Tam(nLinha).Código = Grade_Tam.Columns(0).Text
    Tabe_Tam(nLinha).Nome = Grade_Tam.Columns(1).Text
  End If
End Sub

Private Sub Grade_Velha_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade_Velha.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p

End Sub

Private Sub Grade_Velha_UnboundReadData(ByVal RowBuf As ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  Dim r As Integer
  Dim i As Integer
  Dim J As Integer
  Dim p As Integer

  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      p = Grade_Velha.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
    If ReadPriorRows Then
      p = p - 1
    Else
      p = p + 1
    End If
  End If
  
  For i = 0 To RowBuf.RowCount - 1
    If p < 0 Or p >= Grade_Velha.Rows Then
      Exit For
    Else
      RowBuf.Value(i, 0) = Tabe_Velha(p).Código
      RowBuf.Value(i, 1) = Tabe_Velha(p).Tamanho
      RowBuf.Value(i, 2) = Tabe_Velha(p).Cor
      RowBuf.Bookmark(i) = p
      If ReadPriorRows Then
        p = p - 1
      Else
        p = p + 1
      End If
      r = r + 1
    End If
  Next i
  RowBuf.RowCount = r
End Sub

Private Sub Grade_Velha_UnboundWriteData(ByVal RowBuf As ssRowBuffer, WriteLocation As Variant)
  Dim nLinha As Integer
  
  nLinha = WriteLocation
  With Tabe_Velha(nLinha)
    .Código = Grade_Velha.Columns(0).Text
    .Tamanho = Grade_Velha.Columns(1).Text
    .Cor = Grade_Velha.Columns(2).Text
  End With
End Sub
