VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmDigitaGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digitação de Produtos com Grade"
   ClientHeight    =   6315
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   11280
   Icon            =   "DigitaGrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DigitaGrade.frx":0CCA
   ScaleHeight     =   6315
   ScaleWidth      =   11280
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fechar"
      Height          =   400
      Left            =   9825
      TabIndex        =   34
      ToolTipText     =   "Fechar a janela sem confirmar digitação"
      Top             =   150
      Width           =   1335
   End
   Begin VB.CheckBox O_Mostra_Custo 
      Caption         =   "Mostra_Custo"
      Height          =   225
      Left            =   5070
      TabIndex        =   33
      Top             =   950
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Data Data1 
      Caption         =   "Produto"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   540
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Con_Produto"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1905
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown1 
      Bindings        =   "DigitaGrade.frx":1254
      Height          =   1095
      Left            =   480
      TabIndex        =   31
      Top             =   4440
      Width           =   6735
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   6720
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4445
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   11880
      _ExtentY        =   1931
      _StockProps     =   77
   End
   Begin VB.Data Data2 
      Caption         =   "Cor"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2535
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cor"
      Top             =   8295
      Visible         =   0   'False
      Width           =   2010
   End
   Begin SSDataWidgets_B.SSDBDropDown DropDown2 
      Bindings        =   "DigitaGrade.frx":1268
      Height          =   1095
      Left            =   480
      TabIndex        =   30
      Top             =   3240
      Width           =   6735
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8520
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1614
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   11880
      _ExtentY        =   1931
      _StockProps     =   77
   End
   Begin VB.CommandButton B_Confirma 
      Caption         =   "&Confirmar"
      Height          =   400
      Left            =   9840
      TabIndex        =   25
      ToolTipText     =   "Confirmar a digitação e retornar"
      Top             =   1215
      Width           =   1335
   End
   Begin VB.CheckBox O_Etiqueta 
      Caption         =   "Etiqueta"
      Height          =   225
      Left            =   5070
      TabIndex        =   24
      Top             =   735
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CheckBox O_Impostos 
      Caption         =   "Impostos"
      Height          =   225
      Left            =   5070
      TabIndex        =   23
      Top             =   525
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CheckBox O_Desconto 
      Caption         =   "Desconto"
      Height          =   225
      Left            =   5070
      TabIndex        =   22
      Top             =   315
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CheckBox O_Preço 
      Caption         =   "Preço"
      Height          =   225
      Left            =   5070
      TabIndex        =   21
      Top             =   105
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ListBox Lista 
      Height          =   1620
      Left            =   105
      MultiSelect     =   1  'Simple
      TabIndex        =   16
      Top             =   300
      Width           =   2850
   End
   Begin VB.CommandButton B_Incluir_Cor 
      Caption         =   "&Adicionar"
      Height          =   400
      Left            =   3180
      TabIndex        =   15
      ToolTipText     =   "Adicionar tamanhos selecionados a tabela de digitação"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fundo - por Quantidade"
      Height          =   1380
      Left            =   6570
      TabIndex        =   2
      Top             =   60
      Width           =   3165
      Begin VB.TextBox Fundo4 
         Height          =   285
         Left            =   420
         TabIndex        =   12
         Text            =   "8"
         Top             =   945
         Width           =   645
      End
      Begin VB.TextBox Fundo3 
         Height          =   285
         Left            =   1365
         TabIndex        =   9
         Text            =   "7"
         Top             =   630
         Width           =   540
      End
      Begin VB.TextBox Fundo2 
         Height          =   285
         Left            =   420
         TabIndex        =   7
         Text            =   "4"
         Top             =   630
         Width           =   645
      End
      Begin VB.TextBox Fundo1 
         Height          =   285
         Left            =   1365
         TabIndex        =   3
         Text            =   "3"
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label9 
         Caption         =   "MAGENTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Left            =   1995
         TabIndex        =   14
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "em diante"
         Height          =   225
         Left            =   1155
         TabIndex        =   13
         Top             =   990
         Width           =   750
      End
      Begin VB.Label Label7 
         Caption         =   "de"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   975
         Width           =   330
      End
      Begin VB.Label Label6 
         Caption         =   "AMARELO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   1995
         TabIndex        =   10
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "a"
         Height          =   225
         Left            =   1155
         TabIndex        =   8
         Top             =   690
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "de"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   690
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "de   1   a "
         Height          =   225
         Left            =   315
         TabIndex        =   5
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label4 
         Caption         =   "VERDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   1995
         TabIndex        =   4
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.CommandButton B_Começa 
      Caption         =   "&Iniciar"
      Height          =   400
      Left            =   9840
      TabIndex        =   1
      ToolTipText     =   "Iniciar a digitação dos valores para os tamanhos selecionados na tabela principal"
      Top             =   675
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Height          =   3945
      Left            =   105
      TabIndex        =   0
      Top             =   2010
      Width           =   11100
      _Version        =   196617
      DataMode        =   1
      Rows            =   500
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      ForeColorEven   =   0
      RowHeight       =   423
      SplitterPos     =   4
      Columns.Count   =   26
      Columns(0).Width=   2037
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3201
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1032
      Columns(2).Caption=   "Cor"
      Columns(2).Name =   "Cor"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1455
      Columns(3).Caption=   "Nome"
      Columns(3).Name =   "Nome_Cor"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   688
      Columns(4).Name =   "Tam1"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   688
      Columns(5).Name =   "Tam2"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   688
      Columns(6).Name =   "Tam3"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   688
      Columns(7).Name =   "Tam4"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   688
      Columns(8).Name =   "Tam5"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   688
      Columns(9).Name =   "Tam6"
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   688
      Columns(10).Name=   "Tam7"
      Columns(10).CaptionAlignment=   2
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   688
      Columns(11).Name=   "Tam8"
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   688
      Columns(12).Name=   "Tam9"
      Columns(12).CaptionAlignment=   2
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   688
      Columns(13).Name=   "Tam10"
      Columns(13).CaptionAlignment=   2
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   688
      Columns(14).Name=   "Tam11"
      Columns(14).CaptionAlignment=   2
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   688
      Columns(15).Name=   "Tam12"
      Columns(15).CaptionAlignment=   2
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   688
      Columns(16).Name=   "Tam13"
      Columns(16).CaptionAlignment=   2
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   688
      Columns(17).Name=   "Tam14"
      Columns(17).CaptionAlignment=   2
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(18).Width=   688
      Columns(18).Name=   "Tam15"
      Columns(18).CaptionAlignment=   2
      Columns(18).DataField=   "Column 18"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(19).Width=   1005
      Columns(19).Caption=   "Total"
      Columns(19).Name=   "Total"
      Columns(19).Alignment=   2
      Columns(19).CaptionAlignment=   2
      Columns(19).DataField=   "Column 19"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(19).Locked=   -1  'True
      Columns(20).Width=   1693
      Columns(20).Caption=   "Preço Unit."
      Columns(20).Name=   "Preço"
      Columns(20).DataField=   "Column 20"
      Columns(20).DataType=   8
      Columns(20).NumberFormat=   "###,##0.00"
      Columns(20).FieldLen=   256
      Columns(21).Width=   1508
      Columns(21).Caption=   "Desc. %"
      Columns(21).Name=   "Desconto"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      Columns(22).Width=   1349
      Columns(22).Caption=   "ICM (%)"
      Columns(22).Name=   "ICM"
      Columns(22).DataField=   "Column 22"
      Columns(22).DataType=   8
      Columns(22).FieldLen=   256
      Columns(23).Width=   1376
      Columns(23).Caption=   "IPI (%)"
      Columns(23).Name=   "IPI"
      Columns(23).DataField=   "Column 23"
      Columns(23).DataType=   8
      Columns(23).FieldLen=   256
      Columns(24).Width=   1852
      Columns(24).Caption=   "Valor Total"
      Columns(24).Name=   "Valor Total"
      Columns(24).DataField=   "Column 24"
      Columns(24).DataType=   8
      Columns(24).FieldLen=   256
      Columns(24).Locked=   -1  'True
      Columns(25).Width=   1720
      Columns(25).Caption=   "Etiqueta"
      Columns(25).Name=   "Etiqueta"
      Columns(25).CaptionAlignment=   2
      Columns(25).DataField=   "Column 25"
      Columns(25).DataType=   8
      Columns(25).FieldLen=   256
      Columns(25).Style=   2
      _ExtentX        =   19579
      _ExtentY        =   6959
      _StockProps     =   79
      Caption         =   "Produtos"
      Enabled         =   0   'False
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
   Begin VB.Label Label11 
      Caption         =   "Selecione os Tamanhos para digitação:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   45
      Width           =   2910
   End
   Begin VB.Label Label13 
      Height          =   405
      Left            =   0
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label_Erro 
      Caption         =   "Este produto não tem esta cor ou este tamanho."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   135
      TabIndex        =   32
      Top             =   6015
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label Itens 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7485
      TabIndex        =   29
      Top             =   5985
      Width           =   960
   End
   Begin VB.Label Label12 
      Caption         =   "Qtde de Itens :"
      Height          =   225
      Left            =   6300
      TabIndex        =   28
      Top             =   6045
      Width           =   1065
   End
   Begin VB.Label Retorno1 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8040
      TabIndex        =   27
      Top             =   1575
      Width           =   1170
   End
   Begin VB.Label Retorno 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   7095
      TabIndex        =   26
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   10350
      TabIndex        =   20
      Top             =   5985
      Width           =   810
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Total :"
      Height          =   225
      Left            =   9570
      TabIndex        =   19
      Top             =   6030
      Width           =   540
   End
   Begin VB.Label Selecionados 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   5055
      TabIndex        =   18
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Tamanhos selecionados :"
      Height          =   225
      Left            =   3165
      TabIndex        =   17
      Top             =   1575
      Width           =   1905
   End
End
Attribute VB_Name = "frmDigitaGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tamanhos As Integer
Dim Linha As Integer
Dim rsProdutos As Recordset
Dim rsCores As Recordset
Dim rsTamanhos As Recordset
Dim rsGrade As Recordset
Dim rsPreços As Recordset

Private Type TabGrade
  Código As String
  Nome As String
  Cor As Integer
  Nome_Cor As String
  Qtde1 As String
  Qtde2 As String
  Qtde3 As String
  Qtde4 As String
  Qtde5 As String
  Qtde6 As String
  Qtde7 As String
  Qtde8 As String
  Qtde9 As String
  Qtde10 As String
  Qtde11 As String
  Qtde12 As String
  Qtde13 As String
  Qtde14 As String
  Qtde15 As String
  Total As String
  Preço As Single
  Desconto As Single
  ICM As Integer
  IPI As Integer
  Valor_Total As Double
  Etiqueta As Integer
End Type

Dim Tabe(500) As TabGrade

Sub Recalcula_Linha()

  Dim Qtde As Long
  Dim i As Integer
  Dim Preço As Double
  Dim Desconto As Double
  
  Qtde = 0
  
  For i = 4 To 18
    If Not IsNull(Grade1.Columns(i).Text) Then
      If Grade1.Columns(i).Text <> "" Then
        If IsNumeric(Grade1.Columns(i).Text) Then
          Qtde = Qtde + Val(Grade1.Columns(i).Text)
        End If
      End If
    End If
  Next i
    
  Grade1.Columns(19).Text = Qtde
  
  Preço = Grade1.Columns(20).Text
  Preço = Preço * Qtde
  
  Desconto = Preço * Grade1.Columns(21).Text / 100
  
  Preço = Preço - Desconto
  
  Desconto = (Preço * Grade1.Columns(23).Text / 100)
  
  Preço = Preço + Desconto 'ipi
  
  Preço = Format(Preço, "##############0.00")
  
  Grade1.Columns(24).Text = Preço
  
End Sub

Public Sub RetornarLinhaGrade(ByVal nLine As Integer, ByRef sCodigo As String, _
  ByRef sNome As String, ByRef nCor As Integer, ByRef nQuantidade() As Integer, _
  ByRef nPreco As Single, ByRef nDesconto As Single, ByRef nICM As Integer, _
  ByRef nIPI As Integer, ByRef nValorTotal As Double, ByRef nEtiqueta As Integer)
  
  Dim nX As Integer
  
  For nX = 0 To 500
    If Tabe(nX).Código <> "" Then
      nLine = nLine - 1
    End If
    If nLine = 0 Then
      With Tabe(nX)
        sCodigo = UCase(.Código)
        sNome = .Nome
        nCor = .Cor
        
        nQuantidade(0) = Val(Tabe(nX).Qtde1)
        nQuantidade(1) = Val(Tabe(nX).Qtde2)
        nQuantidade(2) = Val(Tabe(nX).Qtde3)
        nQuantidade(3) = Val(Tabe(nX).Qtde4)
        nQuantidade(4) = Val(Tabe(nX).Qtde5)
        nQuantidade(5) = Val(Tabe(nX).Qtde6)
        nQuantidade(6) = Val(Tabe(nX).Qtde7)
        nQuantidade(7) = Val(Tabe(nX).Qtde8)
        nQuantidade(8) = Val(Tabe(nX).Qtde9)
        nQuantidade(9) = Val(Tabe(nX).Qtde10)
        nQuantidade(10) = Val(Tabe(nX).Qtde11)
        nQuantidade(11) = Val(Tabe(nX).Qtde12)
        nQuantidade(12) = Val(Tabe(nX).Qtde13)
        nQuantidade(13) = Val(Tabe(nX).Qtde14)
        nQuantidade(14) = Val(Tabe(nX).Qtde15)
        
        nPreco = .Preço
        nDesconto = .Desconto
        nICM = .ICM
        nIPI = .IPI
        nValorTotal = .Valor_Total
        nEtiqueta = .Etiqueta
      End With
      Exit Sub
    End If
  Next nX
End Sub

Public Sub RetornarTamanhos(ByRef nTamanho() As Integer)
  Dim nCol As Integer
  
  For nCol = 4 To 18
    If Grade1.Columns(nCol).Caption <> "" Then
      nTamanho(nCol - 4) = Grade1.Columns(nCol).Caption
    End If
  Next nCol
End Sub

Private Sub B_Começa_Click()
  Dim nCol As Integer
  
  Call StatusMsg("")
  
  '07/10/2003 - Maikel
  '             Mudada a forma de verificação do numero de itens selecionados no listbox. Antes tinha um bug que
  '             quando era selecionado apenas um tamanho, o sistema dava a mensagem abaixo.
  If Lista.SelCount <= 0 Then
    DisplayMsg "Escolha ao menos um tamanho antes."
    Exit Sub
  End If
  
  
  Grade1.Columns(0).DropDownHwnd = DropDown1.hwnd
  Grade1.Enabled = True
  
  For nCol = 4 To 18
    If Grade1.Columns(nCol).Caption <> "" Then
      Grade1.Columns(nCol).Visible = True
    End If
  Next nCol
  
  If O_Preço.Value = 1 Then
    Grade1.Columns(20).Visible = True
    Grade1.Columns(24).Visible = True
  Else
    Grade1.Columns(20).Visible = False
    Grade1.Columns(21).Visible = False
    Grade1.Columns(24).Visible = False
  End If
  
  If O_Desconto.Value = 1 Then
    If O_Preço.Value = 1 Then
      Grade1.Columns(21).Visible = True
    End If
  Else
    Grade1.Columns(21).Visible = False
  End If
  
  If O_Impostos.Value = 1 Then
    Grade1.Columns(22).Visible = True
    Grade1.Columns(23).Visible = True
  Else
    Grade1.Columns(22).Visible = False
    Grade1.Columns(23).Visible = False
  End If
  
  If O_Etiqueta.Value = 1 Then
    Grade1.Columns(25).Visible = True
  Else
    Grade1.Columns(25).Visible = False
  End If
  
  Grade1.MoveLast
  Grade1.MoveFirst
  
  B_Incluir_Cor.Enabled = False
  Lista.Enabled = False
  B_Começa.Enabled = False
  
  Grade1.SetFocus
  SendKeys "{Tab}"
  
End Sub

Private Sub B_Confirma_Click()
  Dim i As Integer
  Dim J As Long
  
  For i = 0 To 499
    If Tabe(i).Código <> "" Then
      J = J + 1
    End If
  Next i
  
  Retorno.Caption = "OK"
  Retorno1.Caption = J
   
  frmDigitaGrade.Hide
End Sub

Private Sub B_Incluir_Cor_Click()
  Dim i         As Integer
  Dim J         As Integer
  Dim Str_Aux   As String
  
  Tamanhos = 0
  
  If Val(Selecionados.Caption) = 0 Then
    DisplayMsg "Escolha os tamanhos a serem usados antes."
    Exit Sub
  End If
  
  If Val(Selecionados.Caption) > 15 Then
    DisplayMsg "Foram escolhidos mais de 15 tamanhos. Retire alguns e tente novamente."
    Exit Sub
  End If
  
  For i = 4 To 18
    Grade1.Columns(i).Caption = ""
  Next i
  
  For i = 0 To (Lista.ListCount - 1)
    If Lista.Selected(i) = True Then
      Tamanhos = Tamanhos + 1
      Str_Aux = Lista.List(i)
      Str_Aux = Left(Str_Aux, 3)
      J = Val(Str_Aux)
      Grade1.Columns(Tamanhos + 3).Caption = J
    End If
  Next i
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub DropDown1_CloseUp()
  Grade1.Columns(0).Text = DropDown1.Columns(1).Text
End Sub

Private Sub DropDown2_CloseUp()
  Grade1.Columns(2).Text = DropDown2.Columns(1).Text
  Grade1.Columns(3).Text = DropDown2.Columns(0).Text
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim Aux_Str As String

  Call CenterForm(Me)
  
  Limpa_Variáveis
  
  Set rsTamanhos = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsCores = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  
  Grade1.StyleSets("verde").BackColor = RGB(0, 255, 0)
  Grade1.StyleSets("amarelo").BackColor = RGB(255, 255, 0)
  Grade1.StyleSets("magenta").BackColor = RGB(255, 0, 255)
  
  Grade1.Scroll -99, -99
  
'  Data1.DatabaseName = gsQuickDBFileName
'  Data2.DatabaseName = gsQuickDBFileName
  
  Set Data1.Recordset = db.OpenRecordset(SQL_CONS_PRODUTO_GRADE, dbOpenSnapshot)
  Set Data2.Recordset = db.OpenRecordset(SQL_CONS_COR, dbOpenSnapshot)
  
  i = 0
  rsTamanhos.Index = "Código"
Lp1:
  rsTamanhos.Seek ">", i
  If rsTamanhos.NoMatch Then GoTo Fim_Tamanhos
  i = rsTamanhos("Código")
  Aux_Str = Format(rsTamanhos("Código"), "000")
  Aux_Str = Aux_Str + " - " + rsTamanhos("Nome")
  Lista.AddItem Aux_Str
  GoTo Lp1
  
Fim_Tamanhos:
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsTamanhos.Close
  rsCores.Close
  rsProdutos.Close
  rsGrade.Close
  rsPreços.Close
  Set rsTamanhos = Nothing
  Set rsCores = Nothing
  Set rsProdutos = Nothing
  Set rsGrade = Nothing
  Set rsPreços = Nothing
End Sub

Sub Limpa_Variáveis()
  Dim nX As Integer
  
  Tamanhos = 0
  Linha = 0
  
  Erase Tabe
    
  For nX = 4 To 18
    Grade1.Columns(nX).Caption = ""
    Grade1.Columns(nX).Visible = False
  Next nX
  
  Grade1.Enabled = False
  
  Grade1.MoveLast
  Grade1.MoveFirst
  
  B_Incluir_Cor.Enabled = True
  Lista.Enabled = True
  B_Começa.Enabled = True
  
End Sub

Private Sub Fundo1_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Fundo2_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Fundo3_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Fundo4_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Grade1_AfterColUpdate(ByVal ColIndex As Integer)
  Dim Aux_Total As Long
  Dim Aux As Variant
  Dim i As Integer
  
  For i = 4 To 18
    Aux = Grade1.Columns(i).Text
    If IsNull(Aux) Then
      Aux = 0
    ElseIf Aux = "" Then
      Aux = 0
    ElseIf Len(Aux) > 6 Then
      Aux = 0
    Else
      Aux = CLng(Aux)
    End If
    Aux_Total = Aux_Total + Aux
  Next i
  Grade1.Columns(19).Text = Aux_Total
  
  Call Recalcula_Linha
  
End Sub

Private Sub Grade1_AfterUpdate(RtnDispErrMsg As Integer)
  Call Recalcula
End Sub

Sub Recalcula()
  Dim Aux_Itens As Long
  Dim Aux_Total As Long
  Dim i As Integer
  
  For i = 0 To 500
   If Tabe(i).Código <> "" Then
     Aux_Itens = Aux_Itens + 1
     Aux_Total = Aux_Total + Tabe(i).Total
   End If
  Next i
  
  Itens.Caption = Aux_Itens
  Total.Caption = Aux_Total
End Sub

Private Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 Dim Erro As Integer
 Dim i As Integer
 
 Aux = Grade1.Columns(ColIndex).Value

 rsProdutos.Index = "Código"
 rsCores.Index = "Código"
 rsPreços.Index = "Tabela"

 If ColIndex = 0 Then 'Produto
   Erro = False
   
   If IsNull(Aux) Or Aux = "" Then  'Apagar linha
      Grade1.Columns(1).Text = ""
      Grade1.Columns(2).Text = 0
      Grade1.Columns(3).Text = ""
      Grade1.Columns(4).Text = ""
      Grade1.Columns(5).Text = ""
      Grade1.Columns(6).Text = ""
      Grade1.Columns(7).Text = ""
      Grade1.Columns(8).Text = ""
      Grade1.Columns(9).Text = ""
      Grade1.Columns(10).Text = ""
      Grade1.Columns(11).Text = ""
      Grade1.Columns(12).Text = ""
      Grade1.Columns(13).Text = ""
      Grade1.Columns(14).Text = ""
      Grade1.Columns(15).Text = ""
      Grade1.Columns(16).Text = ""
      Grade1.Columns(17).Text = ""
      Grade1.Columns(18).Text = ""
      Grade1.Columns(19).Text = ""
      Grade1.Columns(20).Text = 0
      Grade1.Columns(21).Text = 0
      Grade1.Columns(22).Text = 0
      Grade1.Columns(23).Text = 0
      Grade1.Columns(24).Text = 0
      Exit Sub
   End If
   
   If IsNull(Aux) Then Erro = True
   If Erro = False Then If Aux = "" Then Erro = True
  '22/04/2009 - mpdea
  'Permitir de 4 até 14 caracteres no código do produto com grade
   If Erro = False Then If Len(Aux) < 4 Then Erro = True
   If Erro = False Then If Len(Aux) > 14 Then Erro = True
   
   If Erro = True Then
     DisplayMsg "Produto incorreto."
     Cancel = True
   End If
   
   rsProdutos.Seek "=", Aux
   If rsProdutos.NoMatch Then
     DisplayMsg "Produto não encontrado."
     Cancel = True
     Exit Sub
   End If
   
   If rsProdutos("Tipo") <> "G" Then
     DisplayMsg "Este produto não usa grade."
     Cancel = True
     Exit Sub
   End If
   
   Grade1.Columns(1).Text = rsProdutos("Nome")
   If O_Mostra_Custo.Value = 1 Then
     rsPreços.Seek "=", "CUSTO", Aux
     If Not rsPreços.NoMatch Then
       Grade1.Columns(20).Text = rsPreços("Preço")
     End If
   End If
   Cancel = False
 End If
 
 If ColIndex = 2 Then  'Cor
   Erro = False
   If IsNull(Aux) Then Erro = True
   If Erro = False Then If Aux = "" Then Erro = True
   If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
   If Erro = False Then If Val(Aux) < 0 Then Erro = True
   If Erro = False Then If Val(Aux) > 999 Then Erro = True
   
   If Erro = True Then
     DisplayMsg "Cor incorreta."
     Cancel = True
     Exit Sub
   End If
   
   If Val(Aux) = 0 Then
     Grade1.Columns(2).Text = "0"
     Grade1.Columns(3).Text = ""
     Exit Sub
   End If
   
   rsCores.Seek "=", Val(Aux)
   If rsCores.NoMatch Then
     DisplayMsg "Cor não encontrada."
     Cancel = True
     Exit Sub
   End If
   Grade1.Columns(3).Value = rsCores("Nome")
  End If
  
  If ColIndex >= 4 And ColIndex <= 18 Then
    Call StatusMsg("")
    If IsNull(Aux) Then
      Exit Sub
    ElseIf Aux = "" Then
      Exit Sub
    ElseIf Not IsNumeric(Aux) Then
      DisplayMsg "Digite a quantidade."
      Grade1.SetFocus
      Cancel = True
    ElseIf Val(Aux) > 32767 Then
      DisplayMsg "Digite quantidades até 32.767."
      Grade1.SetFocus
      Cancel = True
    End If
  End If

End Sub

Private Sub Grade1_InitColumnProps()
  Grade1.Columns(2).DropDownHwnd = DropDown2.hwnd
End Sub

Private Sub Grade1_KeyPress(KeyAscii As Integer)
  Dim Coluna As Byte
  Dim nCol As Integer
  
  nCol = Grade1.Col
  If nCol >= 4 And nCol <= 18 Then
    KeyAscii = gnSomenteNumero(KeyAscii)
    If Grade1.Columns(nCol).Caption = "" Then
      KeyAscii = 0
    ElseIf Label_Erro.Visible = True Then
      KeyAscii = 0
    ElseIf Grade1.Columns(1).Text = "" Then
      KeyAscii = 0
    ElseIf Grade1.Columns(3).Text = "" Then
      KeyAscii = 0
    ElseIf IsNull(Grade1.Columns(0).Text) Or Grade1.Columns(0).Text = "" Then
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub Grade1_LostFocus()
  If Grade1.RowChanged = True Then
    Grade1.Update
  End If
End Sub

Private Sub Grade1_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  Dim Produto As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Prod As String
  Dim Coluna As Integer
  Dim Aux_Str As String

  Label_Erro.Visible = False

  If Grade1.Col < 4 Or Grade1.Col > 18 Then
    Exit Sub
  End If
  
  Coluna = Grade1.Col
  
  If Grade1.Columns(1).Text = "" Then
    Exit Sub
  End If
  
  If Grade1.Columns(3).Text = "" Then Exit Sub
  
  If Grade1.Columns(Coluna).Caption = "" Then
    Exit Sub
  End If
  
  Cor = Grade1.Columns(2).Text
  Tamanho = Grade1.Columns(Coluna).Caption
  
  Prod = Grade1.Columns(0).Text
  
  Aux_Str = "00000" + Trim(str(Tamanho))
  Aux_Str = Right(Aux_Str, 3)
  
  Prod = Prod + Aux_Str
  
  Aux_Str = "00000" + Trim(str(Cor))
  Aux_Str = Right(Aux_Str, 3)
  
  Prod = Prod + Aux_Str
  
  rsGrade.Index = "Código"
  rsGrade.Seek "=", Prod
  If rsGrade.NoMatch Then
    Label_Erro.Visible = True
  End If
    
End Sub

Private Sub Grade1_RowLoaded(ByVal Bookmark As Variant)
 Dim i As Integer
 
 For i = 3 To 18
  If Val(Grade1.Columns(i).Text) <> 0 Then
    If Val(Grade1.Columns(i).Text) <= Fundo1.Text Then
       Grade1.Columns(i).CellStyleSet "verde"
    End If
    If Val(Grade1.Columns(i).Text) >= Fundo2.Text And Val(Grade1.Columns(i).Text) <= Fundo3.Text Then
       Grade1.Columns(i).CellStyleSet "amarelo"
    End If
    If Val(Grade1.Columns(i).Text) >= Fundo4.Text Then
       Grade1.Columns(i).CellStyleSet "magenta"
    End If
  Else
    Grade1.Columns(i).CellStyleSet ""
  End If
 Next i
End Sub


Private Sub Grade1_UnboundAddData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, NewRowBookmark As Variant)
Dim Linha As Integer

Linha = Grade1.Row

 Tabe(Linha).Código = Grade1.Columns(0).Text
 Tabe(Linha).Nome = Grade1.Columns(1).Text
 Tabe(Linha).Cor = Grade1.Columns(2).Text
 Tabe(Linha).Nome_Cor = Grade1.Columns(3).Text
 Tabe(Linha).Qtde1 = Grade1.Columns(4).Text
 Tabe(Linha).Qtde2 = Grade1.Columns(5).Text
 Tabe(Linha).Qtde3 = Grade1.Columns(6).Text
 Tabe(Linha).Qtde4 = Grade1.Columns(7).Text
 Tabe(Linha).Qtde5 = Grade1.Columns(8).Text
 Tabe(Linha).Qtde6 = Grade1.Columns(9).Text
 Tabe(Linha).Qtde7 = Grade1.Columns(10).Text
 Tabe(Linha).Qtde8 = Grade1.Columns(11).Text
 Tabe(Linha).Qtde9 = Grade1.Columns(12).Text
 Tabe(Linha).Qtde10 = Grade1.Columns(13).Text
 Tabe(Linha).Qtde11 = Grade1.Columns(14).Text
 Tabe(Linha).Qtde12 = Grade1.Columns(15).Text
 Tabe(Linha).Qtde13 = Grade1.Columns(16).Text
 Tabe(Linha).Qtde14 = Grade1.Columns(17).Text
 Tabe(Linha).Qtde15 = Grade1.Columns(18).Text
 Tabe(Linha).Total = Grade1.Columns(19).Text
 Tabe(Linha).Preço = Grade1.Columns(20).Text
 Tabe(Linha).Desconto = Grade1.Columns(21).Text
 Tabe(Linha).ICM = Grade1.Columns(22).Text
 Tabe(Linha).IPI = Grade1.Columns(23).Text
 Tabe(Linha).Valor_Total = Grade1.Columns(24).Text
 Tabe(Linha).Etiqueta = Grade1.Columns(25).Value


End Sub


Private Sub Grade1_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  Dim p As Integer
  
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove < 0 Then
      p = Grade1.Rows
    Else
      p = 0
    End If
  Else
    p = StartLocation
  End If
  
  p = p + NumberOfRowsToMove
  
  NewLocation = p
    

End Sub


Private Sub Grade1_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim r, i, J, p As Integer

If IsNull(StartLocation) Then
  If ReadPriorRows Then
    p = Grade1.Rows
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
  If p < 0 Or p >= Grade1.Rows Then Exit For
     RowBuf.Value(i, 0) = Tabe(p).Código
     RowBuf.Value(i, 1) = Tabe(p).Nome
     RowBuf.Value(i, 2) = Tabe(p).Cor
     RowBuf.Value(i, 3) = Tabe(p).Nome_Cor
     RowBuf.Value(i, 4) = Tabe(p).Qtde1
     RowBuf.Value(i, 5) = Tabe(p).Qtde2
     RowBuf.Value(i, 6) = Tabe(p).Qtde3
     RowBuf.Value(i, 7) = Tabe(p).Qtde4
     RowBuf.Value(i, 8) = Tabe(p).Qtde5
     RowBuf.Value(i, 9) = Tabe(p).Qtde6
     RowBuf.Value(i, 10) = Tabe(p).Qtde7
     RowBuf.Value(i, 11) = Tabe(p).Qtde8
     RowBuf.Value(i, 12) = Tabe(p).Qtde9
     RowBuf.Value(i, 13) = Tabe(p).Qtde10
     RowBuf.Value(i, 14) = Tabe(p).Qtde11
     RowBuf.Value(i, 15) = Tabe(p).Qtde12
     RowBuf.Value(i, 16) = Tabe(p).Qtde13
     RowBuf.Value(i, 17) = Tabe(p).Qtde14
     RowBuf.Value(i, 18) = Tabe(p).Qtde15
     RowBuf.Value(i, 19) = Tabe(p).Total
     RowBuf.Value(i, 20) = Tabe(p).Preço
     RowBuf.Value(i, 21) = Tabe(p).Desconto
     RowBuf.Value(i, 22) = Tabe(p).ICM
     RowBuf.Value(i, 23) = Tabe(p).IPI
     RowBuf.Value(i, 24) = Tabe(p).Valor_Total
     RowBuf.Value(i, 25) = Tabe(p).Etiqueta
     
     
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


Private Sub Grade1_UnboundWriteData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, WriteLocation As Variant)
  Dim nLinha As Integer
  
  nLinha = WriteLocation

  With Tabe(nLinha)
    .Código = Grade1.Columns(0).Text
    .Nome = Grade1.Columns(1).Text
    .Cor = Grade1.Columns(2).Text
    .Nome_Cor = Grade1.Columns(3).Text
    .Qtde1 = Grade1.Columns(4).Text
    .Qtde2 = Grade1.Columns(5).Text
    .Qtde3 = Grade1.Columns(6).Text
    .Qtde4 = Grade1.Columns(7).Text
    .Qtde5 = Grade1.Columns(8).Text
    .Qtde6 = Grade1.Columns(9).Text
    .Qtde7 = Grade1.Columns(10).Text
    .Qtde8 = Grade1.Columns(11).Text
    .Qtde9 = Grade1.Columns(12).Text
    .Qtde10 = Grade1.Columns(13).Text
    .Qtde11 = Grade1.Columns(14).Text
    .Qtde12 = Grade1.Columns(15).Text
    .Qtde13 = Grade1.Columns(16).Text
    .Qtde14 = Grade1.Columns(17).Text
    .Qtde15 = Grade1.Columns(18).Text
    .Total = Grade1.Columns(19).Text
    .Preço = Grade1.Columns(20).Text
    .Desconto = Grade1.Columns(21).Text
    .ICM = Grade1.Columns(22).Text
    .IPI = Grade1.Columns(23).Text
    .Valor_Total = Grade1.Columns(24).Text
    .Etiqueta = Grade1.Columns(25).Value
  End With
End Sub

Private Sub Grade1_Validate(Cancel As Boolean)
  If Grade1.RowChanged Then
    Grade1.Update
  End If
End Sub

Private Sub Lista_Click()
  Selecionados.Caption = Lista.SelCount
End Sub
