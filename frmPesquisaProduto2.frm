VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPesquisaProduto2 
   BackColor       =   &H00FFA324&
   BorderStyle     =   0  'None
   Caption         =   " Pesquisa de Produtos Alfa"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesquisaProduto2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   17610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_abaixo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15870
      Picture         =   "frmPesquisaProduto2.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   600
      Width           =   945
   End
   Begin VB.CommandButton cmd_Acima 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   15870
      Picture         =   "frmPesquisaProduto2.frx":5844C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   60
      Width           =   945
   End
   Begin VB.CommandButton cmd_direita 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   16830
      Picture         =   "frmPesquisaProduto2.frx":61F3E
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmd_esquerda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   15120
      Picture         =   "frmPesquisaProduto2.frx":6BA30
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmd_favoritos5 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   16380
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1170
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos6 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   16380
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos7 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   16380
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2670
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos8 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   16380
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3450
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos1 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1170
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos2 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos3 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2670
      Width           =   1200
   End
   Begin VB.CommandButton cmd_favoritos4 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3450
      Width           =   1200
   End
   Begin VB.CommandButton cmd_fecharTela 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13770
      Picture         =   "frmPesquisaProduto2.frx":75522
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_abaixoCorrecel2 
      BackColor       =   &H0080FFFF&
      Height          =   555
      Left            =   2550
      Picture         =   "frmPesquisaProduto2.frx":7F014
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3630
      Width           =   1200
   End
   Begin VB.CommandButton cmd_acimaCorrecel2 
      BackColor       =   &H0080FFFF&
      Height          =   555
      Left            =   2550
      Picture         =   "frmPesquisaProduto2.frx":81F96
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmd_abaixoCorrecel1 
      BackColor       =   &H00C0FFFF&
      Height          =   555
      Left            =   30
      Picture         =   "frmPesquisaProduto2.frx":84728
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3630
      Width           =   2445
   End
   Begin VB.CommandButton cmd_acimaCorrecel1 
      BackColor       =   &H00C0FFFF&
      Height          =   555
      Left            =   30
      Picture         =   "frmPesquisaProduto2.frx":89472
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   60
      Width           =   2445
   End
   Begin VB.CommandButton cmd_carrocel2_04 
      BackColor       =   &H0080FFFF&
      Height          =   720
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel2_03 
      BackColor       =   &H0080FFFF&
      Height          =   720
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2130
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel2_02 
      BackColor       =   &H0080FFFF&
      Height          =   720
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1380
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel2_01 
      BackColor       =   &H0080FFFF&
      Height          =   720
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   630
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_08 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_07 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2130
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_06 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1380
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_05 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   630
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_04 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_03 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2130
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_02 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1380
      Width           =   1200
   End
   Begin VB.CommandButton cmd_carrocel_01 
      BackColor       =   &H00C0FFFF&
      Height          =   720
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   630
      Width           =   1200
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
      Height          =   1515
      Left            =   11880
      TabIndex        =   4
      Top             =   1860
      Width           =   3135
      Begin VB.CommandButton cmd_menos 
         BackColor       =   &H00C0C0FF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmd_mais 
         BackColor       =   &H00FFFF80&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txt_EntrarQtde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   90
         TabIndex        =   5
         Text            =   "1"
         Top             =   960
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmd_incluir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3465
      Width           =   3135
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultados 
      Height          =   4110
      Left            =   3825
      TabIndex        =   1
      Top             =   60
      Width           =   8010
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
      Col.Count       =   2
      BevelColorHighlight=   16777215
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
      MaxSelectedRows =   0
      ForeColorEven   =   4210752
      BackColorEven   =   12648447
      BackColorOdd    =   16777215
      RowHeight       =   503
      ExtraHeight     =   212
      Columns.Count   =   2
      Columns(0).Width=   3757
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   9313
      Columns(1).Caption=   "Descrição"
      Columns(1).Name =   "Descricao"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   14129
      _ExtentY        =   7250
      _StockProps     =   79
      BackColor       =   15066597
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_codProduto 
      BackColor       =   &H00E5E5E5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   11880
      TabIndex        =   3
      Top             =   90
      Width           =   1845
   End
   Begin VB.Label lbl_nomeProduto 
      BackColor       =   &H00E5E5E5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   11880
      TabIndex        =   2
      Top             =   630
      Width           =   3135
   End
End
Attribute VB_Name = "frmPesquisaProduto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private contador_arrayClasses As Long
Private contador_arraySubClasses As Long
Private ponteiro_arrayClasses As Long
Private ponteiro_arraySubClasses As Long
Dim arrayClasses() As Variant
Dim arraySubClasses() As Variant
Private arrayProdutos() As Variant
Private contador_arrayProdutos As Long

Private lCodigoClasse_SELECIONADA As Long
Private sNomeClasse_SELECIONADA As String
Private lCodigoSubClasse_SELECIONADA As Long
Private sNomeSubClasse_SELECIONADA As String

Private sProdutoFavoritos1 As String
Private sProdutoFavoritos2 As String
Private sProdutoFavoritos3 As String
Private sProdutoFavoritos4 As String
Private sProdutoFavoritos5 As String
Private sProdutoFavoritos6 As String
Private sProdutoFavoritos7 As String
Private sProdutoFavoritos8 As String


Private Sub cmd_abaixo_Click()
    Me.Top = Me.Top + 100
End Sub

Private Sub cmd_abaixoCorrecel1_Click()
On Error GoTo Erro
  Dim lContadorLocal As Long
  Dim i As Integer
  
  
  lContadorLocal = contador_arrayClasses - (ponteiro_arrayClasses * 8)
  
  ' ===============================================
  ' Tratar a pintura do botão clicado
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
                      
  If lContadorLocal > 0 Then
      If lContadorLocal > 8 Then
          For i = 1 To 8
              If lCodigoClasse_SELECIONADA = arrayClasses((ponteiro_arrayClasses * 8) + (i - 1), 0) Then
                  If i = 1 Then
                      cmd_carrocel_01.BackColor = &HFFA324
                  ElseIf i = 2 Then
                      cmd_carrocel_02.BackColor = &HFFA324
                  ElseIf i = 3 Then
                      cmd_carrocel_03.BackColor = &HFFA324
                  ElseIf i = 4 Then
                      cmd_carrocel_04.BackColor = &HFFA324
                  ElseIf i = 5 Then
                      cmd_carrocel_05.BackColor = &HFFA324
                  ElseIf i = 6 Then
                      cmd_carrocel_06.BackColor = &HFFA324
                  ElseIf i = 7 Then
                      cmd_carrocel_07.BackColor = &HFFA324
                  ElseIf i = 8 Then
                      cmd_carrocel_08.BackColor = &HFFA324
                  End If
              End If
          Next
      Else
          For i = 1 To lContadorLocal
              If lCodigoClasse_SELECIONADA = arrayClasses((ponteiro_arrayClasses * 8) + (i - 1), 0) Then
                  If i = 1 Then
                      cmd_carrocel_01.BackColor = &HFFA324
                  ElseIf i = 2 Then
                      cmd_carrocel_02.BackColor = &HFFA324
                  ElseIf i = 3 Then
                      cmd_carrocel_03.BackColor = &HFFA324
                  ElseIf i = 4 Then
                      cmd_carrocel_04.BackColor = &HFFA324
                  ElseIf i = 5 Then
                      cmd_carrocel_05.BackColor = &HFFA324
                  ElseIf i = 6 Then
                      cmd_carrocel_06.BackColor = &HFFA324
                  ElseIf i = 7 Then
                      cmd_carrocel_07.BackColor = &HFFA324
                  ElseIf i = 8 Then
                      cmd_carrocel_08.BackColor = &HFFA324
                  End If
              End If
          Next
      End If
  End If
  '
  ' ===============================================
  
  
  If lContadorLocal > 0 Then
      If lContadorLocal = 1 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = ""
          cmd_carrocel_03.Caption = ""
          cmd_carrocel_04.Caption = ""
          cmd_carrocel_05.Caption = ""
          cmd_carrocel_06.Caption = ""
          cmd_carrocel_07.Caption = ""
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal = 2 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = ""
          cmd_carrocel_04.Caption = ""
          cmd_carrocel_05.Caption = ""
          cmd_carrocel_06.Caption = ""
          cmd_carrocel_07.Caption = ""
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal = 3 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 2, 1)
          cmd_carrocel_04.Caption = ""
          cmd_carrocel_05.Caption = ""
          cmd_carrocel_06.Caption = ""
          cmd_carrocel_07.Caption = ""
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal = 4 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 2, 1)
          cmd_carrocel_04.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 3, 1)
          cmd_carrocel_05.Caption = ""
          cmd_carrocel_06.Caption = ""
          cmd_carrocel_07.Caption = ""
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal = 5 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 2, 1)
          cmd_carrocel_04.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 3, 1)
          cmd_carrocel_05.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 4, 1)
          cmd_carrocel_06.Caption = ""
          cmd_carrocel_07.Caption = ""
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal = 6 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 2, 1)
          cmd_carrocel_04.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 3, 1)
          cmd_carrocel_05.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 4, 1)
          cmd_carrocel_06.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 5, 1)
          cmd_carrocel_07.Caption = ""
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal = 7 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 2, 1)
          cmd_carrocel_04.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 3, 1)
          cmd_carrocel_05.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 4, 1)
          cmd_carrocel_06.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 5, 1)
          cmd_carrocel_07.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 6, 1)
          cmd_carrocel_08.Caption = ""
      ElseIf lContadorLocal >= 8 Then
          cmd_carrocel_01.Caption = arrayClasses((ponteiro_arrayClasses * 8), 1)
          cmd_carrocel_02.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 1, 1)
          cmd_carrocel_03.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 2, 1)
          cmd_carrocel_04.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 3, 1)
          cmd_carrocel_05.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 4, 1)
          cmd_carrocel_06.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 5, 1)
          cmd_carrocel_07.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 6, 1)
          cmd_carrocel_08.Caption = arrayClasses((ponteiro_arrayClasses * 8) + 7, 1)
      End If
          
      ponteiro_arrayClasses = ponteiro_arrayClasses + 1

      cmd_acimaCorrecel1.Enabled = True
      
      If lContadorLocal <= 8 Then
          cmd_abaixoCorrecel1.Enabled = False
      Else
          cmd_abaixoCorrecel1.Enabled = True
      End If
  End If
  
  Exit Sub
  
Erro:
  MsgBox "Erro " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_abaixoCorrecel2_Click()
On Error GoTo Erro
  Dim lContadorLocal As Long
  Dim i As Integer
  
  lContadorLocal = contador_arraySubClasses - (ponteiro_arraySubClasses * 4)
  
  ' ===============================================
  ' Tratar a pintura do botão clicado
  cmd_carrocel2_01.BackColor = &HC0FFFF
  cmd_carrocel2_02.BackColor = &HC0FFFF
  cmd_carrocel2_03.BackColor = &HC0FFFF
  cmd_carrocel2_04.BackColor = &HC0FFFF
                      
  If lContadorLocal > 0 Then
      If lContadorLocal > 4 Then
          For i = 1 To 4
              If lCodigoSubClasse_SELECIONADA = arraySubClasses((ponteiro_arraySubClasses * 4) + (i - 1), 0) Then
                  If i = 1 Then
                      cmd_carrocel2_01.BackColor = &HFFA324
                  ElseIf i = 2 Then
                      cmd_carrocel2_02.BackColor = &HFFA324
                  ElseIf i = 3 Then
                      cmd_carrocel2_03.BackColor = &HFFA324
                  ElseIf i = 4 Then
                      cmd_carrocel2_04.BackColor = &HFFA324
                  End If
              End If
          Next
      Else
          For i = 1 To lContadorLocal
              If lCodigoSubClasse_SELECIONADA = arraySubClasses((ponteiro_arraySubClasses * 4) + (i - 1), 0) Then
                  If i = 1 Then
                      cmd_carrocel2_01.BackColor = &HFFA324
                  ElseIf i = 2 Then
                      cmd_carrocel2_02.BackColor = &HFFA324
                  ElseIf i = 3 Then
                      cmd_carrocel2_03.BackColor = &HFFA324
                  ElseIf i = 4 Then
                      cmd_carrocel2_04.BackColor = &HFFA324
                  End If
              End If
          Next
      End If
  End If
  '
  ' ===============================================
  
  
  If lContadorLocal > 0 Then
      If lContadorLocal = 1 Then
          cmd_carrocel2_01.Caption = arraySubClasses((ponteiro_arraySubClasses * 4), 1)
          cmd_carrocel2_02.Caption = ""
          cmd_carrocel2_03.Caption = ""
          cmd_carrocel2_04.Caption = ""
      ElseIf lContadorLocal = 2 Then
          cmd_carrocel2_01.Caption = arraySubClasses((ponteiro_arraySubClasses * 4), 1)
          cmd_carrocel2_02.Caption = arraySubClasses((ponteiro_arraySubClasses * 4) + 1, 1)
          cmd_carrocel2_03.Caption = ""
          cmd_carrocel2_04.Caption = ""
      ElseIf lContadorLocal = 3 Then
          cmd_carrocel2_01.Caption = arraySubClasses((ponteiro_arraySubClasses * 4), 1)
          cmd_carrocel2_02.Caption = arraySubClasses((ponteiro_arraySubClasses * 4) + 1, 1)
          cmd_carrocel2_03.Caption = arraySubClasses((ponteiro_arraySubClasses * 4) + 2, 1)
          cmd_carrocel2_04.Caption = ""
      ElseIf lContadorLocal >= 4 Then
          cmd_carrocel2_01.Caption = arraySubClasses((ponteiro_arraySubClasses * 4), 1)
          cmd_carrocel2_02.Caption = arraySubClasses((ponteiro_arraySubClasses * 4) + 1, 1)
          cmd_carrocel2_03.Caption = arraySubClasses((ponteiro_arraySubClasses * 4) + 2, 1)
          cmd_carrocel2_04.Caption = arraySubClasses((ponteiro_arraySubClasses * 4) + 3, 1)
      End If
          
      ponteiro_arraySubClasses = ponteiro_arraySubClasses + 1

      cmd_acimaCorrecel2.Enabled = True
      
      If lContadorLocal <= 4 Then
          cmd_abaixoCorrecel2.Enabled = False
      Else
          cmd_abaixoCorrecel2.Enabled = True
      End If
  End If
  
  Exit Sub
  
Erro:
  MsgBox "Erro " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_Acima_Click()
    Me.Top = Me.Top - 100
End Sub

Private Sub cmd_acimaCorrecel1_Click()
On Error GoTo Erro
  Dim lContadorLocal As Long
  Dim i As Integer
  
  ponteiro_arrayClasses = ponteiro_arrayClasses - 1
  
  lContadorLocal = ponteiro_arrayClasses - 1
  
  
  ' ===============================================
  ' Tratar a pintura do botão clicado
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
                      
  If lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8), 0) Then
      cmd_carrocel_01.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 1, 0) Then
      cmd_carrocel_02.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 2, 0) Then
      cmd_carrocel_03.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 3, 0) Then
      cmd_carrocel_04.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 4, 0) Then
      cmd_carrocel_05.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 5, 0) Then
      cmd_carrocel_06.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 6, 0) Then
      cmd_carrocel_07.BackColor = &HFFA324
  ElseIf lCodigoClasse_SELECIONADA = arrayClasses((lContadorLocal * 8) + 7, 0) Then
      cmd_carrocel_08.BackColor = &HFFA324
  End If
  '
  ' ===============================================
  
  cmd_carrocel_01.Caption = arrayClasses((lContadorLocal * 8), 1)
  cmd_carrocel_02.Caption = arrayClasses((lContadorLocal * 8) + 1, 1)
  cmd_carrocel_03.Caption = arrayClasses((lContadorLocal * 8) + 2, 1)
  cmd_carrocel_04.Caption = arrayClasses((lContadorLocal * 8) + 3, 1)
  cmd_carrocel_05.Caption = arrayClasses((lContadorLocal * 8) + 4, 1)
  cmd_carrocel_06.Caption = arrayClasses((lContadorLocal * 8) + 5, 1)
  cmd_carrocel_07.Caption = arrayClasses((lContadorLocal * 8) + 6, 1)
  cmd_carrocel_08.Caption = arrayClasses((lContadorLocal * 8) + 7, 1)
  
  If ponteiro_arrayClasses = 1 Then
      cmd_acimaCorrecel1.Enabled = False
  End If
  
  cmd_abaixoCorrecel1.Enabled = True
  
  Exit Sub
  
Erro:
  MsgBox "Erro " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmd_acimaCorrecel2_Click()
On Error GoTo Erro
  Dim lContadorLocal As Long
  
  ponteiro_arraySubClasses = ponteiro_arraySubClasses - 1
  
  lContadorLocal = ponteiro_arraySubClasses - 1
  
  ' ===============================================
  ' Tratar a pintura do botão clicado
  cmd_carrocel2_01.BackColor = &HC0FFFF
  cmd_carrocel2_02.BackColor = &HC0FFFF
  cmd_carrocel2_03.BackColor = &HC0FFFF
  cmd_carrocel2_04.BackColor = &HC0FFFF
                      
  If lCodigoSubClasse_SELECIONADA = arraySubClasses((lContadorLocal * 4), 0) Then
      cmd_carrocel2_01.BackColor = &HFFA324
  ElseIf lCodigoSubClasse_SELECIONADA = arraySubClasses((lContadorLocal * 4) + 1, 0) Then
      cmd_carrocel2_02.BackColor = &HFFA324
  ElseIf lCodigoSubClasse_SELECIONADA = arraySubClasses((lContadorLocal * 4) + 2, 0) Then
      cmd_carrocel2_03.BackColor = &HFFA324
  ElseIf lCodigoSubClasse_SELECIONADA = arraySubClasses((lContadorLocal * 4) + 3, 0) Then
      cmd_carrocel2_04.BackColor = &HFFA324
  End If
  '
  ' ===============================================
  
  
  cmd_carrocel2_01.Caption = arraySubClasses((lContadorLocal * 4), 1)
  cmd_carrocel2_02.Caption = arraySubClasses((lContadorLocal * 4) + 1, 1)
  cmd_carrocel2_03.Caption = arraySubClasses((lContadorLocal * 4) + 2, 1)
  cmd_carrocel2_04.Caption = arraySubClasses((lContadorLocal * 4) + 3, 1)
  
  If ponteiro_arraySubClasses = 1 Then
      cmd_acimaCorrecel2.Enabled = False
  End If
  
  cmd_abaixoCorrecel2.Enabled = True
  
  
  Exit Sub
  
Erro:
  MsgBox "Erro " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub LimpaProdutosFavoritos()
    cmd_favoritos1.BackColor = &HC0FFFF
    cmd_favoritos2.BackColor = &HC0FFFF
    cmd_favoritos3.BackColor = &HC0FFFF
    cmd_favoritos4.BackColor = &HC0FFFF
    cmd_favoritos5.BackColor = &HC0FFFF
    cmd_favoritos6.BackColor = &HC0FFFF
    cmd_favoritos7.BackColor = &HC0FFFF
    cmd_favoritos8.BackColor = &HC0FFFF
End Sub

Private Sub cmd_carrocel_01_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_01.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_01.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_01.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      
      cmd_carrocel_01.BackColor = &HFFA324
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmd_carrocel_02_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_02.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_02.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_02.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_02.BackColor = &HFFA324
  
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel_03_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_03.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_03.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_03.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_03.BackColor = &HFFA324
  
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel_04_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_04.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_04.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_04.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_04.BackColor = &HFFA324
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel_05_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_05.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_05.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              
              End If
          Next
      End If
      cmd_carrocel_05.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_05.BackColor = &HFFA324
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel_06_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_06.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_06.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5))
              End If
          Next
      End If
      cmd_carrocel_06.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_06.BackColor = &HFFA324
  
  End If
  grdResultados.Redraw = True
  
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel_07_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_07.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_07.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_07.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_07.BackColor = &HFFA324
  
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel_08_Click()
On Error GoTo Erro
  Dim lContaReg As Long
  Dim lCodigoClasse As Long
  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arrayClasses - 1
      If arrayClasses(lContaReg, 1) = cmd_carrocel_08.Caption Then
          lCodigoClasse_SELECIONADA = arrayClasses(lContaReg, 0)
          sNomeClasse_SELECIONADA = arrayClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel_08.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoClasse_SELECIONADA = 0
      
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_08.BackColor = &HC0FFFF
  Else
      If lCodigoSubClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel_08.BackColor = &HFFA324
  
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel2_01_Click()
On Error GoTo Erro
  Dim lContaReg As Long

  
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel2_02.BackColor = &HC0FFFF
  cmd_carrocel2_03.BackColor = &HC0FFFF
  cmd_carrocel2_04.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoSubClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arraySubClasses - 1
      If arraySubClasses(lContaReg, 1) = cmd_carrocel2_01.Caption Then
          lCodigoSubClasse_SELECIONADA = arraySubClasses(lContaReg, 0)
          sNomeSubClasse_SELECIONADA = arraySubClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel2_01.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoSubClasse_SELECIONADA = 0
      
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_01.BackColor = &HC0FFFF
  Else
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_01.BackColor = &HFFA324
  End If
  
  grdResultados.Redraw = True
  
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel2_02_Click()
On Error GoTo Erro
  Dim lContaReg As Long

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  LimpaProdutosFavoritos
  cmd_carrocel2_01.BackColor = &HC0FFFF
  cmd_carrocel2_03.BackColor = &HC0FFFF
  cmd_carrocel2_04.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoSubClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arraySubClasses - 1
      If arraySubClasses(lContaReg, 1) = cmd_carrocel2_02.Caption Then
          lCodigoSubClasse_SELECIONADA = arraySubClasses(lContaReg, 0)
          sNomeSubClasse_SELECIONADA = arraySubClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel2_02.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoSubClasse_SELECIONADA = 0
      
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_02.BackColor = &HC0FFFF
  Else
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_02.BackColor = &HFFA324
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel2_03_Click()
On Error GoTo Erro
  Dim lContaReg As Long

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"

  LimpaProdutosFavoritos
  cmd_carrocel2_01.BackColor = &HC0FFFF
  cmd_carrocel2_02.BackColor = &HC0FFFF
  cmd_carrocel2_04.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoSubClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arraySubClasses - 1
      If arraySubClasses(lContaReg, 1) = cmd_carrocel2_03.Caption Then
          lCodigoSubClasse_SELECIONADA = arraySubClasses(lContaReg, 0)
          sNomeSubClasse_SELECIONADA = arraySubClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel2_03.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoSubClasse_SELECIONADA = 0
      
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_03.BackColor = &HC0FFFF
  Else
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_03.BackColor = &HFFA324
  
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_carrocel2_04_Click()
On Error GoTo Erro
  Dim lContaReg As Long

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"

  LimpaProdutosFavoritos
  cmd_carrocel2_01.BackColor = &HC0FFFF
  cmd_carrocel2_02.BackColor = &HC0FFFF
  cmd_carrocel2_03.BackColor = &HC0FFFF
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  DoEvents
  
  ' ===================================================
  ' Achar código da Classe
  lCodigoSubClasse_SELECIONADA = 0
  For lContaReg = 0 To contador_arraySubClasses - 1
      If arraySubClasses(lContaReg, 1) = cmd_carrocel2_04.Caption Then
          lCodigoSubClasse_SELECIONADA = arraySubClasses(lContaReg, 0)
          sNomeSubClasse_SELECIONADA = arraySubClasses(lContaReg, 1)
      End If
  Next
  ' ===================================================

  grdResultados.RowHeight = 530
  
  If cmd_carrocel2_04.BackColor = &HFFA324 Then
      ' Já estava selecionado...então será DESELECIONADO
      lCodigoSubClasse_SELECIONADA = 0
      
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_04.BackColor = &HC0FFFF
  Else
      If lCodigoClasse_SELECIONADA <> 0 Then
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 2) = lCodigoClasse_SELECIONADA And arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      Else
          For lContaReg = 0 To contador_arrayProdutos - 1
              If arrayProdutos(lContaReg, 4) = lCodigoSubClasse_SELECIONADA Then
                  grdResultados.AddItem arrayProdutos(lContaReg, 0) & vbTab & _
                  LCase(arrayProdutos(lContaReg, 1) & " (" & arrayProdutos(lContaReg, 3) & " " & arrayProdutos(lContaReg, 5)) & ")"
              End If
          Next
      End If
      cmd_carrocel2_04.BackColor = &HFFA324
  End If
  
  grdResultados.Redraw = True
  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
    MsgBox "Erro na pesquisa de produtos " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_direita_Click()
    Me.Left = Me.Left + 100
End Sub

Private Sub cmd_esquerda_Click()
    Me.Left = Me.Left - 100
End Sub

Private Sub cmd_favoritos1_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos1.BackColor = &HC0FFFF Then
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos1.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos1
      lbl_nomeProduto.Caption = cmd_favoritos1.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos1.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos2_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos2.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos2.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos2
      lbl_nomeProduto.Caption = cmd_favoritos2.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos2.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos3_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos3.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos3.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos3
      lbl_nomeProduto.Caption = cmd_favoritos3.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos3.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos4_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos4.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos4.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos4
      lbl_nomeProduto.Caption = cmd_favoritos4.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos4.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos5_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos5.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos5.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos5
      lbl_nomeProduto.Caption = cmd_favoritos5.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos5.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos6_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos6.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos6.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos6
      lbl_nomeProduto.Caption = cmd_favoritos6.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos6.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos7_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos7.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos8.BackColor = &HC0FFFF
      
      cmd_favoritos7.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos7
      lbl_nomeProduto.Caption = cmd_favoritos7.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos7.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_favoritos8_Click()
On Error GoTo Erro

  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  txt_EntrarQtde.Text = "1"
  
  If cmd_favoritos8.BackColor = &HC0FFFF Then
      cmd_favoritos1.BackColor = &HC0FFFF
      cmd_favoritos2.BackColor = &HC0FFFF
      cmd_favoritos3.BackColor = &HC0FFFF
      cmd_favoritos4.BackColor = &HC0FFFF
      cmd_favoritos5.BackColor = &HC0FFFF
      cmd_favoritos6.BackColor = &HC0FFFF
      cmd_favoritos7.BackColor = &HC0FFFF
      
      cmd_favoritos8.BackColor = &HFFA324
      
      lbl_codProduto.Caption = sProdutoFavoritos8
      lbl_nomeProduto.Caption = cmd_favoritos8.Caption
      txt_EntrarQtde.SetFocus
  Else
      cmd_favoritos8.BackColor = &HC0FFFF
  End If

  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_fecharTela_Click()
    Unload Me
End Sub

Private Sub cmd_incluir_Click()
On Error GoTo Erro

  Dim intCancel As Integer
  Dim frmX As Form
  
  txt_EntrarQtde.Text = Replace(txt_EntrarQtde.Text, ".", ",")
  
  If Not IsNumeric(txt_EntrarQtde.Text) Then
      MsgBox "Informe uma quantidade válida.", vbInformation, "Atenção"
      txt_EntrarQtde.SetFocus
      Exit Sub
  End If
  
  If Trim(lbl_codProduto.Caption) = "" Then
      MsgBox "Selecione um produto na grade.", vbInformation, "Atenção"
      Exit Sub
  End If
  
  'Form de origem (chamadas comuns)
  Set frmX = g_frmVendaRapida
      
  With frmX
      'Insere o item
      .Grade1.Columns(0).Text = lbl_codProduto.Caption
        
      If txt_EntrarQtde.Text = "" Or txt_EntrarQtde.Text = "0" Then
          .Grade1.Columns(1).Text = "1"
      Else
          .Grade1.Columns(1).Text = txt_EntrarQtde.Text
      End If
        
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
  
  
  '===============================================================================
  ' Limpar Classe e SubClasse selecionadas, mas manter a grade carregada
  lCodigoClasse_SELECIONADA = 0
  lCodigoSubClasse_SELECIONADA = 0
  cmd_carrocel_01.BackColor = &HC0FFFF
  cmd_carrocel_02.BackColor = &HC0FFFF
  cmd_carrocel_03.BackColor = &HC0FFFF
  cmd_carrocel_04.BackColor = &HC0FFFF
  cmd_carrocel_05.BackColor = &HC0FFFF
  cmd_carrocel_06.BackColor = &HC0FFFF
  cmd_carrocel_07.BackColor = &HC0FFFF
  cmd_carrocel_08.BackColor = &HC0FFFF
  
  cmd_carrocel2_01.BackColor = &HC0FFFF
  cmd_carrocel2_02.BackColor = &HC0FFFF
  cmd_carrocel2_03.BackColor = &HC0FFFF
  cmd_carrocel2_04.BackColor = &HC0FFFF
  
  ' Limpar tb os botões de produtos favoritos
  cmd_favoritos1.BackColor = &HC0FFFF
  cmd_favoritos2.BackColor = &HC0FFFF
  cmd_favoritos3.BackColor = &HC0FFFF
  cmd_favoritos4.BackColor = &HC0FFFF
  cmd_favoritos5.BackColor = &HC0FFFF
  cmd_favoritos6.BackColor = &HC0FFFF
  cmd_favoritos7.BackColor = &HC0FFFF
  cmd_favoritos8.BackColor = &HC0FFFF
  
  
  txt_EntrarQtde.Text = "1"
  lbl_codProduto.Caption = ""
  lbl_nomeProduto.Caption = ""
  '===============================================================================
  
      
  Exit Sub

Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

End Sub

Private Sub cmd_mais_Click()
On Error GoTo Erro
    Dim dbl_unidades As Double
    
    If Trim(txt_EntrarQtde.Text) = "" Then
        txt_EntrarQtde.Text = "0"
    End If
    
    If Not IsNumeric(txt_EntrarQtde.Text) Then
        txt_EntrarQtde.Text = "0"
    End If
    
    dbl_unidades = CDbl(txt_EntrarQtde.Text)
    
    If dbl_unidades < 0 Then
        txt_EntrarQtde.Text = "1"
        Exit Sub
    End If
    
    dbl_unidades = dbl_unidades + 1
    txt_EntrarQtde.Text = CStr(dbl_unidades)

    Exit Sub

Erro:
  MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    
End Sub

Private Sub cmd_menos_Click()
On Error GoTo Erro
    Dim dbl_unidades As Double
    
    If Trim(txt_EntrarQtde.Text) = "" Then
        txt_EntrarQtde.Text = "0"
    End If
    
    If Not IsNumeric(txt_EntrarQtde.Text) Then
        txt_EntrarQtde.Text = "0"
    End If
    
    dbl_unidades = CDbl(txt_EntrarQtde.Text)
    
    If dbl_unidades > 0 Then
        dbl_unidades = dbl_unidades - 1
        txt_EntrarQtde.Text = CStr(dbl_unidades)
    End If
    
    If dbl_unidades < 0 Then
        txt_EntrarQtde.Text = "0"
        Exit Sub
    End If

    Exit Sub

Erro:
  MsgBox "Erro " & Err.Description, vbInformation, "Atenção"
    

End Sub

Private Sub Form_Load()
On Error GoTo Erro
  Dim lContador As Long
  Dim rsClasse As Recordset
  Dim rsSubClasse As Recordset
  Dim rsProdutos As Recordset
  Dim rsProdutosFavoritos As Recordset
  Dim sSql As String
  
 
  ponteiro_arrayClasses = 1
  ponteiro_arraySubClasses = 1
  
  lCodigoClasse_SELECIONADA = 0
  lCodigoSubClasse_SELECIONADA = 0

  lContador = 0
  Set rsClasse = db.OpenRecordset("select Código, Nome from Classes order by Nome", dbOpenDynaset)
  If Not (rsClasse.EOF And rsClasse.BOF) Then
      rsClasse.MoveLast
      rsClasse.MoveFirst
      
      ReDim arrayClasses(rsClasse.RecordCount, 2)
      contador_arrayClasses = rsClasse.RecordCount
      While Not rsClasse.EOF
          arrayClasses(lContador, 0) = rsClasse.Fields(0).Value
          If Not IsNull(rsClasse.Fields(1).Value) Then
              arrayClasses(lContador, 1) = rsClasse.Fields(1).Value
          Else
              arrayClasses(lContador, 1) = ""
          End If
          
          lContador = lContador + 1
          rsClasse.MoveNext
      Wend
  End If
  rsClasse.Close
  Set rsClasse = Nothing


  lContador = 0
  Set rsSubClasse = db.OpenRecordset("select Código, Nome from [Sub Classes] order by Nome", dbOpenDynaset)
  If Not (rsSubClasse.EOF And rsSubClasse.BOF) Then
      rsSubClasse.MoveLast
      rsSubClasse.MoveFirst
      
      ReDim arraySubClasses(rsSubClasse.RecordCount, 2)
      contador_arraySubClasses = rsSubClasse.RecordCount
      While Not rsSubClasse.EOF
          arraySubClasses(lContador, 0) = rsSubClasse.Fields(0).Value
          If Not IsNull(rsSubClasse.Fields(1).Value) Then
              arraySubClasses(lContador, 1) = rsSubClasse.Fields(1).Value
          Else
              arraySubClasses(lContador, 1) = ""
          End If
          lContador = lContador + 1
          rsSubClasse.MoveNext
      Wend
  End If
  rsSubClasse.Close
  Set rsSubClasse = Nothing
  
  lContador = 0
  sSql = "select P.Código, P.Nome, P.Classe, C.Nome, P.[Sub Classe], S.Nome "
  sSql = sSql & " from Produtos P, Classes C, [Sub Classes] S "
  sSql = sSql & " Where P.Classe = C.Código and P.[Sub Classe] = S.Código and P.Desativado = false "
  sSql = sSql & " order by P.Nome"
  
  Set rsProdutos = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If Not (rsProdutos.EOF And rsProdutos.BOF) Then
      rsProdutos.MoveLast
      rsProdutos.MoveFirst
      
      ReDim arrayProdutos(rsProdutos.RecordCount, 6)
      contador_arrayProdutos = rsProdutos.RecordCount
      While Not rsProdutos.EOF
          arrayProdutos(lContador, 0) = rsProdutos.Fields(0).Value
          arrayProdutos(lContador, 1) = rsProdutos.Fields(1).Value
          arrayProdutos(lContador, 2) = rsProdutos.Fields(2).Value
          arrayProdutos(lContador, 3) = rsProdutos.Fields(3).Value
          arrayProdutos(lContador, 4) = rsProdutos.Fields(4).Value
          arrayProdutos(lContador, 5) = rsProdutos.Fields(5).Value
          lContador = lContador + 1
          rsProdutos.MoveNext
      Wend
  End If
  rsProdutos.Close
  Set rsProdutos = Nothing
  
  
  ' ==========================================================================
  ' Carregar carrocel 01 CLASSES
  If contador_arrayClasses > 0 Then
      If contador_arrayClasses = 1 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
      ElseIf contador_arrayClasses = 2 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
      ElseIf contador_arrayClasses = 3 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
          cmd_carrocel_03.Caption = arrayClasses(2, 1)
      ElseIf contador_arrayClasses = 4 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
          cmd_carrocel_03.Caption = arrayClasses(2, 1)
          cmd_carrocel_04.Caption = arrayClasses(3, 1)
      ElseIf contador_arrayClasses = 5 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
          cmd_carrocel_03.Caption = arrayClasses(2, 1)
          cmd_carrocel_04.Caption = arrayClasses(3, 1)
          cmd_carrocel_05.Caption = arrayClasses(4, 1)
      ElseIf contador_arrayClasses = 6 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
          cmd_carrocel_03.Caption = arrayClasses(2, 1)
          cmd_carrocel_04.Caption = arrayClasses(3, 1)
          cmd_carrocel_05.Caption = arrayClasses(4, 1)
          cmd_carrocel_06.Caption = arrayClasses(5, 1)
      ElseIf contador_arrayClasses = 7 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
          cmd_carrocel_03.Caption = arrayClasses(2, 1)
          cmd_carrocel_04.Caption = arrayClasses(3, 1)
          cmd_carrocel_05.Caption = arrayClasses(4, 1)
          cmd_carrocel_06.Caption = arrayClasses(5, 1)
          cmd_carrocel_07.Caption = arrayClasses(6, 1)
      ElseIf contador_arrayClasses >= 8 Then
          cmd_carrocel_01.Caption = arrayClasses(0, 1)
          cmd_carrocel_02.Caption = arrayClasses(1, 1)
          cmd_carrocel_03.Caption = arrayClasses(2, 1)
          cmd_carrocel_04.Caption = arrayClasses(3, 1)
          cmd_carrocel_05.Caption = arrayClasses(4, 1)
          cmd_carrocel_06.Caption = arrayClasses(5, 1)
          cmd_carrocel_07.Caption = arrayClasses(6, 1)
          cmd_carrocel_08.Caption = arrayClasses(7, 1)
      End If
      
      If contador_arrayClasses <= 8 Then
          cmd_acimaCorrecel1.Enabled = False
          cmd_abaixoCorrecel1.Enabled = False
      Else
          cmd_acimaCorrecel1.Enabled = False
          cmd_abaixoCorrecel1.Enabled = True
      End If
  End If
  ' ==========================================================================
  
  ' ==========================================================================
  ' Carregar carrocel 02 SUB CLASSES
  If contador_arraySubClasses > 0 Then
      If contador_arraySubClasses = 1 Then
          cmd_carrocel2_01.Caption = arraySubClasses(0, 1)
      ElseIf contador_arraySubClasses = 2 Then
          cmd_carrocel2_01.Caption = arraySubClasses(0, 1)
          cmd_carrocel2_02.Caption = arraySubClasses(1, 1)
      ElseIf contador_arraySubClasses = 3 Then
          cmd_carrocel2_01.Caption = arraySubClasses(0, 1)
          cmd_carrocel2_02.Caption = arraySubClasses(1, 1)
          cmd_carrocel2_03.Caption = arraySubClasses(2, 1)
      ElseIf contador_arraySubClasses >= 4 Then
          cmd_carrocel2_01.Caption = arraySubClasses(0, 1)
          cmd_carrocel2_02.Caption = arraySubClasses(1, 1)
          cmd_carrocel2_03.Caption = arraySubClasses(2, 1)
          cmd_carrocel2_04.Caption = arraySubClasses(3, 1)
      End If
      
      If contador_arraySubClasses <= 8 Then
          cmd_acimaCorrecel2.Enabled = False
          cmd_abaixoCorrecel2.Enabled = False
      Else
          cmd_acimaCorrecel2.Enabled = False
          cmd_abaixoCorrecel2.Enabled = True
      End If
  End If
  ' ==========================================================================

  ' ==========================================================================
  sSql = "Select F.Produto, P.Nome from ProdutoFavoritos F, Produtos P "
  sSql = sSql & " where F.Filial = " & gnCodFilial & " and F.Produto = P.Código "
  
  Set rsProdutosFavoritos = db.OpenRecordset(sSql, dbOpenDynaset)
  If Not (rsProdutosFavoritos.EOF And rsProdutosFavoritos.BOF) Then
      rsProdutosFavoritos.MoveLast
      rsProdutosFavoritos.MoveFirst
      
      If rsProdutosFavoritos.RecordCount > 0 Then
          If rsProdutosFavoritos.RecordCount = 1 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 2 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 3 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos3 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos3.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 4 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos3 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos4 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos4.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 5 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos3 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos4 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos5 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos5.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 6 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos3 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos4 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos5 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos5.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos6 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos6.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 7 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos3 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos4 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos5 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos5.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos6 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos6.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos7 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos7.Caption = rsProdutosFavoritos.Fields(1).Value
          ElseIf rsProdutosFavoritos.RecordCount = 8 Then
              sProdutoFavoritos1 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos1.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos2 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos2.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos3 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos3.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos4 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos4.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos5 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos5.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos6 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos6.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos7 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos7.Caption = rsProdutosFavoritos.Fields(1).Value
              rsProdutosFavoritos.MoveNext
              sProdutoFavoritos8 = rsProdutosFavoritos.Fields(0).Value
              cmd_favoritos8.Caption = rsProdutosFavoritos.Fields(1).Value
          End If
      End If
  Else
      cmd_favoritos1.Caption = ""
      cmd_favoritos2.Caption = ""
      cmd_favoritos3.Caption = ""
      cmd_favoritos4.Caption = ""
      cmd_favoritos5.Caption = ""
      cmd_favoritos6.Caption = ""
      cmd_favoritos7.Caption = ""
      cmd_favoritos8.Caption = ""
      
      sProdutoFavoritos1 = ""
      sProdutoFavoritos2 = ""
      sProdutoFavoritos3 = ""
      sProdutoFavoritos4 = ""
      sProdutoFavoritos5 = ""
      sProdutoFavoritos6 = ""
      sProdutoFavoritos7 = ""
      sProdutoFavoritos8 = ""
  End If
  rsProdutosFavoritos.Close
  Set rsProdutosFavoritos = Nothing
  ' ==========================================================================




  ' Altura Tela pesquisaProdutosAlfa = 4220
  ' Altura Tela vendaRapidaCheckOut  = 5400
  ' Total de                         = 9620
      
  If Screen.Height > 9620 Then
      Me.Top = (Screen.Height - 9620) / 2
      Me.Left = (Screen.Width - Me.Width) / 2
      Me.Show
  Else
      Me.Top = 300
      Me.Left = (Screen.Width - Me.Width) / 2
      Me.Show
  End If

  Exit Sub
  
Erro:
  MsgBox "Erro na abertura da tela " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Sub grdResultados_Click()
On Error GoTo Erro

  Dim sCodigoProduto As String
  
  txt_EntrarQtde.Text = "1"
  sCodigoProduto = grdResultados.Columns(0).Text
  lbl_codProduto.Caption = sCodigoProduto
  lbl_nomeProduto.Caption = grdResultados.Columns(1).Text
  txt_EntrarQtde.SetFocus
 
  Exit Sub
Erro:
  MsgBox "Erro na função Click da grade " & Err.Description, vbInformation, "Atenção"
End Sub

