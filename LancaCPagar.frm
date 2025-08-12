VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLancaCPagar 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lançamentos/Manutenção de Contas a Pagar"
   ClientHeight    =   5730
   ClientLeft      =   1350
   ClientTop       =   1035
   ClientWidth     =   7860
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
   HelpContextID   =   1360
   Icon            =   "LancaCPagar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   7860
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
      Height          =   345
      Left            =   -210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   5190
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.TextBox Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2400
      Width           =   5880
   End
   Begin VB.TextBox Nota 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1515
      Width           =   1335
   End
   Begin VB.Data Data3 
      Appearance      =   0  'Flat
      Caption         =   "Data3"
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
      Height          =   315
      Left            =   1080
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   5250
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
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
      Height          =   315
      Left            =   2790
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   5220
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSMask.MaskEdBox Valor_Pago 
      Height          =   360
      Left            =   6225
      TabIndex        =   11
      Top             =   4350
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Height          =   360
      Left            =   6225
      TabIndex        =   10
      Top             =   3780
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Height          =   360
      Left            =   6225
      TabIndex        =   9
      Top             =   3322
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   360
      Left            =   6225
      TabIndex        =   8
      Top             =   2872
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Custo 
      Bindings        =   "LancaCPagar.frx":4E95A
      DataSource      =   "Data2"
      Height          =   360
      Left            =   1680
      TabIndex        =   2
      Top             =   1065
      Width           =   1335
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
      Columns(0).Width=   6482
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1826
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   635
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "LancaCPagar.frx":4E96E
      DataSource      =   "Data3"
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   615
      Width           =   1335
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
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1958
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   635
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "LancaCPagar.frx":4E982
      DataSource      =   "Data1"
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   165
      Width           =   1335
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
      Columns(0).Width=   6826
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
      _ExtentY        =   635
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox Data_Pagto 
      Height          =   360
      Left            =   6225
      TabIndex        =   12
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   4785
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   360
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   3322
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      ForeColor       =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin MSMask.MaskEdBox Emissão 
      Height          =   360
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2872
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "* ATENÇÃO: Caso efetue o pagamento nesta tela...o CAIXA NÃO será sensibilizado"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   990
      TabIndex        =   29
      Top             =   4350
      Width           =   3510
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6225
      X2              =   7560
      Y1              =   4215
      Y2              =   4215
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   3600
      Top             =   4905
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "LancaCPagar.frx":4E996
   End
   Begin VB.Label Nome_Custo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3060
      TabIndex        =   28
      Top             =   1065
      Width           =   4530
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Custo"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1118
      Width           =   1335
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pagamento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4665
      TabIndex        =   26
      Top             =   4845
      Width           =   1350
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago : (=)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4665
      TabIndex        =   25
      Top             =   4395
      Width           =   1215
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Acréscimo :  (+)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4665
      TabIndex        =   24
      Top             =   3795
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto :   (-)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4665
      TabIndex        =   23
      Top             =   3375
      Width           =   1215
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4665
      TabIndex        =   22
      Top             =   2925
      Width           =   975
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Seqüência"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2925
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Nota"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2025
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   668
      Width           =   1095
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3045
      TabIndex        =   15
      Top             =   615
      Width           =   4530
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   218
      Width           =   855
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3045
      TabIndex        =   13
      Top             =   165
      Width           =   4530
   End
End
Attribute VB_Name = "frmLancaCPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'05/05/2005 - Daniel
'
'Projeto: Melhorias para o Centro de Custo
'
'A partir da versão 6.52.0.38 todo campo de Centro de Custo
'estará carregando apenas os Centros que estão ativos no sistema

Dim Num_Registro As Variant
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsCP As Recordset
Dim rsCentros As Recordset
Dim Conta As Long

Private gsSql As String
Private gsWhere As String
Private gsOrder As String

Private Sub Acréscimo_GotFocus()
  Acréscimo.SelStart = 0
  Acréscimo.SelLength = 16
End Sub

Private Sub Acréscimo_LostFocus()
'  Valor_Pago.Text = CCur(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text))

' Calcula valores dos camos Desconto, Acréscimo e Valor_Pago (22/06/2022 - Pablo)
Call CalculaValores("Acréscimo")
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
    Case "miOpSearch"
      Call SearchRecord
    Case "miOpPrintCheques"
      Call PrintCheque
  End Select
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  gsOrder = ""
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0 '"Por Filial, Vencimento, Fornecedor, Contador"
          gsOrder = "ORDER BY Filial, Vencimento, Fornecedor, Contador"
        Case 1 '"Por Filial, Fornecedor"
          gsOrder = "ORDER BY Filial, Fornecedor, Vencimento"
        Case 2 '"Por Filial, Data Pagamento"
          gsOrder = "ORDER BY Filial, Pagamento, Fornecedor"
        Case 3 '"Por Filial, Centro de Custo"
          gsOrder = "ORDER BY Filial, [Centro de Custo]"
        Case 4 '"Por Nota, Fornecedor"
          gsOrder = "ORDER BY Nota, Fornecedor"
      End Select
  End Select
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsCP
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsCP
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsCP
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsCP
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Private Sub SearchRecord()

  If Not IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Apague todos os campos da tela com o botão NOVO."
    gsMsg = gsMsg & vbCrLf & "Selecione a Ordem de Pesquisa na lista e preencha com dados iniciais os campos respectivos."
    gsMsg = gsMsg & vbCrLf & "Pressione novamente este botão PROCURAR."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsWhere = ""
  
  If Len(Trim(Combo_Empresa.Text)) = 0 Then
    Combo_Empresa.Text = "0"
  End If
  
  Select Case ActiveBar1.Tools("miOpOrdem").CBListIndex
    
    Case -1, 0  '"Por Filial, Vencimento"
      If Not IsDate(Vencimento.Text) Then
        Vencimento.Text = Date - 3
      End If
      gsWhere = "WHERE Filial >= " & Combo_Empresa.Text & " AND Vencimento >= #" & Format(Vencimento.Text, "mm/dd/yyyy") & "#"
    Case 1  '"Por Filial, Fornecedor"
      If Len(Trim(Combo_Cliente.Text)) = 0 Then
        Combo_Cliente.Text = "0"
      End If
      gsWhere = "WHERE Filial >= " & Combo_Empresa.Text & " AND Fornecedor >= " & Combo_Cliente.Text
    Case 2  '"Por Filial, Data Pagamento"
      If Not IsDate(Data_Pagto.Text) Then
        Data_Pagto.Text = Date - 3
      End If
      gsWhere = "WHERE Filial >= " & Combo_Empresa.Text & " AND Pagamento >= #" & Format(Data_Pagto.Text, "mm/dd/yyyy") & "#"
    Case 3  '"Por Filial, Centro de Custo"
      If Len(Trim(Combo_Custo.Text)) = 0 Then
        Combo_Custo.Text = "0"
      End If
      gsWhere = "WHERE Filial >= " & Combo_Empresa.Text & " AND [Centro de Custo] >= " & Combo_Custo.Text
    Case 4  '"Por Nota, Fornecedor"
      If Len(Trim(Combo_Cliente.Text)) = 0 Then
        Combo_Cliente.Text = "0"
      End If
      gsWhere = "WHERE Nota >= '" & Trim(Nota.Text) & "' AND Fornecedor >= " & Combo_Cliente.Text
  End Select
  
  Set rsCP = db.OpenRecordset(gsSql & " " & gsWhere & " " & gsOrder, dbOpenDynaset)
  If Not rsCP.EOF Then
    Call ShowRecord
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em função dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar esta conta?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    
    '11/07/2007 - Anderson
    'Criação de log para registro de exclusão de registro
    'Efetua registro do Log
    db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
      Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '" & Left("Usu:" & gnUserCode & " Forn:" & rsCP("Fornecedor") & " Seq:" & rsCP("Sequência") & " NF:" & rsCP("Nota") & " Venc:" & rsCP("Vencimento") & " Vr:" & rsCP("Valor"), 80) & "', 'CNT_PAG: excluir')", dbFailOnError

    rsCP.Delete
    rsCP.MovePrevious
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim Conta As Variant
  Dim Aux_Filial As Integer
  Dim Aux_Cliente As Long
  Dim Aux_Vencimento As Variant
  Dim Valor_Correto As Double
  Dim sTexto As String
  
  Dim intRepeatUpdateLocked As Integer
  Dim blnInTransaction As Boolean
  
  On Error GoTo Trata_Erro:
  
  Call StatusMsg("")
  
  Rem Verifica Empresa
  If Nome_Empresa.Caption = "" Then
    DisplayMsg "Filial inválida, verifique."
    Combo_Empresa.SetFocus
    Exit Sub
  End If
  
  If IsNull(Sequência.Text) Then Sequência.Text = 0
  If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
  If Val(Sequência.Text) < 0 Then Sequência.Text = 0
  
  If Nome_Cliente.Caption = "" Then
    DisplayMsg "Fornecedor inválido, verifique."
    Combo_Cliente.SetFocus
    Exit Sub
  End If
  
  If Nome_Custo.Caption = "" Then
    DisplayMsg "Centro de custo inváldo, verifique."
    Combo_Custo.SetFocus
    Exit Sub
  End If
  
  If IsNull(Sequência.Text) Then Sequência.Text = 0
  If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
  
  If Not IsDate(Emissão.Text) Then
   DisplayMsg "Data de emissão inválida."
   Emissão.SetFocus
   Exit Sub
  End If
  
  If Not IsDate(Vencimento.Text) Then
   DisplayMsg "Data de vencimento inválida."
   Vencimento.SetFocus
   Exit Sub
  End If
  
  If CDate(Vencimento.Text) < CDate(Emissão.Text) Then
    DisplayMsg "Data de vencimento não pode ser anterior à data de emissão."
    Vencimento.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If Not IsNumeric(Valor.Text) Then Valor.Text = 0
  If Erro = False Then If CDbl(Valor.Text) <= 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Valor incorreto."
    Valor.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0
  If Erro = False Then If CDbl(Desconto.Text) < 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Desconto incorreto."
    Desconto.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If Not IsNumeric(Acréscimo.Text) Then Acréscimo.Text = 0
  If CDbl(Acréscimo.Text) < 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Acréscimo incorreto."
    Acréscimo.SetFocus
    Exit Sub
  End If
  
  
  If IsDate(Data_Pagto.Text) Then
    If Not IsNumeric(Valor_Pago.Text) Then Valor_Pago.Text = 0
    If CDbl(Valor_Pago.Text) <> 0 Then
       Rem verifica soma
       Valor_Correto = CDbl(Valor.Text) - CDbl(Desconto.Text) + CDbl(Acréscimo.Text)
       If Abs((Valor_Correto - CDbl(Valor_Pago.Text))) > 0.001 Then
         DisplayMsg "Valor pago incorreto, valor correto seria " + str$(Valor_Correto)
         Valor_Pago.SetFocus
         Exit Sub
       End If
    End If
  End If
  
  Call StatusMsg("Gravando ...")
 
  ws.BeginTrans
  blnInTransaction = True
 
  With rsCP
    If IsNull(Num_Registro) Then
       .AddNew
       sTexto = "Lançamento efetuado."
    Else
      .LockEdits = True
      .Edit
      sTexto = "Lançamento alterado."
    End If

     Conta = .Fields("Contador")
    .Fields("Filial") = Combo_Empresa.Text
    .Fields("Fornecedor") = Combo_Cliente.Text
    .Fields("Centro de Custo") = Val(Combo_Custo.Text)
    .Fields("Sequência") = Sequência.Text
    .Fields("Nota") = Nota.Text
    .Fields("Descrição") = Descrição.Text
    .Fields("Data Emissão") = Emissão.Text
    .Fields("Vencimento") = Vencimento.Text
    .Fields("Valor") = CDbl(Valor.Text)
    .Fields("Desconto") = CDbl(Desconto.Text)
    .Fields("Acréscimo") = CDbl(Acréscimo.Text)
    .Fields("Valor Pago") = CDbl(gsHandleNull(Valor_Pago.Text & ""))
    If Not IsDate(Data_Pagto.Text) Then
      .Fields("Pagamento") = Null
    Else
      .Fields("Pagamento") = Data_Pagto.Text
    End If
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
    
    Aux_Filial = .Fields("Filial")
    Aux_Cliente = .Fields("Fornecedor")
    Aux_Vencimento = .Fields("Vencimento")
    
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  
  ws.CommitTrans
  blnInTransaction = False
  
  End With
  
  Call StatusMsg("")
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & Sequência.Text & " Cli:" & Combo_Cliente.Text & " Vr:" & Valor.Text & " VrPg:" & Valor_Pago.Text & " DtVc:" & Vencimento.Text, 80) & "', 'CNT_PAG: novo-atu')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************

  Exit Sub
  
Trata_Erro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
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
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Manutenção - Contas a receber")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Exit Sub
        Case 3 'Encerrar
          End
      End Select
  End Select
End Sub

Private Sub PrintCheque()

  Call StatusMsg("")
  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre ou grave um lançamento antes."
    Exit Sub
  End If
  
  frmImprimeCheque2.Favorecido.Text = Nome_Cliente.Caption
  frmImprimeCheque2.Valor.Text = Valor.Text
  frmImprimeCheque2.Show
  
End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  Combo_Empresa.Text = ""
  Nome_Empresa.Caption = ""
  Combo_Cliente.Text = ""
  Nome_Cliente.Caption = ""
  Combo_Custo.Text = ""
  Nome_Custo.Caption = ""
  Sequência.Text = ""
  Nota.Text = ""
  Descrição.Text = ""
  Emissão.Mask = ""
  Emissão.Text = ""
  Emissão.Mask = "##/##/####"
  Vencimento.Mask = ""
  Vencimento.Text = ""
  Vencimento.Mask = "##/##/####"
  Data_Pagto.Mask = ""
  Data_Pagto.Text = ""
  Data_Pagto.Mask = "##/##/####"
  Valor.Text = "0"
  Desconto.Text = "0"
  Acréscimo.Text = "0"
  Valor_Pago.Text = "0"
  
  If Not rsCP.EOF Then
    On Error Resume Next
    rsCP.MoveFirst
    rsCP.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  
  Combo_Empresa.SetFocus

End Sub

Private Sub ShowRecord()
  Combo_Empresa.Text = rsCP("Filial")
  Combo_Empresa_LostFocus
  Combo_Cliente.Text = rsCP("Fornecedor")
  Combo_Cliente_LostFocus
  Combo_Custo.Text = rsCP("Centro de Custo")
  Combo_Custo_LostFocus
  Sequência.Text = rsCP("Sequência")
  Nota.Text = rsCP("Nota") & ""
  Descrição.Text = rsCP("Descrição") & ""
  Emissão.Text = gsFormatDate(rsCP("Data Emissão"))
  Vencimento.Text = gsFormatDate(rsCP("Vencimento"))
  Valor.Text = rsCP("Valor")
  Desconto.Text = rsCP("Desconto")
  Acréscimo.Text = rsCP("Acréscimo")
  Valor_Pago.Text = rsCP("Valor Pago")
  If IsDate(rsCP("Pagamento")) Then
    Data_Pagto.Text = gsFormatDate(rsCP("Pagamento"))
  Else
    Data_Pagto.Mask = ""
    Data_Pagto.Text = ""
    Data_Pagto.Mask = "##/##/####"
  End If
  Num_Registro = rsCP.Bookmark
End Sub

' Botão removido em 22/06/2022 (Pablo)
'Private Sub cmdTotal_Click()
'  Valor_Pago.Text = Format(CDbl(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text)), "##,###,##0.00")
'End Sub

' criado em: 22/06/2022
' autor: Pablo Verçosa Silva
' descrição: calcula os campos Desconto, Acréscimo e Valor_Pago conforme edição
Private Sub CalculaValores(ByVal pStart As String)
  Dim nValor As Double
  Dim nDesconto As Double
  Dim nAcrescimo As Double
  Dim nValor_Pago As Double
  
  nValor = IIf(IsNumeric(Valor.Text), IIf(CDbl(gsHandleNull(Valor.Text)) > 0, CDbl(gsHandleNull(Valor.Text)), 0), 0)
  nDesconto = IIf(IsNumeric(Desconto.Text), IIf(CDbl(gsHandleNull(Desconto.Text)) > 0, CDbl(gsHandleNull(Desconto.Text)), 0), 0)
  nAcrescimo = IIf(IsNumeric(Acréscimo.Text), IIf(CDbl(gsHandleNull(Acréscimo.Text)) > 0, CDbl(gsHandleNull(Acréscimo.Text)), 0), 0)
  nValor_Pago = IIf(IsNumeric(Valor_Pago.Text), IIf(CDbl(gsHandleNull(Valor_Pago.Text)) > 0, CDbl(gsHandleNull(Valor_Pago.Text)), 0), 0)
  
  If StrComp(pStart, Desconto.Name, 1) = 0 Or StrComp(pStart, Acréscimo.Name, 1) = 0 Then nValor_Pago = nValor - nDesconto + nAcrescimo
  If StrComp(pStart, Valor_Pago.Name, 1) = 0 Then
    If nValor > nValor_Pago Then
      nDesconto = nValor - nValor_Pago
      nAcrescimo = 0
    ElseIf nValor < nValor_Pago Then
      nDesconto = 0
      nAcrescimo = nValor_Pago - nValor
    ElseIf nValor = nValor_Pago Then
      nDesconto = 0
      nAcrescimo = 0
    End If
  End If
  
  Desconto.Text = Format(nDesconto, "##,###,##0.00")
  Acréscimo.Text = Format(nAcrescimo, "##,###,##0.00")
  Valor_Pago.Text = Format(nValor_Pago, "##,###,##0.00")
  Data_Pagto.Text = Date
End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 0 Or Val(Combo_Cliente.Text) > 99999999 Then Exit Sub

  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  Nome_Cliente.Caption = rsClientes("Nome")

End Sub

Private Sub Combo_Custo_CloseUp()
 Combo_Custo.Text = Combo_Custo.Columns(1).Text
 Combo_Custo_LostFocus

End Sub

Private Sub Combo_Custo_LostFocus()
  Nome_Custo.Caption = ""
  If IsNull(Combo_Custo.Text) Then Exit Sub
  If Not IsNumeric(Combo_Custo.Text) Then Exit Sub
  If Val(Combo_Custo.Text) < 0 Or Val(Combo_Custo.Text) > 9999 Then Exit Sub

  rsCentros.Index = "Código"
  rsCentros.Seek "=", Val(Combo_Custo.Text)
  If rsCentros.NoMatch Then Exit Sub
  Nome_Custo.Caption = rsCentros("Nome")

End Sub

Private Sub Combo_Empresa_CloseUp()
 Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
 Combo_Empresa_LostFocus
End Sub

Private Sub Combo_Empresa_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Empresa.Text) Then Exit Sub
  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
  If Val(Combo_Empresa.Text) < 0 Or Val(Combo_Empresa.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Empresa.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")
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

Private Sub Desconto_GotFocus()
  Desconto.SelStart = 0
  Desconto.SelLength = 16
End Sub

Private Sub Desconto_LostFocus()
'  Valor_Pago.Text = CCur(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text))

' Calcula valores dos camos Desconto, Acréscimo e Valor_Pago (22/06/2022 - Pablo)
Call CalculaValores("Desconto")
End Sub

Private Sub Emissão_LostFocus()
  Emissão.Text = Ajusta_Data(Emissão.Text)
End Sub

Private Sub Emissão_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Emissão.Text = frmCalendario.gsDateCalender(Emissão.Text)
  End Select
End Sub

Private Sub Valor_GotFocus()
  Valor.SelStart = 0
  Valor.SelLength = 16
End Sub

Private Sub Valor_LostFocus()
'  Valor_Pago.Text = CCur(gsHandleNull(Valor.Text)) - CCur(gsHandleNull(Desconto.Text)) + CCur(gsHandleNull(Acréscimo.Text))
End Sub

Private Sub Valor_Pago_GotFocus()
  Valor_Pago.SelStart = 0
  Valor_Pago.SelLength = 16
End Sub

Private Sub Valor_Pago_LostFocus()
  ' Calcula valores dos camos Desconto, Acréscimo e Valor_Pago (22/06/2022 - Pablo)
  Call CalculaValores("Valor_Pago")
End Sub

Private Sub Vencimento_LostFocus()
  Vencimento.Text = Ajusta_Data(Vencimento.Text)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()

  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Filial, Vencimento"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Filial, Fornecedor"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 2, "Por Filial, Data Pagamento"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 3, "Por Filial, Centro de Custo"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 4, "Por Nota, Fornecedor"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  '25/04/2005 - Daniel
  'Otimizado rotina para abrir a tela de lançamentos de contas
  'com a conta selecionada a partir do duplo click na tela de
  'manutenções
  '
  'Solicitante: Consultor Carlos (Petrópolis - RJ)
  If frmManContasPagar.g_blnFind Then
    'Carregamos o CP com um único registro escolhido
    Set rsCP = db.OpenRecordset(frmManContasPagar.g_strQuery, dbOpenDynaset)
  Else
    gsSql = "SELECT * FROM [Contas a Pagar] "
    gsOrder = "ORDER BY Filial, Vencimento, Fornecedor, Contador"
    Set rsCP = db.OpenRecordset(gsSql & " " & gsOrder, dbOpenDynaset)
  End If
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsCentros = db.OpenRecordset("Centros de Custo", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName

  Call ActiveBarLoadToolTips(Me)
  
  Me.Show
  DoEvents
  
  Call ClearScreen
  
  '06/06/2005 - Daniel
  'Carregar automaticamente a Filial corrente
  'e a data atual para a Data de Emissão
  Combo_Empresa.Text = gnCodFilial
  Combo_Empresa_LostFocus
  
  Emissão.Text = Format(Data_Atual, "DD/MM/YYYY")
  Combo_Empresa.SetFocus
  '----------------------------------------------
  
  '25/04/2005 - Daniel
  'Exibição do registro a partir da tela de
  'manutenções
  If frmManContasPagar.g_blnFind Then
    Call MoveFirst
    frmManContasPagar.g_blnFind = False
  End If
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsClientes.Close
  rsCP.Close
  rsCentros.Close
  Set rsParametros = Nothing
  Set rsClientes = Nothing
  Set rsCP = Nothing
  Set rsCentros = Nothing
End Sub

