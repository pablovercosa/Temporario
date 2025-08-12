VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmGeraPagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lançar Parcelas de Contas a Pagar"
   ClientHeight    =   4470
   ClientLeft      =   2310
   ClientTop       =   1680
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GeraPagar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   13125
   Begin VB.TextBox txtNumParcInicial 
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
      Height          =   345
      Left            =   1695
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "1"
      Top             =   3570
      Width           =   525
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F7F7F7&
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   1875
   End
   Begin VB.CommandButton B_Gera 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gravar parcelas"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2070
      Width           =   1875
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
      Height          =   345
      Left            =   6060
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Centros de Custo] WHERE Ativo ORDER BY Nome"
      Top             =   4230
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Data Data2 
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
      Left            =   4500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Fornecedor"
      Top             =   4230
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Data Data1 
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
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4230
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ListBox Lista1 
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
      Height          =   4230
      Left            =   9150
      TabIndex        =   14
      Top             =   105
      Width           =   1890
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Custo 
      Bindings        =   "GeraPagar.frx":4E95A
      DataSource      =   "Data3"
      Height          =   375
      Left            =   1695
      TabIndex        =   2
      Top             =   945
      Width           =   960
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
      _ExtentX        =   1693
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
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
   Begin VB.TextBox Nota 
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
      Height          =   345
      Left            =   1695
      MaxLength       =   15
      TabIndex        =   7
      Top             =   3990
      Width           =   1275
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   375
      Left            =   1695
      TabIndex        =   6
      Top             =   2490
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
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
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton B_Mostra 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar parcelas >>"
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
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2070
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Próximos Vencimentos"
      Height          =   600
      Left            =   3090
      TabIndex        =   26
      Top             =   1830
      Width           =   3960
      Begin VB.OptionButton O_Período 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "A cada"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   285
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton O_Dia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todo dia"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2310
         TabIndex        =   12
         Top             =   225
         Width           =   930
      End
      Begin VB.TextBox Período 
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
         Height          =   345
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "30"
         Top             =   195
         Width           =   450
      End
      Begin VB.TextBox Dia 
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
         Height          =   345
         Left            =   3270
         MaxLength       =   2
         TabIndex        =   13
         Top             =   195
         Width           =   450
      End
      Begin VB.Label Label10 
         Caption         =   "dias"
         Height          =   225
         Left            =   1665
         TabIndex        =   27
         Top             =   255
         Width           =   375
      End
   End
   Begin VB.TextBox Parcelas 
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
      Height          =   375
      Left            =   4590
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2490
      Width           =   540
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Left            =   1695
      TabIndex        =   5
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2025
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Descrição 
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
      Height          =   375
      Left            =   1695
      MaxLength       =   14
      TabIndex        =   4
      Top             =   2970
      Width           =   1965
   End
   Begin MSMask.MaskEdBox Data_Emissão 
      Height          =   375
      Left            =   1695
      TabIndex        =   3
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1575
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "GeraPagar.frx":4E96E
      DataSource      =   "Data2"
      Height          =   375
      Left            =   1695
      TabIndex        =   1
      Top             =   525
      Width           =   960
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
      _ExtentX        =   1693
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "GeraPagar.frx":4E982
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1695
      TabIndex        =   0
      Top             =   105
      Width           =   960
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
      _ExtentX        =   1693
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
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
   Begin VB.Label Label2 
      Caption         =   "Número Parcela Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   32
      Top             =   3630
      Width           =   1635
   End
   Begin VB.Label Nome_Custo 
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2700
      TabIndex        =   31
      Top             =   945
      Width           =   4365
   End
   Begin VB.Label Label15 
      Caption         =   "Centro de Custo"
      Height          =   195
      Left            =   60
      TabIndex        =   30
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Nota"
      Height          =   195
      Left            =   60
      TabIndex        =   29
      Top             =   4050
      Width           =   510
   End
   Begin VB.Label Label13 
      Caption         =   "Valor da Parcela"
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   2610
      Width           =   1170
   End
   Begin VB.Label Label8 
      Caption         =   "Qtde de Parcelas"
      Height          =   195
      Left            =   3300
      TabIndex        =   25
      Top             =   2580
      Width           =   1290
   End
   Begin VB.Label Label7 
      Caption         =   "Vencimento Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Descrição"
      Height          =   195
      Left            =   60
      TabIndex        =   23
      Top             =   3060
      Width           =   960
   End
   Begin VB.Label Label5 
      Caption         =   "Data Emissão"
      Height          =   195
      Left            =   60
      TabIndex        =   22
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label Nome_Fornecedor 
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2700
      TabIndex        =   21
      Top             =   525
      Width           =   4365
   End
   Begin VB.Label Label3 
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   60
      TabIndex        =   20
      Top             =   615
      Width           =   960
   End
   Begin VB.Label Nome_Empresa 
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2700
      TabIndex        =   19
      Top             =   105
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   "Filial"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   180
      Width           =   465
   End
End
Attribute VB_Name = "frmGeraPagar"
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

Dim rsParametros As Recordset
Dim rsCustos As Recordset
Dim rsFornecedores As Recordset
Dim rsContas_Pagar As Recordset

Function Verifica_Dados() As Integer
 
 Dim Erro As Integer
 

 Call StatusMsg("")

 If Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo_Empresa.SetFocus
   Verifica_Dados = 1
   Exit Function
 End If
 
 If Nome_Fornecedor.Caption = "" Then
   DisplayMsg "Escolha o fornecedor."
   Combo_Fornecedor.SetFocus
   Verifica_Dados = 1
   Exit Function
 End If

 If Nome_Custo.Caption = "" Then Combo_Custo.Text = ""
 
 If Not IsDate(Data_Emissão.Text) Then
   DisplayMsg "Data de emissão incorreta."
   Data_Emissão.SetFocus
   Verifica_Dados = 1
   Exit Function
 End If
 
 If Not IsDate(Vencimento.Text) Then
   DisplayMsg "Data de vencimento incorreta."
   Vencimento.SetFocus
   Verifica_Dados = 1
   Exit Function
 End If

 Erro = False
 If IsNull(Valor.Text) Then Erro = True
 If Erro = False Then If Valor.Text = "" Then Erro = True
 If Erro = False Then If Not IsNumeric(Valor.Text) Then Erro = True
 If Erro = False Then If CDbl(Valor.Text) <= 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Valor incorreto, verifique."
   Valor.SetFocus
   Verifica_Dados = 1
   Exit Function
 End If
 
 Erro = False
 If IsNull(Parcelas.Text) Then Erro = True
 If Erro = False Then If Parcelas.Text = "" Then Erro = True
 If Erro = False Then If Not IsNumeric(Parcelas.Text) Then Erro = True
 If Erro = False Then If Val(Parcelas.Text) <= 0 Then Erro = True
 If Erro = False Then If Val(Parcelas.Text) > 180 Then Erro = True
 If Erro = True Then
   DisplayMsg "Digite de 2 a 180 parcelas."
   Parcelas.SetFocus
   Verifica_Dados = 1
   Exit Function
 End If
 
 If O_Período.Value = True Then
  Erro = False
  If IsNull(Período.Text) Then Erro = True
  If Erro = False Then If Período.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Período.Text) Then Erro = True
  If Erro = False Then If (Val(Período.Text) < 1 Or Val(Período.Text) > 180) Then Erro = True
  If Erro = True Then
    DisplayMsg "Digite de 2 a 180 dias de intervalo."
    Período.SetFocus
    Verifica_Dados = 1
    Exit Function
  End If
 End If
  
 If O_Dia.Value = True Then
  Erro = False
  If IsNull(Dia.Text) Then Erro = True
  If Erro = False Then If Dia.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Dia.Text) Then Erro = True
  If Erro = False Then If (Val(Dia.Text) < 1 Or Val(Dia.Text) > 180) Then Erro = True
  If Erro = True Then
    DisplayMsg "Dia deve estar entre 1 e 31."
    Dia.SetFocus
    Verifica_Dados = 1
    Exit Function
  End If
 End If

 Verifica_Dados = 0

End Function

Private Sub B_Gera_Click()
  Dim nI As Integer
  Dim nQtdeParc As Integer
  Dim sQtdeFinal As String
  Dim nIni As Integer
  
  txtNumParcInicial.Text = gsHandleNull(txtNumParcInicial.Text)
  If Not IsNumeric(txtNumParcInicial.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Número da Parcela Inicial inválido."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    txtNumParcInicial.SetFocus
    Exit Sub
  End If
  If CInt(txtNumParcInicial.Text) <= 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Número da Parcela Inicial inválido."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    txtNumParcInicial.SetFocus
    Exit Sub
  End If
  
  If Verifica_Dados = 1 Then Exit Sub
  
  B_Mostra_Click
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente gerar as parcelas de contas a pagar ? Este processo NÃO poderá ser desfeito."
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    DisplayMsg "Contas não geradas."
    Exit Sub
  End If
  
  nQtdeParc = Lista1.ListCount
  nIni = CInt(txtNumParcInicial.Text)
  sQtdeFinal = CStr(nIni + nQtdeParc - 1)
  
  Screen.MousePointer = vbHourglass
  
  For nI = 0 To nQtdeParc - 1
    Call StatusMsg("Gerando parcela " & nI)
    rsContas_Pagar.AddNew
    rsContas_Pagar("Filial") = Combo_Empresa.Text
    rsContas_Pagar("Fornecedor") = Combo_Fornecedor.Text
    rsContas_Pagar("Data Emissão") = CDate(Data_Emissão.Text)
    rsContas_Pagar("Descrição") = Descrição.Text & " Parc. " & nIni & "/" & sQtdeFinal
    rsContas_Pagar("Vencimento") = CDate(Lista1.List(nI))
    rsContas_Pagar("Valor") = CDbl(Valor.Text)
    rsContas_Pagar("Desconto") = 0
    rsContas_Pagar("Acréscimo") = 0
    rsContas_Pagar("Valor Pago") = 0
    rsContas_Pagar("Sequência") = 0
    rsContas_Pagar("Nota") = Nota.Text
    rsContas_Pagar("Centro de Custo") = Val(Combo_Custo.Text)
    rsContas_Pagar("Data Alteração") = Format(Data_Atual, "dd/mm/yyyy")
    rsContas_Pagar.Update
    nIni = nIni + 1
  Next nI

  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Forn:" & Combo_Fornecedor.Text & " Venc:" & Vencimento.Text & " Vlr:" & Valor.Text & " Qtd Parc:" & Parcelas.Text, 80) & "', 'CNT_PAG: gera parc')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************

  Screen.MousePointer = vbDefault
  
  DisplayMsg "Parcelas de Contas a Pagar geradas."
  
  Call StatusMsg("")
  
  B_Gera.Enabled = False

End Sub

Private Sub B_Mostra_Click()
 Dim i As Integer
 Dim Aux_Vcto As String
 Dim Aux_Str As String
 Dim Aux_Mês As Integer
 Dim Aux_Ano As Integer
 Dim Aux_Dia As Integer
   

 If Verifica_Dados = 1 Then Exit Sub
 
 Lista1.Clear
 
 Aux_Vcto = Vencimento.Text
 Aux_Vcto = Format(Aux_Vcto, "dd/mm/yyyy")
 
 
 For i = 1 To Val(Parcelas.Text)
  Lista1.AddItem (Aux_Vcto)
   
  Aux_Str = Vencimento.Text
  If O_Período.Value = True Then
    Aux_Vcto = str(CDate(Aux_Vcto) + Val(Período.Text))
    Aux_Vcto = Format(Aux_Vcto, "dd/mm/yyyy")
  End If
  
  If O_Dia.Value = True Then
    Aux_Str = Aux_Vcto
    Aux_Mês = Mid(Aux_Str, 4, 2)
    Aux_Ano = Mid(Aux_Str, 7, 4)
  
    Aux_Mês = Aux_Mês + 1
    If Aux_Mês = 13 Then
      Aux_Ano = Aux_Ano + 1
      Aux_Mês = 1
    End If
    
    Aux_Dia = Dia.Text
    
    If Aux_Mês = 2 Then
     If Aux_Dia > 28 Then
       Aux_Dia = 28
     End If
    End If
    
    If Aux_Mês = 4 Or Aux_Mês = 6 Or Aux_Mês = 9 Or Aux_Mês = 11 Then
      If Aux_Dia > 30 Then Aux_Dia = 30
    End If
    
    Aux_Str = str(Aux_Dia) + "/" + str(Aux_Mês) + "/" + str(Aux_Ano)
    Aux_Str = Trim(Aux_Str)
    Aux_Vcto = Format(Aux_Str, "dd/mm/yyyy")
    
  End If
  
 Next i
 
End Sub

Private Sub cmdClear_Click()
  Combo_Empresa.Text = ""
  Combo_Fornecedor.Text = ""
  Combo_Custo.Text = ""
  Data_Emissão.Mask = ""
  Data_Emissão.Text = ""
  Data_Emissão.Mask = "##/##/####"
  Descrição.Text = ""
  Vencimento.Mask = ""
  Vencimento.Text = ""
  Vencimento.Mask = "##/##/####"
  Valor.Text = ""
  Nota.Text = ""
  Parcelas.Text = ""
  Dia.Text = ""
  Lista1.Clear
  Nome_Empresa.Caption = ""
  Nome_Fornecedor.Caption = ""
  Nome_Custo.Caption = ""
  B_Gera.Enabled = True
End Sub

Private Sub Combo_Custo_CloseUp()
 Combo_Custo.Text = Combo_Custo.Columns(1).Text
 Combo_Custo_LostFocus
End Sub

Private Sub Combo_Custo_LostFocus()
  
  Call StatusMsg("")

  Nome_Custo.Caption = ""
  If IsNull(Combo_Custo.Text) Then Exit Sub
  If Combo_Custo.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Custo.Text) Then Exit Sub
  
  If Val(Combo_Custo.Text) < 0 Or Val(Combo_Custo.Text) > 9999 Then Exit Sub
  
  rsCustos.Index = "Código"
  rsCustos.Seek "=", Val(Combo_Custo.Text)
  If rsCustos.NoMatch Then Exit Sub
  
  Nome_Custo.Caption = rsCustos("Nome") & ""

End Sub

Private Sub Combo_Empresa_CloseUp()
 Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
 Combo_Empresa_LostFocus
End Sub

Private Sub Combo_Empresa_LostFocus()
  Call StatusMsg("")

  Nome_Empresa.Caption = ""
  If IsNull(Combo_Empresa.Text) Then Exit Sub
  If Combo_Empresa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
  
  If Val(Combo_Empresa.Text) < 0 Or Val(Combo_Empresa.Text) > 99 Then Exit Sub
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Empresa.Text)
  If rsParametros.NoMatch Then Exit Sub
  
  Nome_Empresa.Caption = rsParametros("Nome") & ""
End Sub

Private Sub Combo_Fornecedor_CloseUp()
 Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
 Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Call StatusMsg("")

  Nome_Fornecedor.Caption = ""
  If IsNull(Combo_Fornecedor.Text) Then Exit Sub
  If Combo_Fornecedor.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  
  If Val(Combo_Fornecedor.Text) < 0 Or Val(Combo_Fornecedor.Text) > 99999999 Then Exit Sub
  
  rsFornecedores.Index = "Código"
  rsFornecedores.Seek "=", Val(Combo_Fornecedor.Text)
  If rsFornecedores.NoMatch Then Exit Sub
  
  Nome_Fornecedor.Caption = rsFornecedores("Nome") & ""
  
End Sub

Private Sub Data_Emissão_LostFocus()
  Data_Emissão.Text = Ajusta_Data(Data_Emissão.Text)
End Sub

Private Sub Data_Emissão_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Emissão.Text = frmCalendario.gsDateCalender(Data_Emissão.Text)
  End Select
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsCustos = db.OpenRecordset("Centros de Custo", , dbReadOnly)
  Set rsFornecedores = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsContas_Pagar = db.OpenRecordset("Contas A Pagar")
  
  Data_Emissão.Text = Format(Data_Atual, "dd/MM/yyyy")

End Sub

Private Sub O_Dia_Click()
 Dia.SetFocus
End Sub

Private Sub O_Período_Click()
 Período.SetFocus
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
