VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProdutosBatchUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização em grupo do Cadastro de Produtos"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   Icon            =   "frmProdutosBatchUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Atualizar"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   6120
      Width           =   1935
   End
   Begin TabDlg.SSTab sstItens 
      Height          =   3975
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Loja Virtual (WEB)"
      TabPicture(0)   =   "frmProdutosBatchUpdate.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBonus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTableOfPrices"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAttributes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraOffer"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "datTabelaPreco"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkChangeAttributes"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkChangeBonus"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkChangeTableOfPrices"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkChangeOffer"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CheckBox chkChangeOffer 
         Caption         =   "Alterar Promoção"
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkChangeTableOfPrices 
         Caption         =   "Alterar Tabela de Preços"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkChangeBonus 
         Caption         =   "Alterar Bônus"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkChangeAttributes 
         Caption         =   "Alterar Atributos"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Data datTabelaPreco 
         Caption         =   "TabPrecos"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Tabela From [Tabela de Preços] ORDER BY Tabela"
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame fraOffer 
         Caption         =   "Promoção"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   3120
         TabIndex        =   25
         Top             =   840
         Width           =   6615
         Begin MSComCtl2.DTPicker dtpDataInicial 
            Height          =   315
            Left            =   2880
            TabIndex        =   6
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   97255425
            CurrentDate     =   37333
         End
         Begin MSComCtl2.DTPicker dtpDataFinal 
            Height          =   315
            Left            =   4800
            TabIndex        =   7
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   97255425
            CurrentDate     =   37333
         End
         Begin SSDataWidgets_B.SSDBCombo cboTablePromocional 
            Bindings        =   "frmProdutosBatchUpdate.frx":05A6
            DataSource      =   "datTabelaPreco"
            Height          =   315
            Left            =   360
            TabIndex        =   5
            Top             =   600
            Width           =   1815
            DataFieldList   =   "Tabela"
            _Version        =   196617
            BackColorOdd    =   14737632
            Columns(0).Width=   3200
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "de"
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   28
            Top             =   660
            Width           =   180
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Index           =   1
            Left            =   4440
            TabIndex        =   27
            Top             =   660
            Width           =   225
         End
         Begin VB.Label Label62 
            Caption         =   "Tabela de Preços"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraAttributes 
         Caption         =   "Atributos"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   2775
         Begin VB.CheckBox chkFabricante 
            Caption         =   "Fabricante"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkPesquisa123 
            Caption         =   "Pesquisa 123"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame fraTableOfPrices 
         Caption         =   "Preço de Venda"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   2775
         Begin SSDataWidgets_B.SSDBCombo cboTableVenda 
            Bindings        =   "frmProdutosBatchUpdate.frx":05C3
            DataSource      =   "datTabelaPreco"
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   1815
            DataFieldList   =   "Tabela"
            _Version        =   196617
            BackColorOdd    =   14737632
            Columns(0).Width=   3200
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label Label61 
            Caption         =   "Tabela de Preços"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraBonus 
         Caption         =   "Bônus"
         Enabled         =   0   'False
         Height          =   855
         Left            =   3120
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   2775
         Begin VB.TextBox txtBonus 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   720
            MaxLength       =   6
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção de produtos"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   9975
      Begin VB.Data datSubClasse 
         Caption         =   "Classe"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "Con_Sub_Classe"
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data datClasse 
         Caption         =   "Classe"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "Con_Classe"
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin Threed.SSPanel sspClasse 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   600
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmProdutosBatchUpdate.frx":05E0
         DataSource      =   "datClasse"
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1215
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmProdutosBatchUpdate.frx":05F8
         DataSource      =   "datSubClasse"
         Height          =   315
         Left            =   5160
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin Threed.SSPanel sspSubClasse 
         Height          =   315
         Left            =   6480
         TabIndex        =   19
         Top             =   600
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Sub Classe"
         Height          =   195
         Index           =   2
         Left            =   5160
         TabIndex        =   20
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Label lblTitle 
      Caption         =   "Selecione a Classe (opcional) e Sub Classe (opcional) dos produtos que deseja realizar as alterações e clique em Atualizar. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "frmProdutosBatchUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'26/07/2002 - Form implementado por mpdea

'10/10/2002 - mpdea
'Comentado referências ao TextBox txtBonus da pasta WEB
'(aguardando definições para implementação futura)

Option Explicit

Private Sub cboClasse_CloseUp()
  cboClasse.Text = cboClasse.Columns(1).Text
  Call cboClasse_Validate(False)
End Sub

Private Sub cboClasse_InitColumnProps()
  With cboClasse
    .Columns(0).Width = 3500
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub cboClasse_Validate(Cancel As Boolean)
  Dim intRet As Integer
  Dim strRet As String
  
  If IsDataType(dtInteger, cboClasse.Text, intRet) Then
    cboClasse.Text = intRet
    strRet = gsGetNameClasse(intRet)
    sspClasse.Caption = strRet
  End If
  
  If strRet = "" Then
    cboClasse.Text = "0"
    sspClasse.Caption = "Todas"
  End If
  
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(1).Text
  Call cboSubClasse_Validate(False)
End Sub

Private Sub cboSubClasse_InitColumnProps()
  With cboSubClasse
    .Columns(0).Width = 3500
    .Columns(1).Width = 1000
  End With
End Sub

Private Sub cboSubClasse_Validate(Cancel As Boolean)
  Dim intRet As Integer
  Dim strRet As String
  
  If IsDataType(dtInteger, cboSubClasse.Text, intRet) Then
    cboSubClasse.Text = intRet
    strRet = gsGetNameSubClasse(intRet)
    sspSubClasse.Caption = strRet
  End If
  
  If strRet = "" Then
    cboSubClasse.Text = "0"
    sspSubClasse.Caption = "Todas"
  End If
  
End Sub

Private Sub chkChangeAttributes_Click()
  fraAttributes.Enabled = chkChangeAttributes.Value = vbChecked
End Sub

Private Sub chkChangeBonus_Click()
  fraBonus.Enabled = chkChangeBonus.Value = vbChecked
End Sub

Private Sub chkChangeOffer_Click()
  fraOffer.Enabled = chkChangeOffer.Value = vbChecked
End Sub

Private Sub chkChangeTableOfPrices_Click()
  fraTableOfPrices.Enabled = chkChangeTableOfPrices.Value = vbChecked
End Sub

Private Sub cmdUpdate_Click()
  Dim strAux As String
  Dim lngRet As Long
  Dim blnBegin As Boolean
  
  On Error GoTo ErrHandler
  
  If chkChangeTableOfPrices.Value = vbChecked Then
    'Verifica se a tabela de venda está correta
    If Not gbCheckTabPreco(cboTableVenda.Text) Then
      DisplayMsg "WEB - Tabela de venda incorreta."
      SelectAllText cboTableVenda, True
      Exit Sub
    End If
    blnBegin = True
  End If
  
  If chkChangeOffer.Value = vbChecked Then
    'Promoção
    If cboTablePromocional.Text <> "" Then
      If Not gbCheckTabPreco(cboTablePromocional.Text) Then
        DisplayMsg "Web - Tabela de promoção incorreta."
        SelectAllText cboTablePromocional, True
        Exit Sub
      End If
    End If
    
    If gbCheckTabPreco(cboTablePromocional.Text) Then
      If Not IsDate(dtpDataInicial.Value) Then
        DisplayMsg "Web - Data início da Promoção incorreta."
        SelectAllText dtpDataInicial, True
        Exit Sub
      End If
  
      If Not IsDate(dtpDataFinal.Value) Then
        DisplayMsg "Web - Data Final da Promoção incorreta."
        SelectAllText dtpDataFinal, True
        Exit Sub
      End If
    
      If CDate(dtpDataFinal.Value) < CDate(dtpDataInicial.Value) Then
        DisplayMsg "Web - Data final inferior a data inicial."
        SelectAllText dtpDataFinal, True
        Exit Sub
      End If
    End If
      
    blnBegin = True
  End If
  
  If chkChangeAttributes.Value = vbChecked Then
    blnBegin = True
  End If
    
'  If chkChangeBonus.Value = vbChecked Then
'    'Bônus
'    If Not IsDataType(dtLong, txtBonus.Text, lngRet) Then
'      DisplayMsg "WEB - Quantidade de Bônus incorreta."
'      SelectAllText txtBonus, True
'      Exit Sub
'    End If
'    txtBonus.Text = lngRet
'    blnBegin = True
'  End If
  
  
  'Verificação
  If Not blnBegin Then
    DisplayMsg "Selecione ao menos um grupo para alteração."
    chkChangeAttributes.SetFocus
    Exit Sub
  End If
  
  strAux = "As alterações para a pasta 'Loja Virtual (WEB)' " & _
           "serão realizadas somente para produtos que não sejam do " & _
           "tipo 'Fracionado' e que a opção 'Controlar estoque deste produto' " & _
           "esteja ativada. Deseja continuar"
  
  If MsgBox(strAux, vbExclamation + vbYesNo, "Atenção") = vbNo Then Exit Sub
    
  
  'Monta string SQL
  blnBegin = False
  strAux = "UPDATE Produtos SET"
  
  If chkChangeAttributes.Value = vbChecked Then
    strAux = strAux & " WebAttribFabricante = " & _
             IIf((chkFabricante.Value = vbChecked), True, False)
    strAux = strAux & ", WebAttribPesquisa123 = " & _
             IIf((chkPesquisa123.Value = vbChecked), True, False)
    blnBegin = True
  End If
  
'  If chkChangeBonus.Value = vbChecked Then
'    If blnBegin Then strAux = strAux & ","
'    strAux = strAux & " WebBonus = " & txtBonus.Text
'    blnBegin = True
'  End If
  
  If chkChangeTableOfPrices.Value = vbChecked Then
    If blnBegin Then strAux = strAux & ","
    strAux = strAux & " WebSaleTablePrice = '" & cboTableVenda.Text & "'"
    blnBegin = True
  End If
  
  If chkChangeOffer.Value = vbChecked Then
    If blnBegin Then strAux = strAux & ","
    strAux = strAux & " WebOfferTablePrice = '" & cboTablePromocional.Text & "'"
    strAux = strAux & ", WebOfferDateStart = #" & _
             Format(dtpDataInicial.Value, SQL_DATE_MASK) & "#"
    strAux = strAux & ", WebOfferDateEnd = #" & _
             Format(dtpDataFinal.Value, SQL_DATE_MASK) & "#"
    blnBegin = True
  End If
  
  'Produto deve ser sincronizado
  strAux = strAux & ", WebSynchronize = True"
  
  
  'Seleção a alterar
  strAux = strAux & " WHERE Código <> '0' AND NOT Fracionado AND Estoque AND WebIncluded"
  
  Call cboClasse_Validate(False)
  Call cboSubClasse_Validate(False)
  
  If cboClasse.Text <> "0" Then
    strAux = strAux & " AND Classe = " & cboClasse.Text
  End If
  
  If cboSubClasse.Text <> "0" Then
    strAux = strAux & " AND [Sub Classe] = " & cboSubClasse.Text
  End If
  
  'Executa SQL
  db.Execute strAux, dbFailOnError
  
  DisplayMsg "Cadastro de Produtos atualizado."
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & "-" & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Form_Load()
  
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datTabelaPreco.DatabaseName = gsQuickDBFileName
  
  dtpDataInicial.Value = Date
  dtpDataFinal.Value = Date
  
  '13/08/2002 - mpdea
  'Atributo da pasta WEB para envio de informações da Pesquisa 123
  chkPesquisa123.Visible = gsPesq1 <> "" Or gsPesq2 <> "" Or gsPesq3 <> ""
  
End Sub
