VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmVendaValorXQtde 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incluir Produto informando o Valor Total"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmVendaValorXQtde.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValorUnit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtQtdeXValor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtDifQtdeXValor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtQuantidade 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   480
      TabIndex        =   2
      Text            =   "0,00"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   4440
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Informe o produto e o valor total para o cálculo da quantidade"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produto"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   9615
      Begin SSDataWidgets_B.SSDBCombo cboProdCod 
         Bindings        =   "frmVendaValorXQtde.frx":058A
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2655
         DataFieldList   =   "Código"
         _Version        =   196617
         BackColorOdd    =   8438015
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   10954
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Data datProdNome 
         Caption         =   "Produto Nome"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Nome, Código FROM Produtos WHERE Código <> '0' AND NOT Desativado"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Data datProdCod 
         Caption         =   "Produto Código"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' AND NOT Desativado"
         Top             =   1320
         Width           =   2535
      End
      Begin SSDataWidgets_B.SSDBCombo cboProdNome 
         Bindings        =   "frmVendaValorXQtde.frx":05A3
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   6375
         DataFieldList   =   "Nome"
         _Version        =   196617
         BackColorOdd    =   8438015
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   10954
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   11245
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Valor Unitário para cálculo"
      Height          =   195
      Index           =   7
      Left            =   7200
      TabIndex        =   17
      Top             =   1920
      Width           =   1860
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Valor Final"
      Height          =   195
      Index           =   6
      Left            =   7200
      TabIndex        =   15
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Diferença"
      Height          =   195
      Index           =   4
      Left            =   7200
      TabIndex        =   13
      Top             =   3360
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   11
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   765
   End
End
Attribute VB_Name = "frmVendaValorXQtde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnPressOK As Boolean

Private mstrTablePrice As String
Private mblnCalculateIPI As Boolean
Private mintQtdeCasasDecimais As Integer

Private mstrRET_ProdCod As String
Private mstrRET_Quantidade As String

Public Sub Start(ByVal strTablePrice As String, ByVal blnCalculateIPI As Boolean)
  Dim intCancel As Integer
  Dim intX As Integer
  
  mblnPressOK = False
  mstrTablePrice = strTablePrice
  mblnCalculateIPI = blnCalculateIPI
  
  Me.Show vbModal
  
  'Insere o item
  If mblnPressOK Then
    
    '18/01/2006 - mpdea
    'Alterado objeto frmVendaRap2 -> g_frmVendaRapida
    With g_frmVendaRapida
      'Procura linha em branco
      .Grade1.MoveFirst
      
      '30/03/2006 - mpdea
      'Corrigido verificação de próxima linha sem produto
      For intX = 1 To .Grade1.Rows
        If .Grade1.Columns(0).Text = "0" Or .Grade1.Columns(0).Text = "" Then Exit For
        .Grade1.MoveNext
      Next intX
      
      'Insere o item
      .Grade1.Columns(0).Text = mstrRET_ProdCod
      .Grade1.Columns(1).Text = mstrRET_Quantidade
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
    
  End If
  
End Sub

Private Sub ClearScreen()

  '05/09/2006 - Anderson - Implementação de 5 casas decimais
  If g_bln5CasasDecimais Then
    txtValorUnit.Text = Format(0, "##,###,##0.00000")
    txtValorTotal.Text = Format(0, "##,###,##0.00000")
    txtQuantidade.Text = "0"
    txtDifQtdeXValor.Text = Format(0, "##,###,##0.00000")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    txtValorUnit.Text = Format(0, "##,###,##0.000")
    txtValorTotal.Text = Format(0, "##,###,##0.000")
    txtQuantidade.Text = "0"
    txtDifQtdeXValor.Text = Format(0, "##,###,##0.000")
  Else
    txtValorUnit.Text = Format(0, FORMAT_VALUE)
    txtValorTotal.Text = Format(0, FORMAT_VALUE)
    txtQuantidade.Text = "0"
    txtDifQtdeXValor.Text = Format(0, FORMAT_VALUE)
  End If
  
  txtQtdeXValor.Text = Format(0, "##,###,##0.00")

End Sub

Private Sub cboProdCod_Change()
  cboProdNome.Text = ""
End Sub

Private Sub cboProdCod_Click()
  Call cboProdCod_LostFocus
End Sub

Private Sub cboProdCod_CloseUp()
  cboProdCod.Text = cboProdCod.Columns("Código").Text
  Call cboProdCod_LostFocus
End Sub

Private Sub cboProdCod_LostFocus()
  Dim strCodigo As String
  Dim curPrice As Currency
  Dim sngDesconto As Single
  Dim sngIPI As Single

  'Retira espaços em branco do código
  strCodigo = Replace(cboProdCod.Text, " ", "")
  cboProdCod.Text = UCase(strCodigo)
  
  If strCodigo = "" Then Exit Sub
  
  With datProdCod.Recordset
    .FindFirst "Código = '" & strCodigo & "'"
    If .NoMatch Then
      Call ClearScreen
      DisplayMsg "Produto não encontrado."
      Call SelectAllText(cboProdCod, True)
    Else
      cboProdNome.Text = .Fields("Nome").Value & ""
      
      curPrice = gcGetPrecoProduto(strCodigo, mstrTablePrice)
      
      'Aplica desconto
      Call IsDataType(dtSingle, .Fields("Desconto").Value, sngDesconto)
      If sngDesconto > 0 Then
        curPrice = curPrice * (1 - sngDesconto / 100)
      End If
      
      'Aplica IPI
      If mblnCalculateIPI Then
        Call IsDataType(dtSingle, .Fields("Percentual IPI").Value, sngIPI)
        If sngIPI > 0 Then
          curPrice = curPrice * (1 + sngIPI / 100)
        End If
      End If
      
      '05/09/2006 - Anderson - Implementação de 5 casas decimais
      If g_bln5CasasDecimais Then
        txtValorUnit.Text = Format(curPrice, "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        txtValorUnit.Text = Format(curPrice, "##,###,##0.000")
      Else
        txtValorUnit.Text = Format(curPrice, FORMAT_VALUE)
      End If
      
      If .Fields("Fracionado").Value Then
        mintQtdeCasasDecimais = CInt("0" & .Fields("QtdeCasasDecimais").Value)
      Else
        mintQtdeCasasDecimais = 0
      End If
      Call txtValorTotal_Change
    End If
  End With

End Sub

Private Sub cboProdNome_Click()
  cboProdCod.Text = cboProdNome.Columns("Código").Text
  Call cboProdCod_LostFocus
End Sub

Private Sub cboProdNome_CloseUp()
  cboProdCod.Text = cboProdNome.Columns("Código").Text
  Call cboProdCod_LostFocus
End Sub

Private Sub cboProdNome_LostFocus()
  Call cboProdCod_LostFocus
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim sngQuantidade As Single
      
  Call IsDataType(dtSingle, txtQuantidade.Text, sngQuantidade)
  
  If sngQuantidade = 0 Then
    Call DisplayMsg("Quantidade incompleta.")
    Call SelectAllText(txtValorTotal, True)
    Exit Sub
  End If
  
  mstrRET_ProdCod = cboProdCod.Text
  mstrRET_Quantidade = txtQuantidade.Text
  
  mblnPressOK = True
  Unload Me
  
End Sub

Private Sub Form_Load()
  Dim strSQL As String
  
  strSQL = "SELECT * FROM Produtos WHERE Código <> '0' AND NOT Desativado ORDER BY "
  
  datProdCod.DatabaseName = gsQuickDBFileName
  datProdCod.RecordSource = strSQL & "[Código Ordenação]"
  
  datProdNome.DatabaseName = gsQuickDBFileName
  datProdNome.RecordSource = strSQL & "Nome"
  
  Call ClearScreen

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmVendaValorXQtde = Nothing
End Sub

Private Sub txtValorTotal_Change()
  Dim curValorTotal As Currency
  Dim sngQuantidade As Single
  Dim curValorUnit As Currency
  
  'Valida entrada de dados
  Call IsDataType(dtCurrency, txtValorTotal.Text, curValorTotal)
  Call IsDataType(dtCurrency, txtValorUnit.Text, curValorUnit)
  
  '05/09/2006 - Anderson - Implementação de 5 casas decimais
  If g_bln5CasasDecimais Then
    txtDifQtdeXValor.Text = Format(0, "##,###,##0.00000")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    txtDifQtdeXValor.Text = Format(0, "##,###,##0.000")
  Else
    txtDifQtdeXValor.Text = Format(0, FORMAT_VALUE)
  End If
  txtQtdeXValor.Text = Format(0, "##,###,##0.00")
  
  If curValorTotal = 0 Or curValorUnit = 0 Then
    sngQuantidade = 0
  Else
    
    'Trunca a quantidade
    sngQuantidade = Truncate(curValorTotal / curValorUnit, mintQtdeCasasDecimais)
    
    txtQtdeXValor.Text = Format(sngQuantidade * curValorUnit, "##,###,##0.00")
    
    'Diferença entre o valor total e o final
    If sngQuantidade * curValorUnit <> curValorTotal Then
      '05/09/2006 - Anderson - Implementação de 5 casas decimais
      If g_bln5CasasDecimais Then
        txtDifQtdeXValor.Text = Format(curValorTotal - curValorUnit * sngQuantidade, "##,###,##0.00000")
      '30/04/2007 - Anderson - Implementação de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        txtDifQtdeXValor.Text = Format(curValorTotal - curValorUnit * sngQuantidade, "##,###,##0.000")
      Else
        txtDifQtdeXValor.Text = Format(curValorTotal - curValorUnit * sngQuantidade, FORMAT_VALUE)
      End If
    End If
     
  End If
  txtQuantidade.Text = sngQuantidade
  
End Sub

Private Sub txtValorTotal_GotFocus()
  Call SelectAllText(Me.ActiveControl)
  cmdOK.Default = True
End Sub

Private Sub txtValorTotal_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtValorTotal_LostFocus()
  cmdOK.Default = False
End Sub
