VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWEB_OFFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procurar registro"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   Icon            =   "frmWEB_OFFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9645
   Begin VB.CommandButton cmdNewFind 
      Caption         =   "&Nova pesquisa"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtOrderID 
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox txtTraceCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin SSDataWidgets_B.SSDBGrid grdResult 
      Height          =   2055
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Dê um duplo-clique para exibir o registro"
      Top             =   3000
      Width           =   9255
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   6
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   2223
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   4048
      Columns(2).Caption=   "Nome do cliente"
      Columns(2).Name =   "ShopperName"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3387
      Columns(3).Caption=   "Status"
      Columns(3).Name =   "Passo"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   2434
      Columns(4).Caption=   "Valor Total"
      Columns(4).Name =   "ValorTotal"
      Columns(4).Alignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   3201
      Columns(5).Caption=   "Data"
      Columns(5).Name =   "Data"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   7
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      TabNavigation   =   1
      _ExtentX        =   16325
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Nenhum registro"
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
   Begin VB.Data datCliente 
      Caption         =   "Cliente"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Cli_For"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   7455
      Begin SSDataWidgets_B.SSDBCombo cboShopper 
         Bindings        =   "frmWEB_OFFind.frx":058A
         DataSource      =   "datCliente"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1695
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
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Codigo"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   3
         Columns(0).FieldLen=   256
         Columns(1).Width=   8864
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblShopperName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   300
         Width           =   5415
      End
   End
   Begin MSComCtl2.DTPicker dtpDataInicial 
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   96665601
      CurrentDate     =   37333
   End
   Begin MSComCtl2.DTPicker dtpDataFinal 
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   96665601
      CurrentDate     =   37333
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Pesquisar"
      Default         =   -1  'True
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtSequence 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9375
      _Version        =   65536
      _ExtentX        =   16536
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Selecione uma ou mais opções para utilizar na pesquisa dos registros"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Código chave identificador"
      Height          =   195
      Index           =   12
      Left            =   5040
      TabIndex        =   18
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Código de rastreamento"
      Height          =   195
      Index           =   14
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Width           =   1680
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   450
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   720
      Width           =   795
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frmWEB_OFFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboShopper_CloseUp()
  cboShopper.Text = cboShopper.Columns("Codigo").Text
  lblShopperName.Caption = cboShopper.Columns("Nome").Text
End Sub

Private Sub cboShopper_LostFocus()
  Dim lngRet As Long
  
  Call IsDataType(dtLong, cboShopper.Text, lngRet)
  If lngRet > 0 Then
    cboShopper.Text = lngRet
    lblShopperName.Caption = gstrGetCliForName(CLng(cboShopper.Text))
  Else
    cboShopper.Text = ""
    lblShopperName.Caption = ""
  End If
End Sub

Private Sub cmdFind_Click()
  Dim strX As String
  Dim strID As String
  Dim blnDataInicial As Boolean
  Dim blnDataFinal As Boolean
  Dim rsResult As Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim strShopperName As String
  
  With grdResult
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  'Passo
  Select Case cboStatus.ItemData(cboStatus.ListIndex)
    Case ofsAll
      strX = "Passo >= " & ofsReceived
    Case Else
      strX = "Passo = " & cboStatus.ItemData(cboStatus.ListIndex)
  End Select
  strWHERE = strX
  
  'Data
  blnDataInicial = Not IsNull(dtpDataInicial.Value)
  blnDataFinal = Not IsNull(dtpDataFinal.Value)
  
  If blnDataInicial Or blnDataFinal Then
    If blnDataInicial And blnDataFinal Then 'Entre datas
      strX = "Data BETWEEN #" & Format(dtpDataInicial.Value, SQL_DATE_MASK) & _
        "# AND #" & Format(dtpDataFinal.Value, SQL_DATE_MASK) & " 23:59:59#"
    ElseIf blnDataInicial Then
      strX = "Data >=#" & Format(dtpDataInicial.Value, SQL_DATE_MASK) & "#"
    ElseIf blnDataFinal Then
      strX = "Data <=#" & Format(dtpDataFinal.Value, SQL_DATE_MASK) & " 23:59:59#"
    End If
    strWHERE = strWHERE & " AND " & strX
  End If
  
  'Comprador (Cliente)
  Call cboShopper_LostFocus
  
  If lblShopperName.Caption <> "" Then
    Call WEB_GetShopperData(strID, CLng(cboShopper.Text), "")
    strX = "ShopperID = '" & strID & "'"
    strWHERE = strWHERE & " AND " & strX
  End If
  
  'Sequência
  If txtSequence.Text <> "" Then
    strX = "Sequencia = " & txtSequence.Text
    strWHERE = strWHERE & " AND " & strX
  End If
  
  'Código de rastreamento
  If txtTraceCode.Text <> "" Then
    strX = "TraceCode = '" & txtTraceCode.Text & "'"
    strWHERE = strWHERE & " AND " & strX
  End If
  
  'Código chave identificador
  If txtOrderID.Text <> "" Then
    strX = "OrderID = '" & txtOrderID.Text & "'"
    strWHERE = strWHERE & " AND " & strX
  End If
  
  strSQL = "SELECT * FROM WEB_OrderForms WHERE " & strWHERE & " ORDER BY ID"
  Set rsResult = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rsResult
    If Not .BOF And Not .EOF Then
      Do Until .EOF
        Call WEB_GetShopperData(.Fields("ShopperID").Value, 0, strShopperName)
        grdResult.AddItem .Fields("ID").Value & vbTab & _
                          .Fields("Boleto").Value & vbTab & _
                          strShopperName & vbTab & _
                          gstrWEB_GetDescPasso(.Fields("Passo").Value) & vbTab & _
                          .Fields("Total").Value & vbTab & _
                          .Fields("Data").Value
        .MoveNext
      Loop
      grdResult.Caption = "Resultado: " & .RecordCount & " registro(s)"
      grdResult.SetFocus
    Else
      grdResult.Caption = "Nenhum registro"
    End If
    .Close
  End With
  Set rsResult = Nothing
  
End Sub

Private Sub cmdNewFind_Click()
  
  cboStatus.ListIndex = 0
  dtpDataInicial.Value = Date
  dtpDataInicial.Value = Null
  dtpDataFinal.Value = Date
  dtpDataFinal.Value = Null
  cboShopper.Text = ""
  lblShopperName.Caption = ""
  txtSequence.Text = ""
  txtTraceCode.Text = ""
  txtOrderID.Text = ""
  
  With grdResult
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  cboStatus.SetFocus
  
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  With cboStatus
    .AddItem "Todos"
    .ItemData(.NewIndex) = ofsAll
    .AddItem "Pedido recebido"
    .ItemData(.NewIndex) = ofsReceived
    .AddItem "Pedido com pagamento confirmado"
    .ItemData(.NewIndex) = ofsConfirmedPayment
    .AddItem "Pedido com produto embalado"
    .ItemData(.NewIndex) = ofsPacked
    .AddItem "Pedido enviado"
    .ItemData(.NewIndex) = ofsHasSent
    .AddItem "Pedido cancelado"
    .ItemData(.NewIndex) = ofsCanceled
    .ListIndex = 0
  End With
  
  dtpDataInicial.Value = Date
  dtpDataInicial.Value = Null
  dtpDataFinal.Value = Date
  dtpDataFinal.Value = Null
  
  datCliente.DatabaseName = gsQuickDBFileName
  
End Sub

Private Sub grdResult_DblClick()
  Dim lngID As Long
  Dim varBook As Variant
  
  If grdResult.Rows = 0 Then Exit Sub
  
  varBook = grdResult.AddItemBookmark(grdResult.Row)
  lngID = CLng(grdResult.Columns("ID").CellText(varBook))
  
  Call frmWEB_OrderForms.PosRecordset(lngID)
End Sub

Private Sub txtSequence_Validate(Cancel As Boolean)
  Dim lngRet As Long
  
  Call IsDataType(dtLong, txtSequence.Text, lngRet)
  If lngRet > 0 Then
    txtSequence.Text = lngRet
  Else
    txtSequence.Text = ""
  End If
End Sub
