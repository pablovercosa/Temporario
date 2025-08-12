VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWEB_OFPersonalizeFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro Personalizado"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmWEB_OFPersonalizeFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   7455
      Begin SSDataWidgets_B.SSDBCombo cboShopper 
         Bindings        =   "frmWEB_OFPersonalizeFilter.frx":058A
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "Selecione uma ou mais opções para utilizar como filtro na seleção dos registros"
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3375
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   795
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   450
   End
End
Attribute VB_Name = "frmWEB_OFPersonalizeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSucess As Boolean
Private mstrFilter As String

Public Function Personalize(ByRef strFilter As String) As Boolean
  mblnSucess = False
  mstrFilter = ""
  Me.Show vbModal
  Personalize = mblnSucess
  strFilter = mstrFilter
  Set frmWEB_OFPersonalizeFilter = Nothing
End Function

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

Private Sub cmdCancel_Click()
  mblnSucess = False
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim strX As String
  Dim strID As String
  Dim blnDataInicial As Boolean
  Dim blnDataFinal As Boolean
  
  'Passo
  Select Case cboStatus.ItemData(cboStatus.ListIndex)
    Case ofsAll
      strX = "Passo >= " & ofsReceived
    Case Else
      strX = "Passo = " & cboStatus.ItemData(cboStatus.ListIndex)
  End Select
  mstrFilter = strX
  
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
    mstrFilter = mstrFilter & " AND " & strX
  End If
  
  'Comprador (Cliente)
  Call cboShopper_LostFocus
  
  If lblShopperName.Caption <> "" Then
    Call WEB_GetShopperData(strID, CLng(cboShopper.Text), "")
    strX = "ShopperID = '" & strID & "'"
    mstrFilter = mstrFilter & " AND " & strX
  End If
  
  mblnSucess = True
  Unload Me
  
End Sub

Private Sub Form_Load()
  
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
