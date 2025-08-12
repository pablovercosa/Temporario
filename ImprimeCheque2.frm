VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmImprimeCheque2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Cheques do Estabelecimento"
   ClientHeight    =   2355
   ClientLeft      =   3450
   ClientTop       =   2925
   ClientWidth     =   6015
   Icon            =   "ImprimeCheque2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   6015
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   3465
      TabIndex        =   9
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1365
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Bancos"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   4575
      TabIndex        =   3
      Top             =   1875
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   1365
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   12
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Favorecido 
      Height          =   315
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   4635
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
      Bindings        =   "ImprimeCheque2.frx":058A
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   315
      Width           =   960
      ListAutoValidate=   0   'False
      MaxDropDownItems=   16
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
      Columns.Count   =   2
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1720
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1693
      _ExtentY        =   556
      _StockProps     =   93
      Text            =   "0"
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Código"
   End
   Begin VB.Label Label2 
      Caption         =   "Data :"
      Height          =   225
      Left            =   2730
      TabIndex        =   8
      Top             =   1425
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Valor :"
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   1425
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Favorecido :"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   885
      Width           =   960
   End
   Begin VB.Label Nome_Banco 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2415
      TabIndex        =   5
      Top             =   315
      Width           =   3480
   End
   Begin VB.Label Label1 
      Caption         =   "Banco :"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   315
      Width           =   645
   End
End
Attribute VB_Name = "frmImprimeCheque2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsBancos As Recordset
Dim rsParametros As Recordset

Private Sub B_Cancela_Click()
  Unload Me
End Sub

Private Sub B_Imprime_Click()

  If Nome_Banco.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Informe o banco."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Valor.Text = gsHandleNull(Valor.Text)
      
  If CDbl(Valor.Text) <= 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor incorreto."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If Not IsDate(Data.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Informe a data."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  Call Imprime_Cheque(Favorecido.Text, Combo_Banco.Text, Format(Data.Text, "dd/mm/yy"), Valor.Text)
  
End Sub

Private Sub Combo_Banco_CloseUp()
  Combo_Banco.Text = Combo_Banco.Columns(1).Text
  Combo_Banco_LostFocus
End Sub

Private Sub Combo_Banco_LostFocus()
  Nome_Banco.Caption = ""
  If IsNull(Combo_Banco.Text) Then Exit Sub
  If Combo_Banco.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Banco.Text) Then Exit Sub
  If Val(Combo_Banco.Text) < 1 Then Exit Sub
  If Val(Combo_Banco.Text) > 999 Then Exit Sub
  
  rsBancos.Index = "Código"
  rsBancos.Seek "=", Combo_Banco.Text
  If rsBancos.NoMatch Then Exit Sub
  Nome_Banco.Caption = rsBancos("Nome")
End Sub

Private Sub Data_LostFocus()
  Data.Text = Ajusta_Data(Data.Text)
End Sub

Private Sub Data_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data.Text = frmCalendario.gsDateCalender(Data.Text)
  End Select
End Sub


Private Sub Form_Activate()

  Data.Text = Format(Date, "dd/mm/yyyy")
 
  If Combo_Banco.Text = "0" Then
    Combo_Banco.Text = CStr(rsParametros("Código Banco Cheques"))
  End If

End Sub

Private Sub Form_Load()
  
  Set rsBancos = db.OpenRecordset("Bancos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If Not rsParametros.NoMatch Then
    Combo_Banco.Enabled = True
    Combo_Banco.Text = rsParametros("Código Banco Cheques")
    If Val(Combo_Banco.Text) <> 0 Then
      Combo_Banco.Enabled = False
    End If
    Combo_Banco_LostFocus
  End If
  
  Data1.DatabaseName = gsQuickDBFileName
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsBancos.Close
  rsParametros.Close
  Set rsBancos = Nothing
  Set rsParametros = Nothing

End Sub
