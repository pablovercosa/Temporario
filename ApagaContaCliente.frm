VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmApagaContaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Apaga Contas de Clientes"
   ClientHeight    =   1980
   ClientLeft      =   1875
   ClientTop       =   1830
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ApagaContaCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   7365
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
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   180
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton B_Apaga 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apagar"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
      Width           =   7125
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   345
      Left            =   5460
      TabIndex        =   1
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   690
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "ApagaContaCliente.frx":4E95A
      DataSource      =   "Data1"
      Height          =   375
      Left            =   750
      TabIndex        =   0
      ToolTipText     =   "Use 0 para todos os clientes"
      Top             =   120
      Width           =   1035
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
      Columns(0).Width=   9313
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
      _ExtentX        =   1826
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Apagar todas as contas do(s) cliente (s) já recebidas, com data de recebimento até (inclusive) :"
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   5055
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1830
      TabIndex        =   4
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   195
      Width           =   555
   End
End
Attribute VB_Name = "frmApagaContaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCliFor As Recordset

Private Sub B_Apaga_Click()
  Dim Resposta As Integer
  Dim sSql As String

  Call StatusMsg("")

  If Not IsDate(Dia.Text) Then
    DisplayMsg "Data inválida, verifique."
    Exit Sub
  End If

  sSql = "Delete * From [Conta Cliente] Where Valor = [Valor Pago] And [Data Pagamento] <= DateValue('" + Dia.Text + "')"
  If Nome_Cliente.Caption <> "" Then
     sSql = sSql + " And Cliente = " + Combo_Cliente.Text
  End If

  db.Execute sSql

  'Efetua registro do Log
  db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
    Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '" & Left("Usu:" & gnUserCode & " Exc contas cli recebidas com DtVc até " & Dia.Text, 80) & "', 'CNT_REC: exc-DT ccli')", dbFailOnError

  DisplayMsg "Foram apagados " + str(db.RecordsAffected) + " lançamentos."

End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
 Nome_Cliente.Caption = ""
 If IsNull(Combo_Cliente.Text) Then Exit Sub
 If Combo_Cliente.Text = "" Then Exit Sub
 If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
 If Val(Combo_Cliente.Text) <= 0 Then Exit Sub
 If Val(Combo_Cliente.Text) > 99999999 Then Exit Sub
  
 rsCliFor.Index = "Código"
 rsCliFor.Seek "=", Val(Combo_Cliente.Text)
 If Not rsCliFor.NoMatch Then Nome_Cliente.Caption = rsCliFor("Nome")
 
End Sub

Private Sub Dia_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Dia.Text = frmCalendario.gsDateCalender(Dia.Text)
  End Select
End Sub

Private Sub Dia_LostFocus()
  Dia.Text = Ajusta_Data(Dia.Text)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Data1.DatabaseName = gsQuickDBFileName
End Sub
