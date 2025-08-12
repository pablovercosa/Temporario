VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmApagaCaixa 
   Caption         =   "Apaga lançamentos do Caixa"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "frmApagaCaixa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   5265
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   480
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton B_Apaga 
      Caption         =   "Apagar"
      Height          =   375
      Left            =   4005
      TabIndex        =   3
      Top             =   2640
      Width           =   1140
   End
   Begin Threed.SSPanel Mensagem 
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   3120
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   15
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1080
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd//mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
      Bindings        =   "frmApagaCaixa.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   450
      Width           =   870
      DataFieldList   =   "Descrição"
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
      Columns(0).Width=   9419
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1217
      Columns(1).Caption=   "Caixa"
      Columns(1).Name =   "Caixa"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Caixa"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1535
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label4 
      Caption         =   "Caixa :"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Nome_Caixa 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   3225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Filial :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1050
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Apagar os lançamentos no caixa indicado acima até a data (inclusive)"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   $"frmApagaCaixa.frx":059E
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
End
Attribute VB_Name = "frmApagaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsCaixas As Recordset


Private Sub B_Apaga_Click()
  
  Dim Resposta As Integer
  Dim sSql As String
  
  Call StatusMsg("")
  
  If Not IsDate(Dia.Text) Then
  DisplayMsg "Data inválida, verifique."
  Exit Sub
  End If
  
  'para corrigir bug de versão anterior apaga lançamentos do caixa 0
  
  sSql = "Delete * From [Caixa] Where Caixa = " & 0
  db.Execute sSql
  
  sSql = ""
  sSql = "Delete * From [Caixa] Where Caixa = " & str(Val(Combo_Caixa.Text))
  sSql = sSql & "And Data <= DateValue('" & Dia.Text & "')"
  sSql = sSql & " And Filial = " & str(gnCodFilial)
  db.Execute sSql
  
  DisplayMsg "Foram apagados " & str(db.RecordsAffected) & " lançamentos."
End Sub

Private Sub Combo_Caixa_CloseUp()
  
  Combo_Caixa.Text = Combo_Caixa.Columns(1).Text
  Combo_Caixa_LostFocus
  
End Sub

Private Sub Combo_Caixa_LostFocus()
  
  Nome_Caixa.Caption = ""
  If IsNull(Combo_Caixa.Text) Then Exit Sub
  If Combo_Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Caixa.Text) Then Exit Sub
  If Val(Combo_Caixa.Text) < 0 Then Exit Sub
  If Val(Combo_Caixa.Text) > 99 Then Exit Sub
  
  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", Val(Combo_Caixa.Text)
  If rsCaixas.NoMatch Then Exit Sub
  Nome_Caixa.Caption = rsCaixas("Descrição")

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

Private Sub Form_Load()

  Call CenterForm(Me)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  Nome_Filial.Caption = gnCodFilial & "-" & rsParametros("Nome")
  
  Data2.DatabaseName = gsQuickDBFileName
'  Dia.Text = gsFormatDate(Data_Atual)
  
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If

End Sub








