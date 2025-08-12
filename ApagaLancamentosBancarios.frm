VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmApagaLancamentos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apagar Lançamentos"
   ClientHeight    =   3210
   ClientLeft      =   1440
   ClientTop       =   1725
   ClientWidth     =   6870
   ForeColor       =   &H80000008&
   Icon            =   "ApagaLancamentosBancarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton B_Apaga 
      Caption         =   "Apagar"
      Height          =   400
      Left            =   5385
      TabIndex        =   6
      Top             =   2655
      Width           =   1335
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   4920
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2160
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
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
      Bindings        =   "ApagaLancamentosBancarios.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
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
      Columns.Count   =   3
      Columns(0).Width=   6694
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3731
      Columns(1).Caption=   "Conta"
      Columns(1).Name =   "Conta"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Conta"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1455
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apaga lançamentos até (inclusive) :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   $"ApagaLancamentosBancarios.frx":059E
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Nome_Conta 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conta :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   1095
   End
End
Attribute VB_Name = "frmApagaLancamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Num_Registro As Variant
Dim Ordem As Variant
Dim rsLançamentos As Recordset
Dim rsContas As Recordset

Private Sub B_Apaga_Click()
  Dim Conta As Long
  Dim Ordem As Variant
  Dim Data_Aux As Variant
  Dim Fim As Integer
  
  Call StatusMsg("")

  If Nome_Conta.Caption = "" Then
    DisplayMsg "Conta incorreta, verifique."
    Combo_Conta.SetFocus
    Exit Sub
  End If
 
  If Not IsDate(Dia.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Dia.SetFocus
    Exit Sub
  End If

  Conta = 0
  Data_Aux = CDate("01/01/1980")
  Ordem = 0
  rsLançamentos.Index = "Conta"
  Fim = False

  Call ws.BeginTrans
    
  Do
   rsLançamentos.Seek ">", Val(Combo_Conta.Text), Data_Aux, Ordem
   If rsLançamentos.NoMatch Then Fim = True
   If Fim = False Then If rsLançamentos("Conta") <> Val(Combo_Conta.Text) Then Fim = True
   If Fim = False Then If rsLançamentos("Data") > CDate(Dia.Text) Then Fim = True
   If Fim = False Then
      Data_Aux = rsLançamentos("Data")
      Ordem = rsLançamentos("Ordem")
      rsLançamentos.Delete
      Call StatusMsg("Apagando dia " + str$(Data_Aux))
      Conta = Conta + 1
   End If
  Loop Until Fim = True

  Call ws.CommitTrans

  DisplayMsg "Processo terminado. Foram apagados " + str$(Conta) + " lançamentos."

  Exit Sub
  
ErrTansaction:
  Call ws.Rollback
  
ErrHandle:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar operação de apagar lançamentos."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Private Sub Combo_Conta_CloseUp()
  Combo_Conta.Text = Combo_Conta.Columns(2).Text
  Combo_Conta_LostFocus
End Sub

Private Sub Combo_Conta_LostFocus()
  Nome_Conta.Caption = ""
  If IsNull(Combo_Conta.Text) Then Exit Sub
  If Not IsNumeric(Combo_Conta.Text) Then Exit Sub
  If Val(Combo_Conta.Text) < 0 Or Val(Combo_Conta.Text) > 999999 Then Exit Sub

  rsContas.Index = "Código"
  rsContas.Seek "=", Val(Combo_Conta.Text)
  If rsContas.NoMatch Then Exit Sub
  Nome_Conta.Caption = rsContas("Descrição")
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
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  
  Data1.DatabaseName = gsQuickDBFileName

  Num_Registro = Null
 
End Sub
