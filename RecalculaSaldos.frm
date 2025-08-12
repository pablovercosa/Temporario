VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmRecalcula 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recálculo de Saldo de Conta"
   ClientHeight    =   1695
   ClientLeft      =   1650
   ClientTop       =   2460
   ClientWidth     =   5760
   ForeColor       =   &H80000008&
   HelpContextID   =   1260
   Icon            =   "RecalculaSaldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   5760
   Begin VB.CommandButton B_Calcula 
      Caption         =   "Recalcular"
      Height          =   400
      Left            =   4290
      TabIndex        =   3
      ToolTipText     =   "Recalcula o saldo de uma conta quando movimentos retroativos foram lançados"
      Top             =   1155
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
      Left            =   225
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contas Bancárias"
      Top             =   1260
      Visible         =   0   'False
      Width           =   2055
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
      Bindings        =   "RecalculaSaldos.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   930
      TabIndex        =   2
      Top             =   405
      Width           =   855
      DataFieldList   =   "Conta"
      _Version        =   196617
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3731
      Columns(0).Caption=   "Conta"
      Columns(0).Name =   "Conta"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Conta"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6588
      Columns(1).Caption=   "Descrição"
      Columns(1).Name =   "Descrição"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Descrição"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4260
      Columns(2).Caption=   "Agência"
      Columns(2).Name =   "Agência"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Agência"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2540
      Columns(3).Caption=   "Código"
      Columns(3).Name =   "Código"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Código"
      Columns(3).DataType=   2
      Columns(3).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Conta 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   405
      Width           =   3765
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conta:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   465
      Width           =   555
   End
End
Attribute VB_Name = "frmRecalcula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Num_Registro As Variant
Dim Ordem As Variant
Dim rsContas As Recordset

Private Sub B_Calcula_Click()
  Dim nConta As Long
  
  Call StatusMsg("")

  If Nome_Conta.Caption = "" Then
    DisplayMsg "Conta incorreta, verifique."
    Combo_Conta.SetFocus
    Exit Sub
  End If
 
  nConta = gnAtualizaSaldoBancario(Val(Combo_Conta.Text))
  
  DisplayMsg "Processo terminado. Foram atualizados " + CStr(nConta) + " lançamentos."
  
  Call StatusMsg("")

End Sub

Private Sub Combo_Conta_CloseUp()
  Combo_Conta.Text = Combo_Conta.Columns(3).Text
  Nome_Conta.Caption = Combo_Conta.Columns(1).Text
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

Private Sub Command3D2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)
  Data1.DatabaseName = gsQuickDBFileName

  Num_Registro = Null
End Sub
