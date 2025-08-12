VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmOperador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleção do Operador"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "Operador.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   855
      Width           =   1100
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1410
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   315
      Width           =   1100
   End
   Begin VB.Data datCaixa 
      Caption         =   "Caixa"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Caixas em Uso"
      Top             =   3135
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Data datOperator 
      Caption         =   "Operador"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Funcionários ORDER BY Nome"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1665
   End
   Begin SSDataWidgets_B.SSDBCombo cboOperador 
      Bindings        =   "Operador.frx":08CA
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   735
      DataFieldList   =   "Código"
      MaxDropDownItems=   16
      _Version        =   196617
      BackColorOdd    =   16763025
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   1508
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Código"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   4260
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Código"
   End
   Begin SSDataWidgets_B.SSDBCombo cboCaixa 
      Bindings        =   "Operador.frx":08E4
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   735
      DataFieldList   =   "Caixa"
      MaxDropDownItems=   16
      _Version        =   196617
      BackColorOdd    =   16763025
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   1323
      Columns(0).Caption=   "Caixa"
      Columns(0).Name =   "Caixa"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Caixa"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   256
      Columns(1).Width=   5583
      Columns(1).Caption=   "Descrição"
      Columns(1).Name =   "Descrição"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Descrição"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Caixa"
   End
   Begin VB.Label lblCaixa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblOperador 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Caixa:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Operador:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "frmOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gsCodOper As String
Private gsCaixa As String
Private gsPassword As String

Private Sub cboOperador_GotFocus()
  cboOperador.BackColor = RGB(0, 0, 210)
  cboOperador.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub cboOperador_LostFocus()
  If cboOperador.Text <> "" Then
    datOperator.Recordset.FindFirst "Código = " & Val(cboOperador.Text)
    If Not datOperator.Recordset.NoMatch Then
      cboOperador.Text = datOperator.Recordset("Código")
      lblOperador.Caption = datOperator.Recordset("Nome")
      gsCodOper = datOperator.Recordset("Código")
    Else
      lblOperador.Caption = ""
    End If
  Else
    lblOperador.Caption = ""
  End If
  cboOperador.BackColor = vbWindowBackground
  cboOperador.ForeColor = vbWindowText
End Sub

Private Sub cboCaixa_GotFocus()
  cboCaixa.BackColor = RGB(0, 0, 210)
  cboCaixa.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub cboCaixa_LostFocus()
  If cboCaixa.Text <> "" Then
    datCaixa.Recordset.FindFirst "Caixa = " & Val(cboCaixa.Text)
    If Not datCaixa.Recordset.NoMatch Then
      cboCaixa.Text = datCaixa.Recordset("Caixa")
      lblCaixa.Caption = datCaixa.Recordset("Descrição")
      gsCaixa = datCaixa.Recordset("Caixa")
    Else
      lblCaixa.Caption = ""
    End If
  Else
    lblCaixa.Caption = ""
  End If
  cboCaixa.BackColor = vbWindowBackground
  cboCaixa.ForeColor = vbWindowText
End Sub

Private Sub cboCaixa_Click()
  gsCaixa = cboCaixa.Columns("Caixa").Text
End Sub

Private Sub cboCaixa_CloseUp()
  gsCaixa = cboCaixa.Columns("Caixa").Text
  cboCaixa.Text = cboCaixa.Columns("Caixa").Text
  lblCaixa.Caption = cboCaixa.Columns("Descrição").Text
End Sub

Private Sub cboOperador_Click()
  gsCodOper = cboOperador.Columns("Código").Text
End Sub

Private Sub cboOperador_CloseUp()
  gsCodOper = cboOperador.Columns("Código").Text
  cboOperador.Text = cboOperador.Columns("Código").Text
  lblOperador.Caption = cboOperador.Columns("Nome").Text
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFechar_Click()
  Dim rs As Recordset
  
  Call cboOperador_LostFocus
  Call cboCaixa_LostFocus
    
  If Val(gsCodOper) = 0 Or lblOperador.Caption = "" Then
    Call DisplayMsg("Operador incorreto. Selecione um operador na lista.")
    cboOperador.SetFocus
    Exit Sub
  End If
  
  If Val(gsCaixa) = 0 Or lblCaixa.Caption = "" Then
    Call DisplayMsg("Caixa incorreto. Selecione um caixa na lista.")
    cboCaixa.SetFocus
    Exit Sub
  End If
  
  gsPassword = ""
  'Set rs = datOperator.Recordset.Clone
  Set rs = db.OpenRecordset("SELECT * FROM [Funcionários] ORDER BY Código", dbOpenSnapshot)
  With rs
    .FindFirst "Código = " & gsCodOper
    If Not .NoMatch Then
      If .Fields("Liberado").Value = True Then
        If .Fields("Superusuário").Value = False Then
          If .Fields("Filial Acesso").Value <> 0 And .Fields("Filial Acesso").Value <> gnCodFilial Then
            Call DisplayMsg("Usuário não autorizado para esta Filial.")
            cboOperador.SetFocus
            Exit Sub
          End If
        End If
      End If
      gsPassword = .Fields("ValorP")
    End If
    .Close
  End With
  Set rs = Nothing
  
  If CriptografaSenha(txtPassword.Text) <> gsPassword Then
    Call DisplayMsg("Senha incorreta.")
    txtPassword.SetFocus
  Else
    gnCaixa = Val(gsCaixa)
    gnCodOper = Val(gsCodOper)
    frmMain.stbStatusBar.Panels("pnlUSER").Text = "Oper: " & gsCodOper & "-" & UCase(lblOperador.Caption)
    frmMain.stbStatusBar.Panels("pnlCAIXA").Text = "Caixa: " & gsCaixa
    Unload Me
  End If

End Sub

Private Sub Form_Load()
  Dim sSql As String
  Call CenterForm(Me)
  Call ShowMsg(LimparBarraStatus)
  datOperator.DatabaseName = gsQuickDBFileName
  datCaixa.DatabaseName = gsQuickDBFileName
  'Funcionários
  sSql = "SELECT Código, Nome, ValorP FROM Funcionários WHERE Liberado = True ORDER BY Nome"
  Set datOperator.Recordset = db.OpenRecordset(sSql, dbOpenSnapshot)
  datOperator.Refresh
  'Caixas
  sSql = "SELECT Caixa, Descrição FROM [Caixas em Uso] ORDER BY Caixa"
  Set datCaixa.Recordset = db.OpenRecordset(sSql, dbOpenSnapshot)
  datCaixa.Refresh
  'Insere os valores iniciais
  cboOperador.Text = gnCodOper
  Call cboOperador_LostFocus
  cboCaixa.Text = gnCaixa
  Call cboCaixa_LostFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call ShowMsg(LimparBarraStatus)
  Set datOperator.Recordset = Nothing
  Set datCaixa.Recordset = Nothing
End Sub

Private Sub txtPassword_Change()
  gsPassword = txtPassword.Text
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.BackColor = RGB(0, 0, 210)
  txtPassword.ForeColor = RGB(255, 255, 255)
  txtPassword.SelStart = 0
  txtPassword.SelLength = txtPassword.MaxLength
End Sub

Private Sub txtPassword_LostFocus()
  txtPassword.BackColor = vbWindowBackground
  txtPassword.ForeColor = vbWindowText
End Sub
