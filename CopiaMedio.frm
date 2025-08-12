VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosCopiaCustoMedio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia Custo Médio para Tabela de Preços"
   ClientHeight    =   1845
   ClientLeft      =   1560
   ClientTop       =   2025
   ClientWidth     =   5850
   HelpContextID   =   1660
   Icon            =   "CopiaMedio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1845
   ScaleWidth      =   5850
   Begin VB.CheckBox chkContaClientes 
      Caption         =   "&Refletir alteração também na Conta de Clientes"
      Height          =   345
      Left            =   105
      TabIndex        =   4
      Top             =   1275
      Width           =   3690
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "CopiaMedio.frx":058A
      Height          =   315
      Left            =   3750
      TabIndex        =   0
      Top             =   165
      Width           =   1935
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Data datPrecos 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   525
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   2220
      Width           =   1875
   End
   Begin VB.CommandButton B_Copia 
      Caption         =   "Copiar"
      Height          =   400
      Left            =   4335
      TabIndex        =   2
      Top             =   1275
      Width           =   1335
   End
   Begin VB.CheckBox Sobrepor 
      Caption         =   "Sobrepor preços existentes"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Copiar custo médio para a seguinte tabela :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   195
      Width           =   3255
   End
End
Attribute VB_Name = "frmPrecosCopiaCustoMedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsProdutos As Recordset
Dim rsPreços As Recordset
Dim rsTabelas As Recordset
Private rsConta_Cli As Recordset

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Copia_Click()
  Dim Cód As String
  Dim Copiados As Long
  Dim nTempCopiados As Long
  Dim nNewValue As Double
  
  Call StatusMsg("")

  If IsNull(cboLista.Text) Then
    DisplayMsg "Tabela incorreta."
    cboLista.SetFocus
    Exit Sub
  End If
  If cboLista.Text = "" Then
    DisplayMsg "Tabela incorreta."
    cboLista.SetFocus
    Exit Sub
  End If
  
  Copiados = 0
  Cód = ""
  rsProdutos.Index = "Código"
  rsPreços.Index = "Tabela"
  
  On Error GoTo ErrTrans
  Call ws.BeginTrans
  
Lp1:
  rsProdutos.Seek ">", Cód
  If rsProdutos.NoMatch Then GoTo Fim
  Cód = rsProdutos("Código")
  
  rsPreços.Seek "=", cboLista.Text, Cód
  If rsPreços.NoMatch Then
    rsPreços.AddNew
    rsPreços("Tabela") = cboLista.Text
    rsPreços("Produto") = Cód
    rsPreços("Preço") = rsProdutos("Custo Médio")
    rsPreços("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsPreços.Update
    If chkContaClientes.Value = vbChecked Then
      nNewValue = rsProdutos("Custo Médio")
      Call UpdateContaClientes(cboLista.Text, rsProdutos("Código").Value, nNewValue)
    End If
    'Atualiza o sincronismo para o produto WEB alterado
    Call WEB_SynchronizeProduct(Cód)
    Copiados = Copiados + 1
  Else
    If Sobrepor.Value = 1 Then
      rsPreços.Edit
      rsPreços("Preço") = rsProdutos("Custo Médio")
      rsPreços("Data Alteração") = Format(Date, "dd/mm/yyyy")
      rsPreços.Update
      If chkContaClientes.Value = vbChecked Then
        nNewValue = rsProdutos("Custo Médio")
        Call UpdateContaClientes(cboLista.Text, rsProdutos("Código").Value, nNewValue)
      End If
      Copiados = Copiados + 1
    End If
  End If
  
  If nTempCopiados <> Copiados Then
    nTempCopiados = Copiados
    Call StatusMsg("Foram copiados " & Copiados & " registros.")
  End If
  GoTo Lp1
  
Fim:
  
  'Cria configuração da tabela
  Call CheckConfigTablePrice(cboLista.Text)

  Call ws.CommitTrans
  
  datPrecos.Refresh
  cboLista.Refresh
  
  DisplayMsg "Final de processos, registros copiados : " & Copiados
    
  Exit Sub
  
ErrTrans:
  ws.Rollback
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao copiar valores entre tabelas."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub

Private Sub cboLista_KeyPress(KeyAscii As Integer)
  KeyAscii = gnLimitKeyPress(cboLista, 15, KeyAscii)
  If KeyAscii <> 0 Then
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub chkContaClientes_Click()
  If chkContaClientes.Value = vbChecked Then
    If Not frmGerente.gbSenhaGerente Then
      chkContaClientes.Value = vbUnchecked
      Exit Sub
    End If
  End If
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsTabelas = db.OpenRecordset("Tabela de Preços")
  
  datPrecos.DatabaseName = gsQuickDBFileName
  
End Sub

Private Sub cboLista_LostFocus()
 If IsNull(cboLista.Text) Then Exit Sub
 cboLista.Text = UCase(cboLista.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsProdutos.Close
  rsPreços.Close
  rsTabelas.Close
  Set rsProdutos = Nothing
  Set rsPreços = Nothing
  Set rsTabelas = Nothing
End Sub
