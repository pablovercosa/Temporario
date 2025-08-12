VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmParametrosDevolucaoMateriais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolução de Materiais"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParametrosDevolucaoMateriais.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   7455
   Begin VB.Data datParametros 
      Caption         =   "datParametros"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Data datTabela 
      Caption         =   "datTabela"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Data datCaixa 
      Caption         =   "datCaixa"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Caixa, Descrição FROM [Caixas em Uso] ORDER BY Caixa"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data datOperacao 
      Caption         =   "datOperacao"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome, Tipo FROM [Operações Saída] ORDER BY Código"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Frame fraY 
      Caption         =   "Valores Padrões para cada Saída gerada na Devolução de Materiais"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtNomeOperacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   930
         Width           =   4600
      End
      Begin VB.TextBox txtNomeCaixa 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2300
      End
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   4600
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperacao 
         Bindings        =   "frmParametrosDevolucaoMateriais.frx":058A
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   935
         Width           =   855
         DataFieldList   =   "Código"
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
         Columns(0).Width=   3200
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   8454143
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCaixa 
         Bindings        =   "frmParametrosDevolucaoMateriais.frx":05A4
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1390
         Width           =   855
         DataFieldList   =   "Caixa"
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
         Columns(0).Width=   3200
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   8454143
         DataFieldToDisplay=   "Caixa"
      End
      Begin SSDataWidgets_B.SSDBCombo cboTabela 
         Bindings        =   "frmParametrosDevolucaoMateriais.frx":05BB
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1845
         Width           =   1575
         DataFieldList   =   "Tabela"
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
         Columns(0).Width=   3200
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   8454143
         DataFieldToDisplay=   "Tabela"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmParametrosDevolucaoMateriais.frx":05D3
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   855
         DataFieldList   =   "Filial"
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
         Columns(0).Width=   3200
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   8454143
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operação"
         Height          =   195
         Left            =   315
         TabIndex        =   11
         Top             =   995
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Caixa"
         Height          =   195
         Left            =   615
         TabIndex        =   10
         Top             =   1450
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tabela"
         Height          =   195
         Left            =   540
         TabIndex        =   9
         Top             =   1905
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   540
         Width           =   300
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   6840
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Bands           =   "frmParametrosDevolucaoMateriais.frx":05EF
      End
   End
End
Attribute VB_Name = "frmParametrosDevolucaoMateriais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro    As Variant
Dim rstParamDevoMat As Recordset

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datParametros.DatabaseName = gsQuickDBFileName
  datOperacao.DatabaseName = gsQuickDBFileName
  datCaixa.DatabaseName = gsQuickDBFileName
  datTabela.DatabaseName = gsQuickDBFileName
  
  Set rstParamDevoMat = db.OpenRecordset("SELECT * FROM ParamDevoMat ORDER BY Filial", dbOpenDynaset)
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
  If KeyAscii = 13 Then
     SendKeys "{Tab}"
     KeyAscii = 0
 End If
End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstParametros As Recordset
  
  txtNomeFilial.Text = ""

  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstParametros = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & CByte(cboFilial.Text), dbOpenSnapshot)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome").Value & ""
    End If
    .Close
  End With

  Set rstParametros = Nothing

  '11/08/2004 - Implementado rotina
  'Esta private irá verificar se existe o registro na base
  'caso exista mostrará
  If Len(txtNomeFilial.Text) > 0 Then Call VerificarSeExisteRegistro

End Sub

Private Sub VerificarSeExisteRegistro()
  Dim rstParamDevoMat As Recordset
  Dim strQuery        As String
  
  strQuery = "SELECT * FROM ParamDevoMat "
  strQuery = strQuery & " WHERE Filial = " & CByte(cboFilial.Text)
  
  Set rstParamDevoMat = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  If rstParamDevoMat.RecordCount = 0 Then
    rstParamDevoMat.Close
    Set rstParamDevoMat = Nothing
    Exit Sub
  End If
  
  With rstParamDevoMat
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      cboOperacao.Text = .Fields("Operacao").Value
      cboOperacao_LostFocus
      cboCaixa.Text = .Fields("Caixa").Value
      cboCaixa_LostFocus
      cboTabela.Text = .Fields("Tabela").Value & ""
      
    End If
    .Close
  End With

  Set rstParamDevoMat = Nothing
  
End Sub

Private Sub cboOperacao_CloseUp()
  cboOperacao.Text = cboOperacao.Columns(0).Text
  cboOperacao_LostFocus
End Sub

Private Sub cboOperacao_LostFocus()
  Dim rstOperacao As Recordset
  
  txtNomeOperacao.Text = ""
    
  If Not IsNumeric(cboOperacao.Text) Then Exit Sub
  
  Set rstOperacao = db.OpenRecordset("SELECT Código, Nome FROM [Operações Saída] WHERE Código = " & CInt(cboOperacao.Text), dbOpenSnapshot)
  
  With rstOperacao
    If Not (.BOF And .EOF) Then
      txtNomeOperacao.Text = .Fields("Nome").Value & ""
    End If
    .Close
  End With

  Set rstOperacao = Nothing

End Sub

Private Sub cboCaixa_CloseUp()
  cboCaixa.Text = cboCaixa.Columns(0).Text
  cboCaixa_LostFocus
End Sub

Private Sub cboCaixa_LostFocus()
  Dim rstCaixas As Recordset
  
  txtNomeCaixa.Text = ""
    
  If Not IsNumeric(cboCaixa.Text) Then Exit Sub
  
  Set rstCaixas = db.OpenRecordset("SELECT Caixa, Descrição FROM [Caixas em Uso] WHERE Caixa = " & CByte(cboCaixa.Text), dbOpenSnapshot)
  
  With rstCaixas
    If Not (.BOF And .EOF) Then
      txtNomeCaixa.Text = .Fields("Descrição").Value & ""
    End If
    .Close
  End With

  Set rstCaixas = Nothing

End Sub

Private Function ValidarCampos() As Boolean
  
  If Len(txtNomeFilial.Text) <= 0 Then
    ValidarCampos = True
    MsgBox "Filial inválida, verifique.", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    Exit Function
  End If
  
  If Len(txtNomeOperacao.Text) <= 0 Then
    ValidarCampos = True
    MsgBox "Operação de Saída inválida, verifique.", vbExclamation, "Quick Store"
    cboOperacao.SetFocus
    Exit Function
  End If
  
  If Len(txtNomeCaixa.Text) <= 0 Then
    ValidarCampos = True
    MsgBox "Caixa inválido, verifique.", vbExclamation, "Quick Store"
    cboCaixa.SetFocus
    Exit Function
  End If
  
  If Len(cboTabela.Text) <= 0 Then
    ValidarCampos = True
    MsgBox "Tabela inválida, verifique.", vbExclamation, "Quick Store"
    cboTabela.SetFocus
    Exit Function
  End If
  
End Function

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstParamDevoMat
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  
End Sub

Private Sub MoveLast()
  On Error Resume Next
  
  With rstParamDevoMat
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  
  With rstParamDevoMat
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With

End Sub

Private Sub MoveNext()
  On Error Resume Next
  
  With rstParamDevoMat
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With

End Sub

Private Sub DeleteRecord()
  Dim intResposta  As Integer
  Dim dblAuxCodigo As Double
  Dim strAuxStr    As String
  
  If IsNull(Num_Registro) Then
    Beep
    MsgBox "Não existe nenhum Parâmetro para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar este Parâmetro ? "
  intResposta = MsgBox(strAuxStr, 20, "ATENÇÃO!!")
  If intResposta = 6 Then
    rstParamDevoMat.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim intErro As Integer
  
  On Error GoTo Processa_Erro

  If ValidarCampos Then Exit Sub

  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstParamDevoMat
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Filial").Value = CByte(cboFilial.Text)
     Else
        .Edit
     End If
     
     .Fields("Operacao").Value = CInt(cboOperacao.Text)
     .Fields("Caixa").Value = CByte(cboCaixa.Text)
     .Fields("Tabela").Value = cboTabela.Text

     .Update
     Num_Registro = .LastModified
     .Bookmark = Num_Registro
   End With

  Call StatusMsg("")
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar registro."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & " - " & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  
  cboFilial.Text = ""
  txtNomeFilial.Text = ""
  cboOperacao.Text = ""
  txtNomeOperacao.Text = ""
  cboCaixa.Text = ""
  txtNomeCaixa.Text = ""
  cboTabela.Text = ""
    
  cboFilial.SetFocus
  Num_Registro = Null
  
  If Not rstParamDevoMat.EOF Then
    On Error Resume Next
    rstParamDevoMat.MoveFirst
    rstParamDevoMat.MovePrevious
    On Error GoTo 0
  End If

End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstParamDevoMat.Close
  Set rstParamDevoMat = Nothing
End Sub

Sub ShowRecord()
  With rstParamDevoMat
    cboFilial.Text = .Fields("Filial").Value
    cboFilial_LostFocus
    cboOperacao.Text = .Fields("Operacao").Value
    cboOperacao_LostFocus
    cboCaixa.Text = .Fields("Caixa").Value
    cboCaixa_LostFocus
    cboTabela.Text = .Fields("Tabela").Value
    
    Num_Registro = .Bookmark
  End With

End Sub

