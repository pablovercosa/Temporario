VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Begin VB.Form frmRetencoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Retenções"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRetencoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   5115
   Begin VB.Frame fraX 
      Height          =   3255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5055
      Begin VB.Frame fraT 
         Caption         =   "Tipo"
         Height          =   1560
         Left            =   2640
         TabIndex        =   13
         Top             =   1100
         Width           =   2175
         Begin VB.OptionButton optCarneComJuros 
            Caption         =   "Carnê com Juros"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton optCarneSemJuros 
            Caption         =   "Carnê sem Juros"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   915
            Width           =   1575
         End
         Begin VB.OptionButton optChequeComJuros 
            Caption         =   "Cheque com Juros"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   645
            Width           =   1695
         End
         Begin VB.OptionButton optChequeSemJuros 
            Caption         =   "Cheque sem Juros"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.ComboBox cboNomeFinanceira 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmRetencoes.frx":058A
         Left            =   1200
         List            =   "frmRetencoes.frx":0594
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   3615
      End
      Begin VB.TextBox Código 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "Pressione F5 para buscar o próximo código livre."
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtValorRetencao 
         Height          =   315
         Left            =   1185
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   1620
         Width           =   1020
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   240
         Top             =   2280
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
         Bands           =   "frmRetencoes.frx":05AB
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Financeira"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Retenção"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmRetencoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rstRetencao  As Recordset
'Nota: O Text Box do Código teve que ser chamado de Código ao invés de txtCodigo
'para podermos aproveitar a rotina em QSGeral de busca automática do próximo código livre

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

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rstRetencao = db.OpenRecordset("SELECT * FROM Retencao ORDER BY Código", dbOpenDynaset)
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
  
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstRetencao
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
  
  With rstRetencao
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
  
  With rstRetencao
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
  
  With rstRetencao
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
    MsgBox "Não existe nenhuma Retenção para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar esta Retenção ? "
  intResposta = MsgBox(strAuxStr, 20, "ATENÇÃO!!")
  If intResposta = 6 Then
    rstRetencao.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  
  On Error GoTo Processa_Erro
  
  If ValidaCampos Then Exit Sub
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstRetencao
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Código") = CInt(Código.Text)
     Else
       .Edit
     End If
     .Fields("Nome") = txtNome.Text & ""
     .Fields("NomeDaFinanceira").Value = Trim(cboNomeFinanceira.Text) & ""
     .Fields("ValorRetencao").Value = Format(txtValorRetencao.Text, FORMAT_VALUE)

     If optChequeSemJuros.Value Then .Fields("Tipo").Value = "Cheque sem Juros"
     If optChequeComJuros.Value Then .Fields("Tipo").Value = "Cheque com Juros"
     If optCarneSemJuros.Value Then .Fields("Tipo").Value = "Carne sem Juros"
     If optCarneComJuros.Value Then .Fields("Tipo").Value = "Carne com Juros"


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
  
  Código.Text = ""
  txtNome.Text = ""
  cboNomeFinanceira.Text = ""
  txtValorRetencao.Text = "0,00"
  optChequeSemJuros.Value = True
  
  Código.SetFocus
  Num_Registro = Null
  
  If Not rstRetencao.EOF Then
    On Error Resume Next
    rstRetencao.MoveFirst
    rstRetencao.MovePrevious
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
  rstRetencao.Close
  Set rstRetencao = Nothing
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyF5 Then
    Set rstRetencao = db.OpenRecordset("SELECT * FROM Retencao ORDER BY Código", dbOpenDynaset)
    Call GetNewCode(Me, rstRetencao, 9999)
  End If

End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Código.Text = "" Then Exit Sub
  If Val(Código.Text) <= 0 Then Exit Sub
  
  rstRetencao.FindFirst "Código = " & Código.Text
  If Not rstRetencao.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If

End Sub

Sub ShowRecord()
  With rstRetencao
    Código.Text = Format(.Fields("Código"), String(Código.MaxLength, "0"))
    txtNome.Text = .Fields("Nome") & ""
    cboNomeFinanceira.Text = .Fields("NomeDaFinanceira").Value & ""
    txtValorRetencao.Text = Format(.Fields("ValorRetencao").Value, FORMAT_VALUE)
    
    If .Fields("Tipo").Value = "Cheque sem Juros" Then optChequeSemJuros.Value = True
    If .Fields("Tipo").Value = "Cheque com Juros" Then optChequeComJuros.Value = True
    If .Fields("Tipo").Value = "Carne sem Juros" Then optCarneSemJuros.Value = True
    If .Fields("Tipo").Value = "Carne com Juros" Then optCarneComJuros.Value = True
    
    Num_Registro = .Bookmark
  End With

End Sub

Private Function ValidaCampos() As Boolean
  
  If Not IsNumeric(Código.Text) Then
    ValidaCampos = True
    MsgBox "Código incorreto, verifique.", vbExclamation, "Atenção"
    Código.SetFocus
    Exit Function
  End If
  
  If Len(txtNome.Text) <= 0 Then
    ValidaCampos = True
    MsgBox "Nome incorreto, verifique.", vbExclamation, "Atenção"
    txtNome.SetFocus
    Exit Function
  End If
  
  If Len(cboNomeFinanceira.Text) <= 0 Then
    ValidaCampos = True
    MsgBox "Financeira incorreta, verifique.", vbExclamation, "Atenção"
    cboNomeFinanceira.SetFocus
    Exit Function
  End If
  
  If Not IsNumeric(txtValorRetencao.Text) Then
    ValidaCampos = True
    MsgBox "Retenção incorreta, verifique.", vbExclamation, "Atenção"
    txtValorRetencao.SetFocus
    Exit Function
  End If
  
End Function


