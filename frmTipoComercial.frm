VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Begin VB.Form frmTipoComercial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Comercial"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipoComercial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   5310
   Begin VB.Frame fraX 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtNome 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   3870
      End
      Begin VB.TextBox C�digo 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1185
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "Pressione F5 para buscar o pr�ximo C�digo livre."
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   405
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   120
         Top             =   1320
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
         Bands           =   "frmTipoComercial.frx":058A
      End
   End
End
Attribute VB_Name = "frmTipoComercial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Num_Registro      As Variant
Dim rstTipoComercial  As Recordset
'Nota: O Text Box do C�digo teve que ser chamado de C�digo ao inv�s de txtCodigo
'para podermos aproveitar a rotina em QSGeral de busca autom�tica do pr�ximo c�digo livre

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
  Set rstTipoComercial = db.OpenRecordset("SELECT * FROM TipoComercial ORDER BY C�digo", dbOpenDynaset)
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
  
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstTipoComercial
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
  
  With rstTipoComercial
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
  
  With rstTipoComercial
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
  
  With rstTipoComercial
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
    MsgBox "N�o existe nenhum Tipo Comercial para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar este Tipo Comercial ? "
  intResposta = MsgBox(strAuxStr, 20, "ATEN��O!!")
  If intResposta = 6 Then
    rstTipoComercial.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim intErro As Integer
  
  On Error GoTo Processa_Erro
  
  If IsNull(C�digo.Text) Then intErro = True
  If Not intErro Then If Not IsNumeric(C�digo.Text) Then intErro = True
  If Not intErro Then If Val(C�digo.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "C�digo incorreto, verifique."
    C�digo.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If IsNull(txtNome.Text) Then intErro = True
  If Not intErro Then If txtNome.Text = "" Then intErro = True
  If intErro = True Then
    DisplayMsg "Nome do Tipo Comercial incorreto, verifique."
    txtNome.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstTipoComercial
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("C�digo") = CInt(C�digo.Text)
     Else
       .Edit
     End If
     .Fields("Descricao") = txtNome.Text & ""

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
  
  C�digo.Text = ""
  txtNome.Text = ""
  
  C�digo.SetFocus
  Num_Registro = Null
  
  If Not rstTipoComercial.EOF Then
    On Error Resume Next
    rstTipoComercial.MoveFirst
    rstTipoComercial.MovePrevious
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
  rstTipoComercial.Close
  Set rstTipoComercial = Nothing
End Sub

Private Sub C�digo_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyF5 Then
    Set rstTipoComercial = db.OpenRecordset("SELECT * FROM TipoComercial ORDER BY C�digo", dbOpenDynaset)
    Call GetNewCode(Me, rstTipoComercial, 9999)
  End If

End Sub

Private Sub C�digo_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub C�digo_LostFocus()
  If IsNull(C�digo.Text) Then Exit Sub
  If C�digo.Text = "" Then Exit Sub
  If Val(C�digo.Text) <= 0 Then Exit Sub
  
  rstTipoComercial.FindFirst "C�digo = " & C�digo.Text
  If Not rstTipoComercial.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If

End Sub

Sub ShowRecord()
  With rstTipoComercial
    C�digo.Text = Format(.Fields("C�digo"), String(C�digo.MaxLength, "0"))
    txtNome.Text = .Fields("Descricao") & ""
    Num_Registro = .Bookmark
  End With

End Sub

