VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Begin VB.Form frmRadio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rádios"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRadio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   7005
   Begin VB.Frame fraX 
      Height          =   3735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtContatos 
         Height          =   285
         Left            =   1185
         MaxLength       =   40
         TabIndex        =   8
         Top             =   2880
         Width           =   4110
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2460
         Width           =   4110
      End
      Begin VB.TextBox txtInscricao 
         Height          =   285
         Left            =   4545
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2040
         Width           =   2010
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   285
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2040
         Width           =   2010
      End
      Begin VB.TextBox txtEstado 
         Height          =   285
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1620
         Width           =   435
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1185
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1620
         Width           =   3060
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   5370
      End
      Begin VB.TextBox Código 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1185
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "Pressione F5 para buscar o próximo código livre."
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   3870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Contatos"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2925
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2505
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição"
         Height          =   195
         Left            =   3495
         TabIndex        =   16
         Top             =   2085
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ "
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2085
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   5280
         TabIndex        =   14
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   405
         Width           =   495
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   6120
         Top             =   2400
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
         Bands           =   "frmRadio.frx":058A
      End
   End
End
Attribute VB_Name = "frmRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rstRadio     As Recordset
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
  Set rstRadio = db.OpenRecordset("SELECT * FROM Radio ORDER BY Código", dbOpenDynaset)
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
    
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstRadio
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
  
  With rstRadio
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
  
  With rstRadio
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
  
  With rstRadio
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
    MsgBox "Não existe nenhuma Rádio para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar esta Rádio ? "
  intResposta = MsgBox(strAuxStr, 20, "ATENÇÃO!!")
  If intResposta = 6 Then
    rstRadio.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim intErro As Integer
  
  On Error GoTo Processa_Erro
  
  If IsNull(Código.Text) Then intErro = True
  If Not intErro Then If Not IsNumeric(Código.Text) Then intErro = True
  If Not intErro Then If Val(Código.Text) <= 0 Then intErro = True
  If intErro = True Then
    DisplayMsg "Código incorreto, verifique."
    Código.SetFocus
    Exit Sub
  End If
  
  intErro = False
  If IsNull(txtNome.Text) Then intErro = True
  If Not intErro Then If txtNome.Text = "" Then intErro = True
  If intErro = True Then
    DisplayMsg "Nome da Rádio incorreto, verifique."
    txtNome.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstRadio
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Código") = CInt(Código.Text)
     Else
       .Edit
     End If
     .Fields("Nome") = txtNome.Text & ""
     .Fields("Endereco") = txtEndereco.Text & ""
     .Fields("Cidade") = txtCidade.Text & ""
     .Fields("Estado") = txtEstado.Text & ""
     .Fields("CNPJ") = txtCNPJ.Text & ""
     .Fields("Inscricao") = txtInscricao.Text & ""
     .Fields("Telefone") = txtTelefone.Text & ""
     .Fields("Contatos") = txtContatos.Text & ""

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
  txtEndereco.Text = ""
  txtCidade.Text = ""
  txtEstado.Text = ""
  txtCNPJ.Text = ""
  txtInscricao.Text = ""
  txtTelefone.Text = ""
  txtContatos.Text = ""
  
  Código.SetFocus
  Num_Registro = Null
  
  If Not rstRadio.EOF Then
    On Error Resume Next
    rstRadio.MoveFirst
    rstRadio.MovePrevious
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
  rstRadio.Close
  Set rstRadio = Nothing
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyF5 Then
    Set rstRadio = db.OpenRecordset("SELECT * FROM Radio ORDER BY Código", dbOpenDynaset)
    Call GetNewCode(Me, rstRadio, 9999)
  End If

End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Código.Text = "" Then Exit Sub
  If Val(Código.Text) <= 0 Then Exit Sub
  
  rstRadio.FindFirst "Código = " & Código.Text
  If Not rstRadio.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If

End Sub

Sub ShowRecord()
  With rstRadio
    Código.Text = Format(.Fields("Código"), String(Código.MaxLength, "0"))
    txtNome.Text = .Fields("Nome") & ""
    txtEndereco.Text = .Fields("Endereco") & ""
    txtCidade.Text = .Fields("Cidade") & ""
    txtEstado.Text = .Fields("Estado") & ""
    txtCNPJ.Text = .Fields("CNPJ") & ""
    txtInscricao.Text = .Fields("Inscricao") & ""
    txtTelefone.Text = .Fields("Telefone") & ""
    txtContatos.Text = .Fields("Contatos") & ""
    Num_Registro = .Bookmark
  End With

End Sub
