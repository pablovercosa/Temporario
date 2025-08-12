VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Begin VB.Form frmSupervisores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Supervisores"
   ClientHeight    =   5775
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
   Icon            =   "frmSupervisores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7005
   Begin VB.Frame fraFunc 
      Caption         =   "Visualize os Funcionários deste Supervisor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   6975
      Begin VB.ListBox lstFuncionarios 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   2400
         ItemData        =   "frmSupervisores.frx":058A
         Left            =   120
         List            =   "frmSupervisores.frx":058C
         TabIndex        =   3
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame fraX 
      Caption         =   "Dados do Supervisor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtObs 
         Height          =   855
         Left            =   1320
         MaxLength       =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Digite no máximo 1200 caracteres em Observações do Supervisor."
         Top             =   1200
         Width           =   5370
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   3495
      End
      Begin VB.TextBox Código 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "Pressione F5 para buscar o próximo código livre."
         Top             =   360
         Width           =   615
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   6240
         Top             =   2160
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
         Bands           =   "frmSupervisores.frx":058E
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   690
         TabIndex        =   7
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Observações"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmSupervisores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro    As Variant
Dim rstSupervisores As Recordset

Private Sub Form_Load()
  Call CenterForm(Me)

  Set rstSupervisores = db.OpenRecordset("Supervisores", dbOpenDynaset)
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'  If gbSkipKey = True Then
'    KeyAscii = 0
'    gbSkipKey = False
'  End If
'  If KeyAscii = 13 Then
'     SendKeys "{Tab}"
'     KeyAscii = 0
' End If
'End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstSupervisores
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
  
  With rstSupervisores
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
  
  With rstSupervisores
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
  
  With rstSupervisores
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
    MsgBox "Não existe nenhum Supervisor para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If
  
  'Verificamos se há vendedor relacionado com
  'o supervisor
  If VerificarIntegridade Then
    MsgBox "Não exclua este Supervisor, ele tem Funcionários ligados a ele, verifique.", vbExclamation, "Integridade de Dados"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar este Supervisor ? "
  intResposta = MsgBox(strAuxStr, 20, "ATENÇÃO!!")
  If intResposta = 6 Then
    rstSupervisores.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim intErro As Integer
  
  On Error GoTo Processa_Erro
  
  If Not IsNumeric(Código.Text) Then
    DisplayMsg "Código incorreto, verifique."
    Código.SetFocus
    Exit Sub
  End If
  
  If Len(txtNome.Text) <= 0 Then
    DisplayMsg "Nome do Supervisor inválido, verifique."
    txtNome.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstSupervisores
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Código") = CInt(Código.Text)
     Else
       .Edit
     End If
     .Fields("Nome") = txtNome.Text & ""
     .Fields("Obs").Value = txtObs.Text & ""

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
  txtObs.Text = ""
  lstFuncionarios.Clear
  
  Código.SetFocus
  Num_Registro = Null
  
  If Not rstSupervisores.EOF Then
    On Error Resume Next
    rstSupervisores.MoveFirst
    rstSupervisores.MovePrevious
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
  rstSupervisores.Close
  Set rstSupervisores = Nothing
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    Set rstSupervisores = db.OpenRecordset("SELECT * FROM Supervisores ORDER BY Código", dbOpenDynaset)
    Call GetNewCode(Me, rstSupervisores, 9999)
  End If
End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Código.Text = "" Then Exit Sub
  If Val(Código.Text) <= 0 Then Exit Sub
  
  rstSupervisores.FindFirst "Código = " & Código.Text
  If Not rstSupervisores.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If

End Sub

Sub ShowRecord()
  With rstSupervisores
    Código.Text = Format(.Fields("Código").Value, String(Código.MaxLength, "0"))
    txtNome.Text = .Fields("Nome").Value & ""
    txtObs.Text = .Fields("Obs").Value & ""
    
    PreencherLista
    
    Num_Registro = .Bookmark
  End With
End Sub

Private Function VerificarIntegridade() As Boolean
  'Integridade no Relacionamento
  'Esta função tem finalidade de barrar a exclusão caso
  'exista Funcionário amarrado com o Supervisor
  Dim rstFuncionarios As Recordset
  Dim strQuery        As String
  
  strQuery = "SELECT Supervisor "
  strQuery = strQuery & " FROM Funcionários "
  strQuery = strQuery & " WHERE Supervisor = " & CInt(Código.Text)
  
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        If .Fields("Supervisor").Value = CInt(Código.Text) Then VerificarIntegridade = True
      
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstFuncionarios = Nothing
  
End Function

Private Sub PreencherLista()
  Dim rstFuncionarios As Recordset
  Dim strQuery        As String
  
  lstFuncionarios.Clear
  
  strQuery = "SELECT Código, Nome, Supervisor "
  strQuery = strQuery & " FROM Funcionários "
  strQuery = strQuery & " WHERE Supervisor = " & CInt(Código.Text)
  
  Set rstFuncionarios = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        lstFuncionarios.AddItem (Right(("00000" & CStr(.Fields("Código").Value)), 5) & " - " & .Fields("Nome").Value & "")
        
      .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstFuncionarios = Nothing

End Sub

