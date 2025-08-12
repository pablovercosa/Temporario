VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Begin VB.Form frmTransportadoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Transportadoras"
   ClientHeight    =   3720
   ClientLeft      =   3090
   ClientTop       =   2130
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
   Icon            =   "Transportadoras.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   7005
   Begin VB.TextBox Contatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   8
      Top             =   2745
      Width           =   3870
   End
   Begin VB.TextBox Telefone 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   7
      Top             =   2280
      Width           =   3870
   End
   Begin VB.TextBox Inscrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4620
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1860
      Width           =   2010
   End
   Begin VB.TextBox CGC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1860
      Width           =   2010
   End
   Begin VB.TextBox Estado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6195
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox Cidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1440
      Width           =   3870
   End
   Begin VB.TextBox Endereço 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1020
      Width           =   5370
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Pressione F5  para Próximo Código Livre"
      Top             =   180
      Width           =   1155
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem"
      ForeColor       =   &H00FF0000&
      Height          =   840
      Left            =   5265
      TabIndex        =   9
      Top             =   60
      Width           =   1350
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   255
         TabIndex        =   11
         Top             =   285
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   255
         TabIndex        =   10
         Top             =   555
         Width           =   870
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   180
      Top             =   3420
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
      Bands           =   "Transportadoras.frx":4E95A
   End
   Begin VB.Label Label10 
      Caption         =   "colocar PLACA"
      Height          =   225
      Left            =   2835
      TabIndex        =   21
      Top             =   180
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label Label9 
      Caption         =   "Contatos"
      Height          =   285
      Left            =   315
      TabIndex        =   20
      Top             =   2790
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Telefone"
      Height          =   285
      Left            =   315
      TabIndex        =   19
      Top             =   2325
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Inscrição"
      Height          =   285
      Left            =   3570
      TabIndex        =   18
      Top             =   1905
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "CNPJ/CPF"
      Height          =   285
      Left            =   315
      TabIndex        =   17
      Top             =   1905
      Width           =   810
   End
   Begin VB.Label Label5 
      Caption         =   "Estado"
      Height          =   285
      Left            =   5355
      TabIndex        =   16
      Top             =   1485
      Width           =   645
   End
   Begin VB.Label Label4 
      Caption         =   "Cidade"
      Height          =   285
      Left            =   315
      TabIndex        =   15
      Top             =   1485
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "Endereço"
      Height          =   285
      Left            =   315
      TabIndex        =   14
      Top             =   1065
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nome"
      Height          =   285
      Left            =   315
      TabIndex        =   13
      Top             =   645
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   285
      Left            =   315
      TabIndex        =   12
      Top             =   225
      Width           =   645
   End
End
Attribute VB_Name = "frmTransportadoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsTransportadoras As Recordset

Private Sub MoveFirst()
  On Error Resume Next
  With rsTransportadoras
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
  With rsTransportadoras
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
  With rsTransportadoras
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
  With rsTransportadoras
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
  Dim Resposta As Integer
  Dim Aux_Código As Double
  Dim Aux_Str As String

  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If

  Aux_Str = "Deseja realmente apagar esta transportadora ? "
  Resposta = MsgBox(Aux_Str, 20, "ATENÇÃO!!")
  If Resposta = 6 Then
    rsTransportadoras.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  On Error GoTo Processa_Erro
  
  If IsNull(Código.Text) Then Erro = True
  If Not Erro Then If Not IsNumeric(Código.Text) Then Erro = True
  If Not Erro Then If Val(Código.Text) <= 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Código incorreto, verifique."
    Código.SetFocus
    Exit Sub
  End If
  Código.Text = RTrim(LTrim(Código.Text))
  
  Erro = False
  If IsNull(Nome.Text) Then Erro = True
  If Not Erro Then If Nome.Text = "" Then Erro = True
  If Erro = True Then
    DisplayMsg "Nome da transportadora incorreto, verifique."
    Nome.SetFocus
    Exit Sub
  End If
  
  Nome.Text = RTrim(LTrim(Nome.Text))
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rsTransportadoras
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Código") = Val(Código.Text)
     Else
       .Edit
     End If
     .Fields("Nome") = Nome.Text & ""
     .Fields("Endereço") = RTrim(LTrim(Endereço.Text)) & ""
     .Fields("Cidade") = RTrim(LTrim(Cidade.Text)) & ""
     .Fields("Estado") = RTrim(LTrim(Estado.Text)) & ""
     .Fields("CGC") = RTrim(LTrim(CGC.Text)) & ""
     .Fields("Inscrição") = RTrim(LTrim(Inscrição.Text)) & ""
     .Fields("Telefone") = RTrim(LTrim(Telefone.Text)) & ""
     .Fields("Contatos") = RTrim(LTrim(Contatos.Text)) & ""
     .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
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
  Nome.Text = ""
  Endereço.Text = ""
  Cidade.Text = ""
  Estado.Text = ""
  CGC.Text = ""
  Inscrição.Text = ""
  Telefone.Text = ""
  Contatos.Text = ""
  Código.SetFocus
  Num_Registro = Null
  
  If Not rsTransportadoras.EOF Then
    On Error Resume Next
    rsTransportadoras.MoveFirst
    rsTransportadoras.MovePrevious
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

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    Call O_Código_Click
    Call GetNewCode(Me, rsTransportadoras, 9999)
  End If
End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Código.Text = "" Then Exit Sub
  If Val(Código.Text) <= 0 Then Exit Sub
  
  rsTransportadoras.FindFirst "Código = " & Código.Text
  If Not rsTransportadoras.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If
End Sub

Sub ShowRecord()
  With rsTransportadoras
    Código.Text = Format(.Fields("Código"), String(Código.MaxLength, "0"))
    Nome.Text = .Fields("Nome") & ""
    Endereço.Text = .Fields("Endereço") & ""
    Cidade.Text = .Fields("Cidade") & ""
    Estado.Text = .Fields("Estado") & ""
    CGC.Text = .Fields("CGC") & ""
    Inscrição.Text = .Fields("Inscrição") & ""
    Telefone.Text = .Fields("Telefone") & ""
    Contatos.Text = .Fields("Contatos") & ""
    Num_Registro = .Bookmark
  End With
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

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsTransportadoras = db.OpenRecordset("SELECT * FROM Transportadoras ORDER BY Código", dbOpenDynaset)
  Me.Show
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsTransportadoras.Close
  Set rsTransportadoras = Nothing
End Sub

Private Sub O_Código_Click()
  Set rsTransportadoras = db.OpenRecordset("SELECT * FROM Transportadoras ORDER BY Código", dbOpenDynaset)
End Sub

Private Sub O_Nome_Click()
  Set rsTransportadoras = db.OpenRecordset("SELECT * FROM Transportadoras ORDER BY Nome", dbOpenDynaset)
End Sub
