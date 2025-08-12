VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmContas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Manutenção de Contas Correntes"
   ClientHeight    =   4080
   ClientLeft      =   1740
   ClientTop       =   1950
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1110
   Icon            =   "Contas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   6645
   Begin VB.TextBox Contabilidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2565
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2835
      Width           =   3885
   End
   Begin VB.TextBox Agência 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1050
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1125
      Width           =   1665
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   0
      Top             =   300
      Width           =   1200
   End
   Begin VB.TextBox Telefone 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1050
      MaxLength       =   30
      TabIndex        =   6
      Top             =   2415
      Width           =   5400
   End
   Begin VB.TextBox Gerente 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1050
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1995
      Width           =   5400
   End
   Begin VB.TextBox Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1050
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1575
      Width           =   5400
   End
   Begin VB.TextBox Conta 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4155
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1125
      Width           =   2295
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5895
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Banco"
      Top             =   3870
      Visible         =   0   'False
      Width           =   1470
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
      Bindings        =   "Contas.frx":4E95A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1050
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      DataFieldList   =   "Nome"
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
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   3690
      Top             =   90
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
      Bands           =   "Contas.frx":4E96E
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código para Contabilidade"
      Height          =   255
      Left            =   285
      TabIndex        =   16
      Top             =   2865
      Width           =   2205
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefones"
      Height          =   255
      Left            =   285
      TabIndex        =   15
      Top             =   2460
      Width           =   750
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Gerente"
      Height          =   255
      Left            =   285
      TabIndex        =   14
      Top             =   2070
      Width           =   660
   End
   Begin VB.Label Nome_Banco 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   735
      Width           =   4170
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Agência"
      Height          =   255
      Left            =   285
      TabIndex        =   12
      Top             =   1155
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   255
      Left            =   285
      TabIndex        =   11
      Top             =   330
      Width           =   615
   End
   Begin VB.Label label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   285
      TabIndex        =   9
      Top             =   1620
      Width           =   720
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   285
      TabIndex        =   10
      Top             =   750
      Width           =   630
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Conta"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   1155
      Width           =   540
   End
End
Attribute VB_Name = "frmContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsContas As Recordset
Dim rsBancos As Recordset
Dim rsLançamentos As Recordset

Private Sub DeleteRecord()
  Dim Resposta As Integer

  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If

  Rem Procura o banco nos cheques
  Call StatusMsg("Aguarde, verificando se esta conta não está em uso.")
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários", , dbReadOnly)
  rsLançamentos.Index = "Conta"
  rsLançamentos.Seek ">", rsContas("Código"), CDate("01/01/1980"), 0
  If rsLançamentos.NoMatch Then GoTo Apaga
  If rsLançamentos("Conta") = rsContas("Código") Then
    Call StatusMsg("")
    Beep
    MsgBox ("Esta conta não pode ser apagada." & vbCrLf & "Existem lançamentos bancários que usam esta conta.")
    Exit Sub
  End If

Apaga:
  Resposta = MsgBox(("Deseja realmente apagar esta conta ?"), 20, "ATENÇÃO!!")
  If Resposta = 6 Then
    rsContas.Delete
    Num_Registro = Null
    ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  Rem Verifica Conta
  Erro = False
  If IsNull(Código.Text) Then Erro = True
  If Erro = False Then If Código.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Código.Text) Then Erro = True
  '28/11/2006 - Anderson
  'Alteração do número de contas bancárias de 99 para 255
  'Solicitado por: 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
  If Erro = False Then If Val(Código.Text) <= 0 Or Val(Código.Text) > 255 Then Erro = True
  
  If Erro = True Then
    DisplayMsg "Use códigos de 1 a 255."
    Código.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Agência.Text) Then Erro = True
  If Erro = False Then If Agência.Text = "" Then Erro = True
  If Erro = True Then
    DisplayMsg "Por favor, digite a agência."
    Agência.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Conta.Text) Then Erro = True
  If Erro = False Then If Conta.Text = "" Then Erro = True
  If Erro = True Then
    DisplayMsg "Por favor, digite a conta."
    Conta.SetFocus
    Exit Sub
  End If
  
  
  If Nome_Banco.Caption = "" Then
    DisplayMsg "Banco inválido, verifique."
    Combo_Banco.SetFocus
    Exit Sub
  End If
  
  
  Erro = False
  If IsNull(Descrição.Text) Then Erro = True
  If Erro = False Then If Descrição.Text = "" Then Erro = True
  If Erro = True Then
    DisplayMsg "Por favor, digite a descrição da conta."
    Descrição.SetFocus
    Exit Sub
  End If
  
  If IsNull(Contabilidade.Text) Then Contabilidade.Text = 0
  If Not IsNumeric(Contabilidade.Text) Then Contabilidade.Text = 0
  If Val(Contabilidade.Text) < 0 Then Contabilidade.Text = 0
  
  Call StatusMsg("Gravando ...")
  
  With rsContas
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código") = Val(Código.Text)
    Else
      .Edit
    End If
    .Fields("Agência") = Agência.Text
    .Fields("Conta") = Conta.Text
    .Fields("Banco") = Val(Combo_Banco.Text)
    .Fields("Descrição") = Descrição.Text
    .Fields("Gerente") = Gerente.Text
    .Fields("Telefone") = Telefone.Text
    .Fields("Contabilidade") = Contabilidade.Text
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  End With
  
  Call StatusMsg("")
  
End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  Código.Text = ""
  Agência.Text = ""
  Conta.Text = ""
  Combo_Banco.Text = ""
  Nome_Banco.Caption = ""
  Descrição.Text = ""
  Gerente.Text = ""
  Telefone.Text = ""
  Contabilidade.Text = ""
  
  If Not rsContas.EOF Then
    rsContas.MoveFirst
    rsContas.MovePrevious
  End If
  
  Num_Registro = Null
  Código.SetFocus
End Sub

Private Sub Código_LostFocus()

  If IsNull(Código.Text) Then Exit Sub
  If Not IsNumeric(Código.Text) Then Exit Sub
  '28/11/2006 - Anderson
  'Alteração do número de contas correntes
  'Solicitado por 2227883 - SANTA FÉ DO ARAGUAIA PREFEITURA MUNICIPAL
  If Val(Código.Text) > 255 Then
    MsgBox "Digite um código válido de 1 a 255.", vbInformation, "Atenção!"
    Código.SetFocus
    Exit Sub
  End If
  If Val(Código.Text) < 1 Then Exit Sub
  If Val(Código.Text) > 255 Then Exit Sub
  
  With rsContas
    .FindFirst "Código = " & CInt(Código.Text)
    If Not .NoMatch Then
      ShowRecord
    End If
  End With

End Sub

Private Sub Combo_Banco_CloseUp()
  Combo_Banco.Text = Combo_Banco.Columns(1).Text
  Combo_Banco_LostFocus
End Sub

Private Sub Combo_Banco_LostFocus()
  Nome_Banco.Caption = ""
  If IsNull(Combo_Banco.Text) Then Exit Sub
  If Not IsNumeric(Combo_Banco.Text) Then Exit Sub
  If Val(Combo_Banco.Text) < 0 Or Val(Combo_Banco.Text) > 9999 Then Exit Sub

  rsBancos.Index = "Código"
  rsBancos.Seek "=", Val(Combo_Banco.Text)
  If rsBancos.NoMatch Then Exit Sub
  Nome_Banco.Caption = rsBancos("Nome") & ""

End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsContas
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
  With rsContas
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
  With rsContas
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
  With rsContas
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
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
    Case "miOpSearch"
      'Call SearchRecord
  End Select
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    '11/05/2007 - Anderson
    'Correção para geração do código automático
    'Call GetNewCode(Me, rsContas, 99)
    Call GetNewCode(Me, rsContas, 999)
  End If
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
  Set rsContas = db.OpenRecordset("SELECT * FROM [Contas Bancárias] ORDER BY Código", dbOpenDynaset)
  Set rsBancos = db.OpenRecordset("Bancos")
  Data1.DatabaseName = gsQuickDBFileName
  If Not rsContas.EOF Then
    rsContas.MovePrevious
  Else
    Num_Registro = Null
  End If
  Call ActiveBarLoadToolTips(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsContas.Close
  Set rsContas = Nothing
  rsBancos.Close
  Set rsBancos = Nothing
End Sub

Private Sub ShowRecord()
  Código.Text = Format(rsContas("Código"), String(Código.MaxLength, "0"))
  Conta.Text = rsContas("Conta") & ""
  Agência.Text = rsContas("Agência") & ""
  Combo_Banco.Text = rsContas("Banco") & ""
  Combo_Banco_LostFocus
  Descrição.Text = rsContas("Descrição") & ""
  Gerente.Text = rsContas("Gerente") & ""
  Telefone.Text = rsContas("Telefone") & ""
  Contabilidade.Text = rsContas("Contabilidade")
  Num_Registro = rsContas.Bookmark
End Sub


