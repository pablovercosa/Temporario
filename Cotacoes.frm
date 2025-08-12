VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCotacoes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cotações"
   ClientHeight    =   2505
   ClientLeft      =   1455
   ClientTop       =   2445
   ClientWidth     =   6990
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
   HelpContextID   =   1130
   Icon            =   "Cotacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2505
   ScaleWidth      =   6990
   Begin MSMask.MaskEdBox Dia 
      Height          =   345
      Left            =   915
      TabIndex        =   0
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   90
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem"
      Height          =   615
      Left            =   135
      TabIndex        =   7
      Top             =   1350
      Width           =   6735
      Begin VB.OptionButton O_Moeda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Moeda + Data"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   750
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton O_Data 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data + Moeda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   1455
      End
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3930
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Moeda"
      Top             =   2250
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSMask.MaskEdBox Cotação 
      Height          =   345
      Left            =   930
      TabIndex        =   2
      Top             =   945
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,##0.0000"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Moeda 
      Bindings        =   "Cotacoes.frx":4E95A
      DataSource      =   "Data1"
      Height          =   345
      Left            =   930
      TabIndex        =   1
      Top             =   510
      Width           =   1185
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
      BackColorOdd    =   16777152
      Columns(0).Width=   3200
      _ExtentX        =   2090
      _ExtentY        =   609
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
      Left            =   3870
      Top             =   -30
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
      Bands           =   "Cotacoes.frx":4E96E
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Cotação"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1005
      Width           =   600
   End
   Begin VB.Label Nome_Moeda 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2175
      TabIndex        =   5
      Top             =   510
      Width           =   4695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Moeda"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   570
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   150
      Width           =   345
   End
End
Attribute VB_Name = "frmCotacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Num_Registro As Variant
Dim rsCotacoes As Recordset
Dim rsMoedas As Recordset

Private Sub DeleteRecord()
  Dim Resposta As Integer
  
  If IsNull(Num_Registro) Then
    Beep
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If
  
  Resposta = MsgBox(("Deseja realmente apagar esta cotação ?"), 20, "ATENÇÃO!!")
  If Resposta = 6 Then
    rsCotacoes.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  Rem Verifica Dia
  If IsNull(Dia.Text) Then Erro = True
  If Not Erro Then If Not IsDate(Dia.Text) Then Erro = True
  If Erro Then
    DisplayMsg "Data invalida."
    Dia.SetFocus
    Exit Sub
  End If
  
  Rem Verifica moeda
  If IsNull(Combo_Moeda.Text) Then Erro = True
  If Not Erro Then If Not IsNumeric(Combo_Moeda.Text) Then Erro = True
  If Not Erro Then If Combo_Moeda.Text = 0 Then Erro = True
  If Erro Then
    DisplayMsg "Moeda inválida ."
    Combo_Moeda.SetFocus
    Exit Sub
  End If
  
  If Val(Combo_Moeda.Text) = 1 Then
    DisplayMsg "Não se pode cadastrar cotação para a moeda 1."
    Combo_Moeda.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Cotação.Text) Then Erro = True
  If Not Erro Then If Cotação.Text = "" Then Erro = True
  If Not Erro Then If Not IsNumeric(Cotação.Text) Then Erro = True
  If Not Erro Then If CDbl(Cotação.Text) <= 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Cotação deve ser um número maior do que 0."
    Cotação.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  
  With rsCotacoes
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Data") = Dia.Text
    Else
      .Edit
    End If
    .Fields("Moeda") = Combo_Moeda.Text
    .Fields("Cotação") = CDbl(Cotação.Text)
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  End With
  
  Call StatusMsg("")
  
End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  Dia.Mask = ""
  Dia.Text = ""
  Dia.Mask = "##/##/####"
  Combo_Moeda.Text = ""
  Nome_Moeda.Caption = ""
  Cotação.Text = ""
  Num_Registro = Null
  Dia.SetFocus
End Sub

Private Sub Combo_Moeda_CloseUp()
  Combo_Moeda.Text = Combo_Moeda.Columns(1).Text
  Combo_Moeda_LostFocus
End Sub

Private Sub Combo_Moeda_LostFocus()

  rsMoedas.Index = "Código"
  Nome_Moeda.Caption = ""
  
  If IsNull(Combo_Moeda.Text) Then Exit Sub
  If Combo_Moeda.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Moeda.Text) Then Exit Sub
  If Val(Combo_Moeda.Text) > 99 Then Exit Sub
  If Val(Combo_Moeda.Text) < 1 Then Exit Sub
  
  rsMoedas.Seek "=", Combo_Moeda.Text
  If Not rsMoedas.NoMatch Then
     Nome_Moeda.Caption = rsMoedas("Nome")
  Else
     Combo_Moeda.Text = 0
  End If
  
  Dia_LostFocus

End Sub

Private Sub Cotação_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Dia_GotFocus()
 If IsNull(Dia.Text) Or Dia.Text = "" Then
    Dia.Text = Format$(Date, "dd/mm/yyyy")
 End If
End Sub

Private Sub Dia_LostFocus()
  Dim Dia_Var As Variant
  
  If IsNull(Dia.Text) Then Exit Sub
  If Not IsDate(Dia.Text) Then Exit Sub
  Dia.Text = Format(Dia.Text, "dd/mm/yyyy")
  If IsNull(Combo_Moeda.Text) Then Exit Sub
  If Not IsNumeric(Combo_Moeda.Text) Then Exit Sub
  Dia_Var = CDate(Dia.Text)

  Dia.Text = Ajusta_Data(Dia.Text)
  
  With rsCotacoes
    .FindFirst "Moeda = " & Combo_Moeda.Text & " AND Data = #" & Format(Dia_Var, "mm/dd/yyyy") & "#"
    If Not .NoMatch Then
      Num_Registro = rsCotacoes.Bookmark
      ShowRecord
    Else
      If Not IsNull(Num_Registro) Then
        .Bookmark = Num_Registro
      End If
    End If
  End With
  
End Sub

Private Sub Dia_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Dia.Text = frmCalendario.gsDateCalender(Dia.Text)
  End Select
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsCotacoes
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
  With rsCotacoes
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
  With rsCotacoes
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
  With rsCotacoes
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsCotacoes = db.OpenRecordset("SELECT * FROM Cotações ORDER BY Moeda, Data", dbOpenDynaset)
  Set rsMoedas = db.OpenRecordset("Moedas", , dbReadOnly)
  Num_Registro = Null
  Data1.DatabaseName = gsQuickDBFileName
  Call ActiveBarLoadToolTips(Me)
  '07/06/2005 - Daniel
  'Para facilitação da digitação, colocamos
  'à data atual no objeto "Dia"
  Dia.Text = Format(Data_Atual, "DD/MM/YYYY")
End Sub

Private Sub ShowRecord()
  Call StatusMsg("")
  Dia.Text = Format$(rsCotacoes("Data"), "dd/mm/yyyy")
  Combo_Moeda.Text = rsCotacoes("Moeda")
  Cotação.Text = rsCotacoes("Cotação")
  rsMoedas.Index = "Código"
  Nome_Moeda.Caption = ""
  rsMoedas.Seek "=", Combo_Moeda.Text
  If Not rsMoedas.NoMatch Then
    Nome_Moeda.Caption = rsMoedas("Nome")
  End If
  Num_Registro = rsCotacoes.Bookmark
End Sub

Private Sub O_Data_Click()
  Set rsCotacoes = db.OpenRecordset("SELECT * FROM Cotações ORDER BY Data, Moeda", dbOpenDynaset)
End Sub

Private Sub O_Moeda_Click()
  Set rsCotacoes = db.OpenRecordset("SELECT * FROM Cotações ORDER BY Moeda, Data", dbOpenDynaset)
End Sub
