VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCartoes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cartões"
   ClientHeight    =   2910
   ClientLeft      =   1545
   ClientTop       =   2700
   ClientWidth     =   7665
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
   HelpContextID   =   1060
   Icon            =   "CCredito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   7665
   Begin VB.CheckBox chkTEF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Cartão exige TEF (Transferência Eletrônica de Fundos)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   315
      TabIndex        =   4
      Top             =   1680
      Width           =   4935
   End
   Begin VB.TextBox Dias 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6525
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1140
      Width           =   780
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1155
      MaxLength       =   25
      TabIndex        =   1
      Top             =   585
      Width           =   6150
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Pressione F5 para o Próximo Livre."
      Top             =   67
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Taxa 
      Height          =   360
      Left            =   1155
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
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
      Format          =   "##.###0"
      PromptChar      =   "_"
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   90
      Top             =   2070
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
      Bands           =   "CCredito.frx":4E95A
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Dias para receber"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4965
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Taxa (%)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   315
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   315
      TabIndex        =   6
      Top             =   645
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   315
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCartoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsCartoes As Recordset
Dim rsCR As Recordset

Private gsSql As String
Private gsWhere As String
Private gsOrder As String

Private Sub SearchRecord()

  If Not IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Apague todos os campos da tela com o botão NOVO."
    gsMsg = gsMsg & vbCrLf & "Selecione a Ordem de Pesquisa na lista e preencha com dados iniciais os campos respectivos."
    gsMsg = gsMsg & vbCrLf & "Pressione novamente este botão PROCURAR."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsSql = "SELECT * FROM Cartões "
  gsWhere = ""
  
  Select Case ActiveBar1.Tools("miOpOrdem").CBListIndex
    
    Case -1, 0  '"Por Código"
      If Len(Trim(Código.Text)) = 0 Then
        Código.Text = "0"
      End If
      gsWhere = "WHERE Código >= " & Código.Text
    Case 1  '"Por Nome"
      gsWhere = "WHERE Nome >= '" & Nome.Text & "'"
  End Select
  
  Set rsCartoes = db.OpenRecordset(gsSql & " " & gsWhere & " " & gsOrder, dbOpenDynaset)
  If Not rsCartoes.EOF Then
    Call ShowRecord
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em função dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  Dim Contador As Long
  
  If IsNull(Num_Registro) Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "Não existe registro para apagar !"
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  
  Rem Procura a adminsitradora nos cartões
  Call StatusMsg("Aguarde, verificando se este cartão não está em uso.")
  Contador = 0
  Set rsCR = db.OpenRecordset("Contas a Receber", , dbReadOnly)
  
  rsCR.Index = "Tipo"
  rsCR.Seek ">", "O", Contador
  If rsCR.NoMatch Then GoTo Apaga
  Contador = rsCR("Contador")
  If rsCR("Administradora") = rsCartoes("Código") Then
    Call StatusMsg("")
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "Este cartão não pode ser apagado." & vbCrLf & "Existem lançamentos no cadastro de cartões de crédito que usam este cartão."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
Apaga:
  Call StatusMsg("")
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar " & Nome.Text & "?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    rsCartoes.Delete
    '09/06/2005 - Daniel
    'Adicionado rotina para limpeza da tela evitando
    'a sujeira nos objetos do cadastro
    Call ClearScreen
    '
    'Call MovePrevious
  End If
  
End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  Rem Verifica código
  If IsNull(Código.Text) Then Erro = True
  If Not Erro Then If Not IsNumeric(Código.Text) Then Erro = True
  If Not Erro Then If Val(Código.Text) < 1 Then Erro = True
  If Not Erro Then If Val(Código.Text) > 9999 Then Erro = True
  
  If Erro Then
    gsTitle = LoadResString(201)
    gsMsg = "Código deve ter valor entre 1 e 9999."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Código.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Nome.Text) Then Erro = True
  If Erro = False Then If Nome.Text = "" Then Erro = True
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Digite o nome da administradora."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Nome.SetFocus
    Exit Sub
  End If
  
  If IsNull(Dias.Text) Or Not IsNumeric(Dias.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Digite o número de dias."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Dias.SetFocus
    Exit Sub
  End If
  If Val(Dias.Text) < 0 Then Dias.Text = 0
  
  If IsNull(Taxa.Text) Or Not IsNumeric(Taxa.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Digite a taxa."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Taxa.SetFocus
    Exit Sub
  End If
  
  '28/01/2005 - Daniel
  'Tratamento para o campo Cartões.Código o tipo de dados
  'deste campo é inteiro mas estava dando problemas na tela
  'Lançamento de cartões de crédito quando este campo se torna
  '[Contas a Receber].Administradora e o tipo de dados do campo
  'Administradora é byte, então neste exato ponto viemos tratá-lo
  'para não cadastrarem valores superiores a byte.
  '
  '09/06/2005 - Daniel
  'Correção: Tratado o valor 250 para considerar como número e não como string
  If (Código.Text) > 250 Then
    MsgBox "Escolha um Código entre 1 a 250 que não tenha sido cadastrado.", vbExclamation, "Atenção"
    Código.SetFocus
    Exit Sub
  End If
  '--------------------------------------------------------------
  
  Call StatusMsg("Gravando ...")
  
  With rsCartoes
   If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código") = Código.Text
   Else
      .Edit
   End If
   .Fields("Nome") = Nome.Text
   .Fields("Taxa") = CDbl(Taxa.Text)
   .Fields("Dias Pagar") = Dias.Text
   .Fields("TEF") = IIf((chkTEF.Value = vbChecked), True, False)
   .Update
   Num_Registro = .LastModified
   .Bookmark = Num_Registro
  End With
  
  Call StatusMsg("")
  
End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  Código.Text = ""
  Nome.Text = ""
  Taxa.Text = ""
  Dias.Text = ""
  chkTEF.Value = vbUnchecked
  Num_Registro = Null
  If Not rsCartoes.EOF And Not rsCartoes.BOF Then
    rsCartoes.MoveFirst
    If Not rsCartoes.BOF Then
      rsCartoes.MovePrevious
    End If
  End If
  Código.SetFocus
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsCartoes
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsCartoes
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsCartoes
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsCartoes
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
  On Error GoTo 0
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
      Call SearchRecord
  End Select
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  gsOrder = ""
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0 '"Por Código"
          gsOrder = "ORDER BY Código"
        Case 1 '"Por Nome"
          gsOrder = "ORDER BY Nome"
      End Select
  End Select
End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    gsOrder = "ORDER BY Código"
    Set rsCartoes = db.OpenRecordset("SELECT * FROM Cartões " & gsOrder, dbOpenDynaset)
    Call GetNewCode(Me, rsCartoes, 9999)
  End If
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Código.Text = "" Then Exit Sub
  If Val(Código.Text) <= 0 Then Exit Sub
  
  rsCartoes.FindFirst "Código = " & Código.Text
  If Not rsCartoes.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If
End Sub

Private Sub Dias_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
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

  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Código"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Nome"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  ActiveBar1.RecalcLayout
  ActiveBar1.Refresh
  
  DoEvents
  Me.Show
  
  gsOrder = "ORDER BY Código"
  Set rsCartoes = db.OpenRecordset("SELECT * FROM Cartões " & gsOrder, dbOpenDynaset)
  
  Call ActiveBarLoadToolTips(Me)
  
  Call ClearScreen
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsCartoes.Close
  Set rsCartoes = Nothing
End Sub

Private Sub ShowRecord()
  Código.Text = Format(rsCartoes("Código"), String(Código.MaxLength, "0"))
  Nome.Text = rsCartoes("Nome")
  Taxa.Text = Format(rsCartoes("Taxa"), "##0.######")
  Dias.Text = rsCartoes("Dias Pagar")
  chkTEF.Value = IIf(rsCartoes("TEF"), vbChecked, vbUnchecked)
  Num_Registro = rsCartoes.Bookmark
End Sub

Private Sub Taxa_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub
