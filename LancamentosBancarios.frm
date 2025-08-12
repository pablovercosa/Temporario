VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLancaContas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Contas Correntes - Lan�amentos"
   ClientHeight    =   4770
   ClientLeft      =   2760
   ClientTop       =   1635
   ClientWidth     =   7680
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
   HelpContextID   =   1340
   Icon            =   "LancamentosBancarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdCloseSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Fim Pesquisa"
      Height          =   400
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox Documento 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Data Data1 
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
      Height          =   345
      Left            =   5805
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   3780
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem"
      Height          =   615
      Left            =   5010
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   2430
      Begin VB.OptionButton O_Banco 
         Caption         =   "Data + Conta"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox Descri��o 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1155
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1950
      Width           =   6360
   End
   Begin MSMask.MaskEdBox D�bito 
      Height          =   360
      Left            =   3435
      TabIndex        =   5
      Top             =   3210
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
      ForeColor       =   192
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Cr�dito 
      Height          =   360
      Left            =   1155
      TabIndex        =   4
      Top             =   3210
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
      ForeColor       =   12582912
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Dia 
      Height          =   360
      Left            =   1170
      TabIndex        =   1
      ToolTipText     =   "Pressione F2 para Calend�rio"
      Top             =   1515
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Conta 
      Bindings        =   "LancamentosBancarios.frx":4E95A
      DataSource      =   "Data1"
      Height          =   360
      Left            =   1155
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
      DataFieldList   =   "Descri��o"
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOdd    =   16777152
      RowHeight       =   476
      Columns(0).Width=   3200
      _ExtentX        =   2143
      _ExtentY        =   635
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
      Left            =   75
      Top             =   4110
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
      Bands           =   "LancamentosBancarios.frx":4E96E
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   225
      Left            =   150
      TabIndex        =   20
      Top             =   2415
      Width           =   1005
   End
   Begin VB.Label Nome_Conta2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2415
      TabIndex        =   19
      Top             =   1530
      Width           =   5100
   End
   Begin VB.Label Atual 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1155
      TabIndex        =   17
      Top             =   3630
      Width           =   1215
   End
   Begin VB.Label Anterior 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1155
      TabIndex        =   16
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"LancamentosBancarios.frx":50DF8
      Height          =   735
      Left            =   165
      TabIndex        =   15
      Top             =   120
      Width           =   7380
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Atual"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   6
      Top             =   3675
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "D�bito"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2685
      TabIndex        =   9
      Top             =   3270
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cr�dito"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   3270
      Width           =   735
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Anterior"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   14
      Top             =   2835
      Width           =   855
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descri��o"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   11
      Top             =   1995
      Width           =   855
   End
   Begin VB.Label Nome_Conta 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   5115
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Conta"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   7
      Top             =   1125
      Width           =   735
   End
End
Attribute VB_Name = "frmLancaContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim Num_Registro_Saved As Variant
Dim Ordem As Variant
Dim rsLan�amentos As Recordset
Private rsLancClone As Recordset
Dim rsContas As Recordset
Private gbInSearching As Boolean
Private gsString As String
Private gsOrder As String

'Function Mostra_Reg(Registro As Variant)
'
'  Call StatusMsg("")
'
'  Combo_Conta.Text = ""
'  Nome_Conta.Caption = ""
'  Nome_Conta2.Caption = ""
'  Dia.Mask = ""
'  Dia.Text = ""
'  Dia.Mask = "##/##/####"
'  Descri��o.Text = ""
'  Anterior.Caption = ""
'  Cr�dito.Text = ""
'  D�bito.Text = ""
'  Atual.Caption = ""
'
'  rsLan�amentos.Bookmark = Registro
'  Num_Registro = Registro
'  Ordem = rsLan�amentos("Ordem")
'
'  Call ShowRecord
'
'End Function
'
Private Sub DeleteRecord()
  Dim Resposta As Integer
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "N�o existe registro para apagar !"
    Exit Sub
  End If
  
  Resposta = MsgBox(("Deseja realmente apagar este lan�amento" & " ?"), 20, "ATEN��O!!")
  If Resposta = 6 Then
    rsLan�amentos.Delete
    Num_Registro = Null
    Call ClearScreen
    DisplayMsg "Lan�amento apagado, n�o se esque�a de usar a tela ""Rec�lculo de Saldos""."
  End If

End Sub

Private Sub SearchRecord()
  Dim gsWhere As String
  
  If Not IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Apague todos os campos da tela com o bot�o NOVO."
    gsMsg = gsMsg & vbCrLf & "Selecione a Ordem de Pesquisa na lista e preencha com dados iniciais os campos respectivos."
    gsMsg = gsMsg & vbCrLf & "Pressione novamente este bot�o PROCURAR."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If

  gsWhere = ""
  
  Select Case ActiveBar1.Tools("miOpOrdem").CBListIndex
    Case -1, 0  'Por Conta e Data
      If Nome_Conta.Caption = "" Then
        Combo_Conta.Text = "0"
      End If
      If Not IsDate(Dia.Text) Then
        Dia.Text = Date - 3
      End If
      gsWhere = "WHERE Conta <> 0 AND Conta >= " & Val(Combo_Conta.Text) & " AND Data >=" & gsGetInvDate(Dia.Text)
    Case 1  'Por Descri��o
'      If Len(Trim(Descri��o.Text)) = 0 Then
'        DisplayMsg "Preencha no campo Descri��o o n�mero da Conta ou parte da descri��o da Conta."
'        Descri��o.SetFocus
'        Exit Sub
'      End If
      gsWhere = "WHERE Conta <> 0 AND Descri��o LIKE '" & Descri��o.Text & "*'"
  End Select
  
  Set rsLan�amentos = db.OpenRecordset("SELECT * FROM [Lan�amentos Banc�rios] " & gsWhere & " " & gsOrder, dbOpenDynaset)
  If Not rsLan�amentos.EOF Then
    Call ShowRecord
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em fun��o dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If


'  Call StatusMsg("")
'
'  If Not gbInSearching Then
'
'    If Nome_Conta.Caption = "" Then
'      DisplayMsg "Escolha a conta antes."
'      Combo_Conta.SetFocus
'      Exit Sub
'    End If
'
'    If Len(Trim(Descri��o.Text)) = 0 Then
'      DisplayMsg "Preencha no campo Descri��o o n�mero da Conta ou parte da descri��o da Conta."
'      Descri��o.SetFocus
'      Exit Sub
'    End If
'
'    gsString = Trim(Descri��o.Text)
'
'    Set rsLancClone = rsLan�amentos.Clone
'
'    If Not rsLancClone.EOF And Not rsLancClone.BOF Then
'      rsLancClone.MoveFirst
'      If Not rsLan�amentos.EOF And Not rsLan�amentos.BOF Then
'        Num_Registro_Saved = rsLan�amentos.Bookmark
'      Else
'        Num_Registro_Saved = Null
'      End If
'    End If
'
'    gbInSearching = True
'
'  End If
'
'  With rsLancClone
'    Do While Not .EOF
'      If InStr(UCase(.Fields("Descri��o").Value), UCase(gsString)) > 0 Then
'        cmdCloseSearch.Visible = True
'        Call ShowRecord
'        Exit Sub
'      End If
'      .MoveNext
'    Loop
'  End With
'
'  DisplayMsg "Fim de Pesquisa. "
'  rsLancClone.Close
'  Set rsLancClone = Nothing
'  gbInSearching = False
'  cmdCloseSearch.Visible = False
  
End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim rsLanc As Recordset
  Dim sTexto As String
  
  Call StatusMsg("")
  
  If IsNull(Cr�dito.Text) Then Cr�dito.Text = 0
  If Cr�dito.Text = "" Then Cr�dito.Text = 0
  If Not IsNumeric(Cr�dito.Text) Then Cr�dito.Text = 0
  
  If IsNull(D�bito.Text) Then D�bito.Text = 0
  If D�bito.Text = "" Then D�bito.Text = 0
  If Not IsNumeric(D�bito.Text) Then D�bito.Text = 0
  
  If IsNull(Documento.Text) Then Documento.Text = ""
  
  Call Combo_Conta_LostFocus
  
  Rem Verifica Conta
  If Nome_Conta.Caption = "" Then
    DisplayMsg "Conta inv�lida, verifique."
    Combo_Conta.SetFocus
    Exit Sub
  End If
  
  If IsNull(Dia.Text) Or Dia.Text = "" Or Not IsDate(Dia.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Dia.SetFocus
    Exit Sub
  End If
  
  If IsNull(Descri��o.Text) Or Descri��o.Text = "" Then
    DisplayMsg "Descri��o inv�lida, verifique."
    Descri��o.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Cr�dito.Text) Then Erro = True
  If Erro = False Then If Not IsNumeric(Cr�dito.Text) Then Erro = False
  If Erro = False Then If CDbl(Cr�dito.Text) < 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "Cr�dito incorreto, verifique."
    Cr�dito.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(D�bito.Text) Then Erro = True
  If Erro = False Then If Not IsNumeric(D�bito.Text) Then Erro = False
  If Erro = False Then If CDbl(D�bito.Text) < 0 Then Erro = True
  If Erro = True Then
    DisplayMsg "D�bito incorreto, verifique."
    D�bito.SetFocus
    Exit Sub
  End If
  
  If CDbl(Cr�dito.Text) = 0 And CDbl(D�bito.Text) = 0 Then
    DisplayMsg "Digite um valor em cr�dito ou em d�bito."
    Cr�dito.SetFocus
    Exit Sub
  End If
  
  If CDbl(Cr�dito.Text) <> 0 And CDbl(D�bito.Text) <> 0 Then
    DisplayMsg "Digite apenas um valor: cr�dito ou em d�bito."
    Cr�dito.SetFocus
    Exit Sub
  End If
  
  Rem Se for um reg. novo acha o saldo anterior
  If IsNull(Num_Registro) Then
    Set rsLanc = db.OpenRecordset("SELECT * FROM [Lan�amentos Banc�rios] WHERE Conta = " & Combo_Conta.Text & " ORDER BY Data, Ordem", dbOpenDynaset)
    rsLanc.FindLast "Data <= #" & Mid(Dia.Text, 4, 2) & "/" & Mid(Dia.Text, 1, 2) & "/" & Mid(Dia.Text, 7, 4) & "#"
    If Not rsLanc.NoMatch Then
       Anterior.Caption = rsLanc("Saldo Atual")
    Else
      Anterior.Caption = "0"
    End If
    rsLanc.Close
    Set rsLanc = Nothing
  End If
  
  Atual.Caption = Format(CDbl(Anterior.Caption) + CDbl(Cr�dito.Text) - CDbl(D�bito.Text), FORMAT_VALUE)
  
  Call StatusMsg("Gravando ...")
  
  With rsLan�amentos
  
    If IsNull(Num_Registro) Then
      .AddNew
      Num_Registro = !Ordem
      sTexto = "Lan�amento efetuado."
      .Fields("Conta") = Val(Combo_Conta.Text)
    Else
      .Edit
      sTexto = "Lan�amento alterado."
    End If
    
    .Fields("Data") = Dia.Text
    .Fields("Descri��o") = Descri��o.Text
    .Fields("Saldo Anterior") = CDbl(Anterior.Caption)
    .Fields("D�bito") = CDbl(D�bito.Text)
    .Fields("Cr�dito") = CDbl(Cr�dito.Text)
    .Fields("Saldo Atual") = Format(CDbl(Atual.Caption), "###########0.00")
    .Fields("Cheque") = Documento.Text
    
    .Update
'    Num_Registro = .LastModified
'    .Bookmark = Num_Registro
    .Bookmark = .LastModified
    
  Call StatusMsg("")
  
  End With
  
End Sub

Private Sub ClearScreen()
  Call StatusMsg("")
  
  Combo_Conta.Text = ""
  Nome_Conta.Caption = ""
  Nome_Conta2.Caption = ""
  Dia.Mask = ""
  Dia.Text = ""
  Dia.Mask = "##/##/####"
  Descri��o.Text = ""
  Anterior.Caption = ""
  Cr�dito.Text = ""
  D�bito.Text = ""
  Atual.Caption = ""
  
  Documento.Text = ""
  
  If Not rsLan�amentos.EOF Then
    On Error Resume Next
    rsLan�amentos.MoveFirst
    rsLan�amentos.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  Ordem = Null
  
  Combo_Conta.SetFocus
  
End Sub

'Private Sub MoveFirst()
'  On Error Resume Next
'  If Not gbInSearching Then
'    With rsLan�amentos
'      .MoveFirst
'      If .BOF Then
'        Beep
'      Else
'        Call ShowRecord
'      End If
'    End With
'  Else
'    With rsLancClone
'      With rsLancClone
'        Num_Registro = .Bookmark
'        Do While Not .BOF
'          .MovePrevious
'          If InStr(UCase(.Fields("Descri��o").Value), UCase(gsString)) > 0 Then
'            Num_Registro = .Bookmark
'          End If
'        Loop
'        .Bookmark = Num_Registro
'        cmdCloseSearch.Visible = True
'        Call ShowRecord
'      End With
'    End With
'  End If
'End Sub
'
'Private Sub MoveLast()
'  On Error Resume Next
'  If Not gbInSearching Then
'    With rsLan�amentos
'      .MoveLast
'      If .EOF Then
'        Beep
'      Else
'        Call ShowRecord
'      End If
'    End With
'  Else
'    With rsLancClone
'      With rsLancClone
'        Num_Registro = .Bookmark
'        Do While Not .EOF
'          .MoveNext
'          If InStr(UCase(.Fields("Descri��o").Value), UCase(gsString)) > 0 Then
'            Num_Registro = .Bookmark
'          End If
'        Loop
'        .Bookmark = Num_Registro
'        cmdCloseSearch.Visible = True
'        Call ShowRecord
'      End With
'    End With
'  End If
'End Sub
'
'Private Sub MovePrevious()
'  On Error Resume Next
'  If Not gbInSearching Then
'    With rsLan�amentos
'      .MovePrevious
'      If Not .BOF Then
'        Call ShowRecord
'      Else
'        Beep
'        .MoveNext
'      End If
'    End With
'  Else
'    With rsLancClone
'      With rsLancClone
'        Do While Not .BOF
'          .MovePrevious
'          If InStr(UCase(.Fields("Descri��o").Value), UCase(gsString)) > 0 Then
'            cmdCloseSearch.Visible = True
'            Call ShowRecord
'            Exit Sub
'          End If
'        Loop
'      End With
'    End With
'  End If
'End Sub
'
'Private Sub MoveNext()
'  On Error Resume Next
'  If Not gbInSearching Then
'    With rsLan�amentos
'      .MoveNext
'      If Not .EOF Then
'        Call ShowRecord
'      Else
'        Beep
'        .MovePrevious
'      End If
'    End With
'  Else
'    With rsLancClone
'      With rsLancClone
'        Do While Not .EOF
'          .MoveNext
'          If InStr(UCase(.Fields("Descri��o").Value), UCase(gsString)) > 0 Then
'            cmdCloseSearch.Visible = True
'            Call ShowRecord
'            Exit Sub
'          End If
'        Loop
'      End With
'    End With
'  End If
'End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsLan�amentos
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
  With rsLan�amentos
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
  With rsLan�amentos
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
  With rsLan�amentos
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
      Call SearchRecord
  End Select
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  gsOrder = ""
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case -1, 0
          gsOrder = "ORDER BY Data, Conta, Ordem"
        Case 1
          gsOrder = "ORDER BY Descri��o, Conta, Ordem"
      End Select
  End Select

End Sub

'Private Sub cmdCloseSearch_Click()
'  rsLancClone.Close
'  Set rsLancClone = Nothing
'  cmdCloseSearch.Visible = False
'  gbInSearching = False
'  If Not IsNull(Num_Registro_Saved) Then
'    rsLan�amentos.Bookmark = Num_Registro_Saved
'  End If
'End Sub

Private Sub Combo_Conta_CloseUp()
  Combo_Conta.Text = Combo_Conta.Columns(2).Text
  Combo_Conta_LostFocus
End Sub

Private Sub Combo_Conta_KeyPress(KeyAscii As Integer)
  If Not Combo_Conta.DroppedDown Then
    '05/12/2006 - Anderson
    'Altera��o para aumentar o n�mero de contas banc�rias
    'KeyAscii = gnLimitKeyPress(Combo_Conta, 2, KeyAscii, True)
    KeyAscii = gnLimitKeyPress(Combo_Conta, 3, KeyAscii, True)
  End If
End Sub

Private Sub Combo_Conta_LostFocus()
  Nome_Conta.Caption = ""
  If IsNull(Combo_Conta.Text) Then Exit Sub
  If Not IsNumeric(Combo_Conta.Text) Then Exit Sub
  
  '05/12/2006 - Anderson
  'Altera��o para aumentar o n�mero de contas banc�rias
  'If Val(Combo_Conta.Text) < 0 Or Val(Combo_Conta.Text) > 999999 Then Exit Sub
  If Val(Combo_Conta.Text) < 0 Or Val(Combo_Conta.Text) > 255 Then Exit Sub

  rsContas.Index = "C�digo"
  rsContas.Seek "=", Val(Combo_Conta.Text)
  If rsContas.NoMatch Then Exit Sub
  Nome_Conta.Caption = rsContas("Descri��o")
  Nome_Conta2.Caption = rsContas("Conta")

End Sub

Private Sub Cr�dito_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub D�bito_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub Dia_LostFocus()
  Dia.Text = Ajusta_Data(Dia.Text)
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
  
  With ActiveBar1.Tools("miOpOrdem").CBList
    .Clear
    .AddItem "Por Conta e Data"
    .AddItem "Por Descri��o"
  End With
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  ActiveBar1.RecalcLayout
  ActiveBar1.Refresh
  
  gsOrder = "ORDER BY Data, Conta, Ordem"
  
  Set rsContas = db.OpenRecordset("Contas Banc�rias", , dbReadOnly)
  Set rsLan�amentos = db.OpenRecordset("SELECT * FROM [Lan�amentos Banc�rios] WHERE Conta <> 0 " & gsOrder, dbOpenDynaset)
  
  Data1.DatabaseName = gsQuickDBFileName

  Num_Registro = Null
  gbInSearching = False
  Call ActiveBarLoadToolTips(Me)
  
  Me.Show
  DoEvents
  
  Call ClearScreen
  
End Sub

Private Sub ShowRecord()

'  If Not gbInSearching Then
    Combo_Conta.Text = rsLan�amentos("Conta")
    Dia.Text = Format(rsLan�amentos("Data"), "dd/mm/yyyy")
    Descri��o.Text = rsLan�amentos("Descri��o") & ""
    Anterior.Caption = Format(rsLan�amentos("Saldo Anterior"), FORMAT_VALUE)
    Cr�dito.Text = rsLan�amentos("Cr�dito")
    D�bito.Text = rsLan�amentos("D�bito")
    Atual.Caption = Format(rsLan�amentos("Saldo Atual"), FORMAT_VALUE)
    Documento.Text = rsLan�amentos("Cheque") & ""
    Combo_Conta_LostFocus
'    Num_Registro = rsLan�amentos.Bookmark
    Num_Registro = rsLan�amentos!Ordem
'  Else
'    Combo_Conta.Text = rsLancClone("Conta")
'    Dia.Text = Format(rsLancClone("Data"), "dd/mm/yyyy")
'    Descri��o.Text = rsLancClone("Descri��o") & ""
'    Anterior.Caption = Format(rsLancClone("Saldo Anterior"), "###,###,###,##0.00")
'    Cr�dito.Text = rsLancClone("Cr�dito")
'    D�bito.Text = rsLancClone("D�bito")
'    Atual.Caption = Format(rsLancClone("Saldo Atual"), "###,###,###,##0.00")
'    Documento.Text = rsLancClone("Cheque") & ""
'    Combo_Conta_LostFocus
'    Num_Registro = rsLancClone.Bookmark
'  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsContas.Close
  rsLan�amentos.Close
  Set rsContas = Nothing
  Set rsLan�amentos = Nothing
End Sub

'Private Sub O_Banco_Click()
'  Set rsLan�amentos = db.OpenRecordset("SELECT * FROM [Lan�amentos Banc�rios] ORDER BY Data, Conta, Ordem", dbOpenDynaset)
'End Sub
Private Sub Label8_Click()

End Sub
