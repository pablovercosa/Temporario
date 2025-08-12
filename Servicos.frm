VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmServicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Serviços"
   ClientHeight    =   7275
   ClientLeft      =   3240
   ClientTop       =   2325
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Servicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   6915
   Begin SSDataWidgets_B.SSDBDropDown DropDownCFOP 
      Bindings        =   "Servicos.frx":4E95A
      Height          =   1185
      Left            =   480
      TabIndex        =   16
      Top             =   4440
      Width           =   5565
      DataFieldList   =   "Nome"
      ListAutoValidate=   0   'False
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOdd    =   12648447
      RowHeight       =   423
      ExtraHeight     =   185
      Columns.Count   =   3
      Columns(0).Width=   5609
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2090
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   1773
      Columns(2).Caption=   "CFOP"
      Columns(2).Name =   "CFOP"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Código Fiscal"
      Columns(2).FieldLen=   256
      _ExtentX        =   9816
      _ExtentY        =   2090
      _StockProps     =   77
   End
   Begin VB.Data DataCFOP2 
      Caption         =   "Data4"
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
      Left            =   4680
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código, [Código Fiscal] FROM [Operações Saída] ORDER BY Nome"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Data DataCFOP1 
      Caption         =   "Data4"
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
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkPublicidade 
      Appearance      =   0  'Flat
      Caption         =   "Trata-se de um Serviço de Propagandas ou Publicidades"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   1590
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comissão "
      Height          =   990
      Left            =   180
      TabIndex        =   14
      Top             =   1950
      Width           =   6615
      Begin VB.CheckBox Comissão_Sobre 
         Appearance      =   0  'Flat
         Caption         =   "Comissão do serviço sobrepõe comissão do vendedor / técnico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Width           =   4965
      End
      Begin MSMask.MaskEdBox Comissão 
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%  Comissão"
         Height          =   195
         Left            =   780
         TabIndex        =   15
         Top             =   600
         Width           =   930
      End
   End
   Begin MSMask.MaskEdBox ISS 
      Height          =   315
      Left            =   3015
      TabIndex        =   3
      Top             =   1200
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Preço 
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   1200
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
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
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   690
      Left            =   4350
      TabIndex        =   11
      Top             =   0
      Width           =   2430
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1380
         TabIndex        =   8
         Top             =   270
         Width           =   795
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   780
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Pressione F5 para o Próximo Livre."
      Top             =   210
      Width           =   900
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   780
      MaxLength       =   60
      TabIndex        =   1
      Top             =   780
      Width           =   6015
   End
   Begin SSDataWidgets_B.SSDBGrid GradeCFOP 
      Bindings        =   "Servicos.frx":4E972
      Height          =   3690
      Left            =   180
      TabIndex        =   17
      Top             =   3000
      Width           =   6615
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      SelectTypeRow   =   1
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   11668
      _ExtentY        =   6509
      _StockProps     =   79
      Caption         =   "Operação X CFOP"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   6300
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
      Bands           =   "Servicos.frx":4E98A
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "% ISS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   13
      Top             =   1260
      Width           =   420
   End
   Begin VB.Label L_Comiss1 
      Caption         =   "Preço"
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   1215
      Width           =   540
   End
   Begin VB.Label Label9 
      Caption         =   "Código"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   210
      Width           =   525
   End
   Begin VB.Label Label11 
      Caption         =   "Nome"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   795
      Width           =   525
   End
End
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Num_Registro As Variant
Private rsServicos As Recordset
Private rsParametros As Recordset
'15/12/2006 - Anderson - Alteração para o cadastro de CFOP por serviço
Dim Rec_CFOP As Recordset
Dim rsOperacaoSaida As Recordset
Dim rsServicoCFOP As Recordset

Private Sub ShowRecord()
  Dim Aux_Str As String
  
  On Error GoTo Processa_Erro
  
  Call StatusMsg("")
  
  Código.Text = rsServicos("Código")
  Código.Locked = True
   
  Nome.Text = rsServicos("Descrição") & ""

  Preço.Text = Format(rsServicos("Preço") & "", "########0.00")
  
  '07/04/2004 - Daniel
  'Case: STC de Caxias do Sul
  If rsServicos.Fields("Publicidade").Value = True Then
    chkPublicidade.Value = vbChecked
  Else
    chkPublicidade.Value = vbUnchecked
  End If
  
  Comissão_Sobre.Value = -rsServicos("Comissão Sobrepõe")
  Comissão.Text = rsServicos("Comissão")
  
  ISS.Text = rsServicos("ISS") & ""
  
  Num_Registro = rsServicos.Bookmark
  
  '20/12/2004 - Anderson - Alteração para o registro de CFOP por Serviço
  Call LoadOperacao
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar o registro em Serviços."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

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
  '15/12/2006 - Anderson
  Call LoadOperacao
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsServicos
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
  With rsServicos
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
  With rsServicos
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
  With rsServicos
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
  Dim rsComissão_Serviço As Recordset
  Dim rsSaidas_Serviço As Recordset
  
  Dim Aux_Filial As Integer
  Dim Aux_Sequência As Long
  Dim Aux_Contador As Long
  
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "Não existe nenhum serviço para ser apagado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If rsServicos("Código") = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "O serviço 0 não pode ser apagado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Atenção, caso você já tenha usado este serviço NÃO é aconselhável apagá-lo. Alguns relatórios poderão mostrar informações incorretas. Deseja apagar este serviço ?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    gsTitle = LoadResString(201)
    gsMsg = "Serviço não apagado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  
  Call StatusMsg("Aguarde, atualizando comissões...")
  
  Call ws.BeginTrans
  
  Set rsComissão_Serviço = db.OpenRecordset("Comissão Serviços")
  Set rsSaidas_Serviço = db.OpenRecordset("Saídas - Serviços")
  
  rsComissão_Serviço.Index = "Sequência"
  Aux_Filial = 0
  Aux_Sequência = 0
  Aux_Contador = 0
  
Lp_Comissão:
  rsComissão_Serviço.Seek ">", Aux_Filial, Aux_Sequência, Aux_Contador
  If rsComissão_Serviço.NoMatch Then GoTo Ver_Saídas
  Aux_Filial = rsComissão_Serviço("Filial")
  Aux_Sequência = rsComissão_Serviço("Sequência")
  Aux_Contador = rsComissão_Serviço("Contador")
  
  If rsComissão_Serviço("Serviço") = rsServicos("Código") Then
    rsComissão_Serviço.Edit
    rsComissão_Serviço("Serviço") = 0
    rsComissão_Serviço.Update
  End If
  GoTo Lp_Comissão
  
  
Ver_Saídas:
  rsSaidas_Serviço.Index = "Sequência"
  Aux_Filial = 0
  Aux_Sequência = 0
  Aux_Contador = 0
    
Lp_Saídas:
  rsSaidas_Serviço.Seek ">", Aux_Filial, Aux_Sequência, Aux_Contador
  If rsSaidas_Serviço.NoMatch Then GoTo Fim
  Aux_Filial = rsSaidas_Serviço("Filial")
  Aux_Sequência = rsSaidas_Serviço("Sequência")
  Aux_Contador = rsSaidas_Serviço("Linha")
  
  If rsSaidas_Serviço("Código") = rsServicos("Código") Then
    rsSaidas_Serviço.Edit
    rsSaidas_Serviço("Código") = 0
    rsSaidas_Serviço.Update
  End If
  
  GoTo Lp_Saídas
  
  
Fim:
  rsServicos.Delete
  
  rsComissão_Serviço.Close
  rsSaidas_Serviço.Close
  
  Set rsComissão_Serviço = Nothing
  Set rsSaidas_Serviço = Nothing
  
  Call ws.CommitTrans
  
  Call ClearScreen

End Sub

Private Sub UpdateRecord()
  Dim sTexto As String
  Dim Erro As Integer
  
  Dim dblPreco As Double
  
  
  On Error GoTo ErrHandler
  
  Código.Text = gsHandleNull(Código.Text & "")
  If CLng(Código.Text) <= 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Código incorreto, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Código.SetFocus
    Exit Sub
  End If
  
  If Comissão_Sobre.Value = 0 Then Comissão.Text = 0
  Erro = False
  If IsNull(Comissão.Text) Then Erro = True
  If Erro = False Then If Comissão.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Comissão.Text) Then Erro = True
  If Erro = False Then
     If CDbl(Comissão.Text) < 0 Or Comissão.Text >= 99 Then Erro = True
  End If
  
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Comissão deve ficar entre 0 e 99."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Comissão.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(Nome.Text)) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Descrição incorreta, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Nome.SetFocus
    Exit Sub
  End If
  
  If IsNull(ISS.Text) Then ISS.Text = 0
  If ISS.Text = "" Then ISS.Text = 0
  If Not IsNumeric(ISS.Text) Then ISS.Text = 0
  If CDbl(ISS.Text) < 0 Or ISS.Text > 99 Then ISS.Text = 0
  
  
  '19/01/2004 - mpdea
  'Validação para o AM
  '
  'Verifica o cadastramento de preços > 0
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
    Call IsDataType(dtDouble, CDbl(Preço.Text), dblPreco)
    
    If dblPreco <= 0 Then
      MsgBox "Preço deve ser maior do que zero.", vbExclamation, "Atenção"
      SelectAllText Preço, True
      Exit Sub
    End If
  End If
  
  Call StatusMsg("Gravando ...")
  
  With rsServicos
  
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código") = Val(Código.Text)
      sTexto = "Registro inserido"
    Else
      .Edit
      sTexto = "Registro alterado"
    End If
    
    .Fields("Descrição") = Nome.Text
    .Fields("Preço") = gsHandleNull(Preço.Text)
    .Fields("Comissão Sobrepõe") = (Comissão_Sobre.Value = 1)
    .Fields("Comissão") = CDbl(Comissão.Text)
    .Fields("ISS") = CDbl(ISS.Text)
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
    '07/04/2004 - Daniel
    'Case: STC de Caxias do Sul
    If chkPublicidade.Value = vbChecked Then
      .Fields("Publicidade").Value = True
    Else
      .Fields("Publicidade").Value = False
    End If
    
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
    
  End With
  
  Num_Registro = rsServicos.Bookmark
  
  Código.Locked = True
  
  Call StatusMsg(sTexto & " com sucesso.")
  
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro na tentativa de gravação: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub


Public Sub ClearScreen()

  Call StatusMsg("")

  Código.Text = ""
  Nome.Text = ""
  Preço.Text = 0
  Comissão_Sobre.Value = 0
  Comissão.Text = 0
  ISS.Text = ""
  '07/04/2004 - Daniel
  'Case: STC de Caxias do Sul
  chkPublicidade.Value = vbUnchecked
  
  If Not rsServicos.EOF Then
    On Error Resume Next
    rsServicos.MoveFirst
    rsServicos.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  Código.Locked = False
  Código.SetFocus
  
End Sub

Private Sub Código_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    Call O_Código_Click
    Call GetNewCode(Me, rsServicos, 9999)
  End If
End Sub

Private Sub Código_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub Código_LostFocus()
  If IsNull(Código.Text) Then Exit Sub
  If Código.Text = "" Then Exit Sub
  If Val(Código.Text) <= 0 Then Exit Sub
  
  rsServicos.FindFirst "Código = " & Código.Text
  If Not rsServicos.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If
End Sub

Private Sub Comissão_Sobre_Click()
  Comissão.Enabled = (Comissão_Sobre.Value = 1)
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
  '15/12/2006 - Anderson
  'Alteração para o cadastro de CFOP por servico.
  DataCFOP1.DatabaseName = gsQuickDBFileName
  DataCFOP2.DatabaseName = gsQuickDBFileName
  
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  KeyPreview = True
  
  Set rsServicos = db.OpenRecordset("SELECT * FROM Serviços WHERE Código <> 0 ORDER BY Código", dbOpenDynaset)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  '15/12/2006 - Anderson
  'Alteração para o cadastro de CFOP por servico
  Set rsOperacaoSaida = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsServicoCFOP = db.OpenRecordset("ServicoCFOP", , dbReadOnly)
  
  Me.Show
  DoEvents
  
  Call ActiveBarLoadToolTips(Me)
  
  Call ClearScreen

  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call StatusMsg("")
  On Error Resume Next
  rsServicos.Close
  rsParametros.Close
  '15/12/2006 - Anderson - Alteração para o cadastro de CFOP por servico
  rsOperacaoSaida.Close
  rsServicoCFOP.Close
  
  Set rsServicos = Nothing
  Set rsParametros = Nothing
  '15/12/2006 - Anderson - Alteração para o cadastro de CFOP por servico
  Set rsOperacaoSaida = Nothing
  Set rsServicoCFOP = Nothing

  On Error GoTo 0
End Sub

Private Sub O_Código_Click()
  Set rsServicos = db.OpenRecordset("SELECT * FROM Serviços WHERE Código <> 0 ORDER BY Código", dbOpenDynaset)
End Sub

Private Sub O_Nome_Click()
  Set rsServicos = db.OpenRecordset("SELECT * FROM Serviços WHERE Código <> 0 ORDER BY Descrição", dbOpenDynaset)
End Sub

Private Sub Preço_GotFocus()
  Preço.SelStart = 0
  Preço.SelLength = Preço.MaxLength
End Sub

'15/12/2006 - Anderson
'Alteração para o cadastro de CFOP por Servico
Private Sub DropDownCFOP_Click()

  GradeCFOP.Columns(1).Text = DropDownCFOP.Columns(1).Text
  GradeCFOP.Columns(2).Text = DropDownCFOP.Columns(0).Text

End Sub

'15/12/2006 - Anderson
'Alteração para o cadastro de CFOP por serviço
Private Sub GradeCFOP_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 Dim Resp As Integer
  
  Call StatusMsg("")
  Aux = GradeCFOP.Columns(ColIndex).Text
  If ColIndex = 1 Then  'Código
    If IsNull(Aux) Then
      Cancel = True
      Exit Sub
    End If
    If Not IsNumeric(Aux) Then
      Cancel = True
      Exit Sub
    End If
    If Val(Aux) > 99999999 Or Val(Aux) < 1 Then
      Cancel = True
      Exit Sub
    End If
    
    rsOperacaoSaida.Index = "Código"
    rsOperacaoSaida.Seek "=", Val(Aux)
    If rsOperacaoSaida.NoMatch Then
       MsgBox ("Código de operação inválido.")
       Cancel = True
       Exit Sub
    End If
    
    Rem Verifica se esta operação já está neste servico
    rsServicoCFOP.Index = "PrimaryKey"
    rsServicoCFOP.Seek "=", Código.Text, Val(Aux)
    If Not rsServicoCFOP.NoMatch Then
      MsgBox "Esta operação já está cadastrada neste serviço.", vbExclamation
      Cancel = True
      Exit Sub
    End If
    
  End If
  
  GradeCFOP.Columns(0).Text = Código.Text
  GradeCFOP.Columns(2).Text = rsOperacaoSaida("Nome")

End Sub

'15/12/2006 - Anderson
'Alteração para o cadastro de CFOP por serviço
Private Sub GradeCFOP_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)

  Dim Resposta As Integer
  
  Call StatusMsg("")
  Resposta = MsgBox("Deseja apagar a operação selecionada? ", 260, "Atenção")
  If Resposta = 7 Then
    DisplayMsg "Operação não apagada."
    Cancel = True
  End If
  
  DispPromptMsg = False

End Sub

'15/12/2006 - Anderson
'Alteração para o cadastro de CFOP por serviço
Private Sub GradeCFOP_LostFocus()

 If GradeCFOP.RowChanged = True Then
   GradeCFOP.Update
 End If

End Sub


'-----------------------------------------------------
'15/12/2006 - Anderson
'Alteração para o cadastro de CFOP por serviço
Private Sub LoadOperacao()
  Dim Cód As String
  Dim sSql As String

  'Arruma Grade
  Cód = CInt("0" & Código)
  
    sSql = "SELECT CodServico, CodOperacao, [Operações Saída].Nome, CFOP FROM ServicoCFOP "
    sSql = sSql + " INNER JOIN [Operações Saída] ON ServicoCFOP.CodOperacao = [Operações Saída].Código  WHERE CodServico =" + Cód + " "
    
    GradeCFOP.Visible = False
    GradeCFOP.DataMode = 1
  
    Set Rec_CFOP = db.OpenRecordset(sSql, dbOpenDynaset)
  
    Set DataCFOP1.Recordset = Rec_CFOP
    
  If Cód = "0" Then

    GradeCFOP.Enabled = False
    'Exit Sub
  Else
    GradeCFOP.Enabled = True
  
  End If
  
  'GradeCFOP.Enabled = True
  
  GradeCFOP.DataMode = 0

  GradeCFOP.ReBind
 
  GradeCFOP.Columns(0).Visible = False 'Cod Prod
  GradeCFOP.Columns(1).Width = 1000 'Cod Operação
  GradeCFOP.Columns(1).Caption = "Cód. Operação"
  GradeCFOP.Columns(2).Locked = True 'Nome da Operação
  GradeCFOP.Columns(2).Width = 4000
  GradeCFOP.Columns(3).Width = 1000 'CFOP
  
  GradeCFOP.Visible = True
  GradeCFOP.Columns(1).DropDownHwnd = DropDownCFOP.hwnd
  '-----------------------------------------------------
  
'  Nome_Foto_LostFocus
  
  Call StatusMsg("")

End Sub

