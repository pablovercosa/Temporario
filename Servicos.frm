VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmServicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Servi�os"
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
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   1773
      Columns(2).Caption=   "CFOP"
      Columns(2).Name =   "CFOP"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "C�digo Fiscal"
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
      RecordSource    =   "SELECT Nome, C�digo, [C�digo Fiscal] FROM [Opera��es Sa�da] ORDER BY Nome"
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
      Caption         =   "Trata-se de um Servi�o de Propagandas ou Publicidades"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   1590
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comiss�o "
      Height          =   990
      Left            =   180
      TabIndex        =   14
      Top             =   1950
      Width           =   6615
      Begin VB.CheckBox Comiss�o_Sobre 
         Appearance      =   0  'Flat
         Caption         =   "Comiss�o do servi�o sobrep�e comiss�o do vendedor / t�cnico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Width           =   4965
      End
      Begin MSMask.MaskEdBox Comiss�o 
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
         Caption         =   "%  Comiss�o"
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
   Begin MSMask.MaskEdBox Pre�o 
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
      Begin VB.OptionButton O_C�digo 
         Appearance      =   0  'Flat
         Caption         =   "C�digo"
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
   Begin VB.TextBox C�digo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   780
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Pressione F5 para o Pr�ximo Livre."
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
      Caption         =   "Opera��o X CFOP"
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
      Caption         =   "Pre�o"
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   1215
      Width           =   540
   End
   Begin VB.Label Label9 
      Caption         =   "C�digo"
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
'15/12/2006 - Anderson - Altera��o para o cadastro de CFOP por servi�o
Dim Rec_CFOP As Recordset
Dim rsOperacaoSaida As Recordset
Dim rsServicoCFOP As Recordset

Private Sub ShowRecord()
  Dim Aux_Str As String
  
  On Error GoTo Processa_Erro
  
  Call StatusMsg("")
  
  C�digo.Text = rsServicos("C�digo")
  C�digo.Locked = True
   
  Nome.Text = rsServicos("Descri��o") & ""

  Pre�o.Text = Format(rsServicos("Pre�o") & "", "########0.00")
  
  '07/04/2004 - Daniel
  'Case: STC de Caxias do Sul
  If rsServicos.Fields("Publicidade").Value = True Then
    chkPublicidade.Value = vbChecked
  Else
    chkPublicidade.Value = vbUnchecked
  End If
  
  Comiss�o_Sobre.Value = -rsServicos("Comiss�o Sobrep�e")
  Comiss�o.Text = rsServicos("Comiss�o")
  
  ISS.Text = rsServicos("ISS") & ""
  
  Num_Registro = rsServicos.Bookmark
  
  '20/12/2004 - Anderson - Altera��o para o registro de CFOP por Servi�o
  Call LoadOperacao
  
  Exit Sub
  
Processa_Erro:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apresentar o registro em Servi�os."
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
  Dim rsComiss�o_Servi�o As Recordset
  Dim rsSaidas_Servi�o As Recordset
  
  Dim Aux_Filial As Integer
  Dim Aux_Sequ�ncia As Long
  Dim Aux_Contador As Long
  
  
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    gsTitle = LoadResString(201)
    gsMsg = "N�o existe nenhum servi�o para ser apagado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  If rsServicos("C�digo") = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "O servi�o 0 n�o pode ser apagado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Aten��o, caso voc� j� tenha usado este servi�o N�O � aconselh�vel apag�-lo. Alguns relat�rios poder�o mostrar informa��es incorretas. Deseja apagar este servi�o ?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    gsTitle = LoadResString(201)
    gsMsg = "Servi�o n�o apagado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  
  Call StatusMsg("Aguarde, atualizando comiss�es...")
  
  Call ws.BeginTrans
  
  Set rsComiss�o_Servi�o = db.OpenRecordset("Comiss�o Servi�os")
  Set rsSaidas_Servi�o = db.OpenRecordset("Sa�das - Servi�os")
  
  rsComiss�o_Servi�o.Index = "Sequ�ncia"
  Aux_Filial = 0
  Aux_Sequ�ncia = 0
  Aux_Contador = 0
  
Lp_Comiss�o:
  rsComiss�o_Servi�o.Seek ">", Aux_Filial, Aux_Sequ�ncia, Aux_Contador
  If rsComiss�o_Servi�o.NoMatch Then GoTo Ver_Sa�das
  Aux_Filial = rsComiss�o_Servi�o("Filial")
  Aux_Sequ�ncia = rsComiss�o_Servi�o("Sequ�ncia")
  Aux_Contador = rsComiss�o_Servi�o("Contador")
  
  If rsComiss�o_Servi�o("Servi�o") = rsServicos("C�digo") Then
    rsComiss�o_Servi�o.Edit
    rsComiss�o_Servi�o("Servi�o") = 0
    rsComiss�o_Servi�o.Update
  End If
  GoTo Lp_Comiss�o
  
  
Ver_Sa�das:
  rsSaidas_Servi�o.Index = "Sequ�ncia"
  Aux_Filial = 0
  Aux_Sequ�ncia = 0
  Aux_Contador = 0
    
Lp_Sa�das:
  rsSaidas_Servi�o.Seek ">", Aux_Filial, Aux_Sequ�ncia, Aux_Contador
  If rsSaidas_Servi�o.NoMatch Then GoTo Fim
  Aux_Filial = rsSaidas_Servi�o("Filial")
  Aux_Sequ�ncia = rsSaidas_Servi�o("Sequ�ncia")
  Aux_Contador = rsSaidas_Servi�o("Linha")
  
  If rsSaidas_Servi�o("C�digo") = rsServicos("C�digo") Then
    rsSaidas_Servi�o.Edit
    rsSaidas_Servi�o("C�digo") = 0
    rsSaidas_Servi�o.Update
  End If
  
  GoTo Lp_Sa�das
  
  
Fim:
  rsServicos.Delete
  
  rsComiss�o_Servi�o.Close
  rsSaidas_Servi�o.Close
  
  Set rsComiss�o_Servi�o = Nothing
  Set rsSaidas_Servi�o = Nothing
  
  Call ws.CommitTrans
  
  Call ClearScreen

End Sub

Private Sub UpdateRecord()
  Dim sTexto As String
  Dim Erro As Integer
  
  Dim dblPreco As Double
  
  
  On Error GoTo ErrHandler
  
  C�digo.Text = gsHandleNull(C�digo.Text & "")
  If CLng(C�digo.Text) <= 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "C�digo incorreto, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    C�digo.SetFocus
    Exit Sub
  End If
  
  If Comiss�o_Sobre.Value = 0 Then Comiss�o.Text = 0
  Erro = False
  If IsNull(Comiss�o.Text) Then Erro = True
  If Erro = False Then If Comiss�o.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Comiss�o.Text) Then Erro = True
  If Erro = False Then
     If CDbl(Comiss�o.Text) < 0 Or Comiss�o.Text >= 99 Then Erro = True
  End If
  
  If Erro = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Comiss�o deve ficar entre 0 e 99."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Comiss�o.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(Nome.Text)) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Descri��o incorreta, verifique."
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
  'Valida��o para o AM
  '
  'Verifica o cadastramento de pre�os > 0
  If UCase(gstrGetEstadoFilial(gnCodFilial)) = "AM" Then
    Call IsDataType(dtDouble, CDbl(Pre�o.Text), dblPreco)
    
    If dblPreco <= 0 Then
      MsgBox "Pre�o deve ser maior do que zero.", vbExclamation, "Aten��o"
      SelectAllText Pre�o, True
      Exit Sub
    End If
  End If
  
  Call StatusMsg("Gravando ...")
  
  With rsServicos
  
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("C�digo") = Val(C�digo.Text)
      sTexto = "Registro inserido"
    Else
      .Edit
      sTexto = "Registro alterado"
    End If
    
    .Fields("Descri��o") = Nome.Text
    .Fields("Pre�o") = gsHandleNull(Pre�o.Text)
    .Fields("Comiss�o Sobrep�e") = (Comiss�o_Sobre.Value = 1)
    .Fields("Comiss�o") = CDbl(Comiss�o.Text)
    .Fields("ISS") = CDbl(ISS.Text)
    .Fields("Data Altera��o") = Format(Date, "dd/mm/yyyy")
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
  
  C�digo.Locked = True
  
  Call StatusMsg(sTexto & " com sucesso.")
  
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro na tentativa de grava��o: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub
  
End Sub


Public Sub ClearScreen()

  Call StatusMsg("")

  C�digo.Text = ""
  Nome.Text = ""
  Pre�o.Text = 0
  Comiss�o_Sobre.Value = 0
  Comiss�o.Text = 0
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
  C�digo.Locked = False
  C�digo.SetFocus
  
End Sub

Private Sub C�digo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    Call O_C�digo_Click
    Call GetNewCode(Me, rsServicos, 9999)
  End If
End Sub

Private Sub C�digo_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub C�digo_LostFocus()
  If IsNull(C�digo.Text) Then Exit Sub
  If C�digo.Text = "" Then Exit Sub
  If Val(C�digo.Text) <= 0 Then Exit Sub
  
  rsServicos.FindFirst "C�digo = " & C�digo.Text
  If Not rsServicos.NoMatch Then
    Call ShowRecord
  Else
    Num_Registro = Null
  End If
End Sub

Private Sub Comiss�o_Sobre_Click()
  Comiss�o.Enabled = (Comiss�o_Sobre.Value = 1)
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
  'Altera��o para o cadastro de CFOP por servico.
  DataCFOP1.DatabaseName = gsQuickDBFileName
  DataCFOP2.DatabaseName = gsQuickDBFileName
  
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  KeyPreview = True
  
  Set rsServicos = db.OpenRecordset("SELECT * FROM Servi�os WHERE C�digo <> 0 ORDER BY C�digo", dbOpenDynaset)
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  '15/12/2006 - Anderson
  'Altera��o para o cadastro de CFOP por servico
  Set rsOperacaoSaida = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
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
  '15/12/2006 - Anderson - Altera��o para o cadastro de CFOP por servico
  rsOperacaoSaida.Close
  rsServicoCFOP.Close
  
  Set rsServicos = Nothing
  Set rsParametros = Nothing
  '15/12/2006 - Anderson - Altera��o para o cadastro de CFOP por servico
  Set rsOperacaoSaida = Nothing
  Set rsServicoCFOP = Nothing

  On Error GoTo 0
End Sub

Private Sub O_C�digo_Click()
  Set rsServicos = db.OpenRecordset("SELECT * FROM Servi�os WHERE C�digo <> 0 ORDER BY C�digo", dbOpenDynaset)
End Sub

Private Sub O_Nome_Click()
  Set rsServicos = db.OpenRecordset("SELECT * FROM Servi�os WHERE C�digo <> 0 ORDER BY Descri��o", dbOpenDynaset)
End Sub

Private Sub Pre�o_GotFocus()
  Pre�o.SelStart = 0
  Pre�o.SelLength = Pre�o.MaxLength
End Sub

'15/12/2006 - Anderson
'Altera��o para o cadastro de CFOP por Servico
Private Sub DropDownCFOP_Click()

  GradeCFOP.Columns(1).Text = DropDownCFOP.Columns(1).Text
  GradeCFOP.Columns(2).Text = DropDownCFOP.Columns(0).Text

End Sub

'15/12/2006 - Anderson
'Altera��o para o cadastro de CFOP por servi�o
Private Sub GradeCFOP_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
 Dim Aux As Variant
 Dim Resp As Integer
  
  Call StatusMsg("")
  Aux = GradeCFOP.Columns(ColIndex).Text
  If ColIndex = 1 Then  'C�digo
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
    
    rsOperacaoSaida.Index = "C�digo"
    rsOperacaoSaida.Seek "=", Val(Aux)
    If rsOperacaoSaida.NoMatch Then
       MsgBox ("C�digo de opera��o inv�lido.")
       Cancel = True
       Exit Sub
    End If
    
    Rem Verifica se esta opera��o j� est� neste servico
    rsServicoCFOP.Index = "PrimaryKey"
    rsServicoCFOP.Seek "=", C�digo.Text, Val(Aux)
    If Not rsServicoCFOP.NoMatch Then
      MsgBox "Esta opera��o j� est� cadastrada neste servi�o.", vbExclamation
      Cancel = True
      Exit Sub
    End If
    
  End If
  
  GradeCFOP.Columns(0).Text = C�digo.Text
  GradeCFOP.Columns(2).Text = rsOperacaoSaida("Nome")

End Sub

'15/12/2006 - Anderson
'Altera��o para o cadastro de CFOP por servi�o
Private Sub GradeCFOP_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)

  Dim Resposta As Integer
  
  Call StatusMsg("")
  Resposta = MsgBox("Deseja apagar a opera��o selecionada? ", 260, "Aten��o")
  If Resposta = 7 Then
    DisplayMsg "Opera��o n�o apagada."
    Cancel = True
  End If
  
  DispPromptMsg = False

End Sub

'15/12/2006 - Anderson
'Altera��o para o cadastro de CFOP por servi�o
Private Sub GradeCFOP_LostFocus()

 If GradeCFOP.RowChanged = True Then
   GradeCFOP.Update
 End If

End Sub


'-----------------------------------------------------
'15/12/2006 - Anderson
'Altera��o para o cadastro de CFOP por servi�o
Private Sub LoadOperacao()
  Dim C�d As String
  Dim sSql As String

  'Arruma Grade
  C�d = CInt("0" & C�digo)
  
    sSql = "SELECT CodServico, CodOperacao, [Opera��es Sa�da].Nome, CFOP FROM ServicoCFOP "
    sSql = sSql + " INNER JOIN [Opera��es Sa�da] ON ServicoCFOP.CodOperacao = [Opera��es Sa�da].C�digo  WHERE CodServico =" + C�d + " "
    
    GradeCFOP.Visible = False
    GradeCFOP.DataMode = 1
  
    Set Rec_CFOP = db.OpenRecordset(sSql, dbOpenDynaset)
  
    Set DataCFOP1.Recordset = Rec_CFOP
    
  If C�d = "0" Then

    GradeCFOP.Enabled = False
    'Exit Sub
  Else
    GradeCFOP.Enabled = True
  
  End If
  
  'GradeCFOP.Enabled = True
  
  GradeCFOP.DataMode = 0

  GradeCFOP.ReBind
 
  GradeCFOP.Columns(0).Visible = False 'Cod Prod
  GradeCFOP.Columns(1).Width = 1000 'Cod Opera��o
  GradeCFOP.Columns(1).Caption = "C�d. Opera��o"
  GradeCFOP.Columns(2).Locked = True 'Nome da Opera��o
  GradeCFOP.Columns(2).Width = 4000
  GradeCFOP.Columns(3).Width = 1000 'CFOP
  
  GradeCFOP.Visible = True
  GradeCFOP.Columns(1).DropDownHwnd = DropDownCFOP.hwnd
  '-----------------------------------------------------
  
'  Nome_Foto_LostFocus
  
  Call StatusMsg("")

End Sub

