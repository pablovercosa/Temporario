VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEstados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de ICMS por Estado"
   ClientHeight    =   6525
   ClientLeft      =   3270
   ClientTop       =   1920
   ClientWidth     =   11760
   Icon            =   "Estados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11760
   Begin VB.Frame fraDiferimento 
      Caption         =   "Diferimento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtTotal 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Data datFilial 
         Caption         =   "datFilial"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtObsDiferimento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         MaxLength       =   70
         TabIndex        =   3
         Top             =   3120
         Width           =   4815
      End
      Begin VB.TextBox txtEstadoCorrente 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   2
         Top             =   2400
         Width           =   480
      End
      Begin VB.TextBox txtBase 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   3255
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "Estados.frx":058A
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   960
         Width           =   735
         DataFieldList   =   "Filial"
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
         Columns(0).Width=   3200
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Obs. Diferimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado Corrente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sobre o Valor ICM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   495
         TabIndex        =   12
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sobre o Total da Nota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   1020
         Width           =   300
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
         Left            =   4920
         Top             =   5280
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
         Bands           =   "Estados.frx":05A2
      End
   End
   Begin VB.Frame fraX 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Fechar"
         Height          =   400
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5280
         Width           =   1335
      End
      Begin SSDataWidgets_B.SSDBGrid Grade1 
         Bindings        =   "Estados.frx":2840
         Height          =   4440
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2475
         _Version        =   196617
         RecordSelectors =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1746
         Columns(0).Caption=   "Estado"
         Columns(0).Name =   "Estado"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Estado"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   2143
         Columns(1).Caption=   "ICMS"
         Columns(1).Name =   "ICM"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   2
         Columns(1).DataField=   "ICM"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   4366
         _ExtentY        =   7832
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   $"Estados.frx":2854
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   615
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Estados"
      Top             =   7230
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "frmEstados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro    As Variant
Dim rstDiferimento  As Recordset

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstFiliais As Recordset
  
  txtNomeFilial.Text = ""
  
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstFiliais = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
  
  With rstFiliais
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome") & ""
    End If
    
    If Not rstFiliais Is Nothing Then .Close
    Set rstFiliais = Nothing
  End With
  
End Sub

Private Sub cmdClose_Click()
  Unload Me
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
  
  Data1.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  '23/05/2006 - mpdea
  'A verificação da personalização para uso de Diferimento
  'está centralizada na função LoadCases_CheckSerialCaseMod
  'Somente se usa Diferimento é carregado o recordset
  '
  '07/06/2005 - Daniel
  'A partir da beta 6.52.0.48 deixamos a frame Diferimento habilitada somente para
  'a empresa Embalavi pois foi para a mesma que a implementação foi criada
  'If CheckSerialCaseMod("QS31306-629") Then
  If g_blnDiferimento Then
    Set rstDiferimento = db.OpenRecordset("SELECT * FROM Diferimento ORDER BY Filial ", dbOpenDynaset)
    Me.Width = 11850
  Else
    Me.Width = 5925
  End If
  '
  '07/06/2005 - Daniel
  'Centralizamos o form depois de dimensionar o comprimento do mesmo
  Call CenterForm(Me)
  
  Me.Show
  
  Call ActiveBarLoadToolTips(Me)
  Call ClearScreen
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not rstDiferimento Is Nothing Then rstDiferimento.Close
  Set rstDiferimento = Nothing
End Sub

Private Sub Grade1_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  Dim Aux As Variant
  Dim Erro As Boolean
  
  If ColIndex = 1 Then
    Aux = Grade1.Columns(1).Text
    Erro = False
    If IsNull(Aux) Then Erro = True
    If Erro = False Then If Aux = "" Then Erro = True
    If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
    
    If Erro = True Then
      DisplayMsg "Use o % do ICM ou -1 para usar o percentual do cadastro de produtos."
      Cancel = True
      Exit Sub
    End If
  End If
End Sub

Private Function ValidaCampos() As Boolean
  '17/05/2004 - Daniel
  'Função com a finalidade de validar os campos do Diferimento
  
  ValidaCampos = False
  
  If Len(txtNomeFilial.Text) = 0 Then
    MsgBox "Favor selecionar uma Filial válida.", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    ValidaCampos = False
    Exit Function
  End If

  If Not IsNumeric(txtTotal.Text) Then
    MsgBox "Favor mencionar um percentual válido para o Total da Nota.", vbExclamation, "Quick Store"
    txtTotal.SetFocus
    ValidaCampos = False
    Exit Function
  End If
  
  If Not IsNumeric(txtBase.Text) Then
    MsgBox "Favor mencionar um percentual válido para o Valor ICM.", vbExclamation, "Quick Store"
    txtBase.SetFocus
    ValidaCampos = False
    Exit Function
  End If

  If Len(txtEstadoCorrente.Text) = 0 Then
    MsgBox "Favor mencionar um Estado válido.", vbExclamation, "Quick Store"
    txtEstadoCorrente.SetFocus
    ValidaCampos = False
    Exit Function
  End If

  ValidaCampos = True

End Function

Private Sub txtEstadoCorrente_LostFocus()
  txtEstadoCorrente.Text = UCase(txtEstadoCorrente.Text)
End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  
  cboFilial.Text = ""
  txtNomeFilial.Text = ""
  txtTotal.Text = "0,00"
  txtBase.Text = "0,00"
  txtEstadoCorrente.Text = ""
  txtObsDiferimento.Text = ""
  
  cboFilial.SetFocus
  Num_Registro = Null
  
  If Not rstDiferimento Is Nothing Then
    If Not rstDiferimento.EOF Then
      On Error Resume Next
      rstDiferimento.MoveFirst
      rstDiferimento.MovePrevious
      On Error GoTo 0
    End If
  End If

End Sub

Private Sub MoveFirst()
  On Error Resume Next
  
  With rstDiferimento
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
  
  With rstDiferimento
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
  
  With rstDiferimento
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
  
  With rstDiferimento
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
  Dim strAuxStr    As String
  
  If IsNull(Num_Registro) Then
    Beep
    MsgBox "Não existe nenhum Diferimento para apagar.", vbExclamation, "Quick Store"
    Exit Sub
  End If

  strAuxStr = "Deseja realmente apagar este Diferimento ? "
  intResposta = MsgBox(strAuxStr, 20, "ATENÇÃO!!")
  If intResposta = 6 Then
    rstDiferimento.Delete
    Num_Registro = Null
    Call ClearScreen
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

Sub ShowRecord()
  With rstDiferimento
    cboFilial.Text = .Fields("Filial").Value
    cboFilial_LostFocus
    txtTotal.Text = Format((.Fields("Total").Value), "##,###,##0.00")
    txtBase.Text = Format((.Fields("Base").Value), "##,###,##0.00")
    txtEstadoCorrente.Text = .Fields("EstadoCorrente").Value & ""
    txtObsDiferimento.Text = .Fields("ObsDiferimento").Value & ""
    
    Num_Registro = .Bookmark
  End With

End Sub

Private Sub UpdateRecord()
  
  On Error GoTo Processa_Erro
  
  If Not ValidaCampos Then Exit Sub
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
   With rstDiferimento
     If IsNull(Num_Registro) Then
        .AddNew
        .Fields("Filial") = CByte(cboFilial.Text)
     Else
       .Edit
     End If
     .Fields("Total") = CDbl(Format((txtTotal.Text), "##,###,##0.00"))
     .Fields("Base") = CDbl(Format((txtBase.Text), "##,###,##0.00"))
     .Fields("EstadoCorrente") = txtEstadoCorrente.Text & ""
     .Fields("ObsDiferimento") = txtObsDiferimento.Text & ""
     
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
