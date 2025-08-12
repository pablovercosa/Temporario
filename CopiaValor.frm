VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosCopiaValor 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia Tabela de Preços Aplicando Valor"
   ClientHeight    =   4290
   ClientLeft      =   4005
   ClientTop       =   1065
   ClientWidth     =   5415
   ForeColor       =   &H80000008&
   HelpContextID   =   1660
   Icon            =   "CopiaValor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   5415
   Begin VB.CheckBox chkContaClientes 
      Caption         =   "&Refletir alteração também na Conta de Clientes"
      Enabled         =   0   'False
      Height          =   345
      Left            =   180
      TabIndex        =   15
      Top             =   3765
      Width           =   3690
   End
   Begin VB.Data datPrecos 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   4950
      Width           =   1875
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2220
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Sub_Classe"
      Top             =   5175
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Sobre 
      Caption         =   "&Sobrepõe preços existentes na tabela destino"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   1080
      Width           =   3645
   End
   Begin VB.CommandButton B_Calcula 
      Caption         =   "Copiar"
      Height          =   400
      Left            =   3990
      TabIndex        =   6
      Top             =   3750
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   5385
      Visible         =   0   'False
      Width           =   1725
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Sub 
      Bindings        =   "CopiaValor.frx":058A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Use 0 para todas as Subclasses"
      Top             =   2265
      Width           =   855
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
      Columns(0).Width=   8202
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
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "CopiaValor.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Use 0 para todas as Classes"
      Top             =   1680
      Width           =   855
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
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1852
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin MSMask.MaskEdBox Multiplicador 
      Height          =   315
      Left            =   1425
      TabIndex        =   5
      Top             =   3030
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##0.00"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabOrig 
      Bindings        =   "CopiaValor.frx":05B2
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   165
      Width           =   1935
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabDest 
      Bindings        =   "CopiaValor.frx":05CA
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   660
      Width           =   1935
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Label Nome_Sub 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Top             =   2250
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Subclasse :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Classe :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1725
      Width           =   735
   End
   Begin VB.Label Nome_Classe 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Valor :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   10
      Top             =   3045
      Width           =   1095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Use valores positivos para aumentar o preços e valores negativos para diminui-los."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2460
      TabIndex        =   9
      Top             =   2910
      Width           =   2775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Tabela DESTINO:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Tabela 
      Appearance      =   0  'Flat
      Caption         =   "Tabela ORIGINAL:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "frmPrecosCopiaValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsPreços As Recordset
Dim rsPreços2 As Recordset
Dim rsClasses As Recordset
Dim rsSub_Classes As Recordset
Dim rsProdutos As Recordset
Dim rsTabelas As Recordset
Private rsConta_Cli As Recordset

Private Sub cboTabDest_CloseUp()
  chkContaClientes.Enabled = True
  chkContaClientes.Value = vbUnchecked
  If cboTabDest.Text = cboTabOrig.Text And cboTabDest.Text <> "" Then
    chkContaClientes.Value = vbChecked
  Else
    chkContaClientes.Enabled = False
  End If
End Sub

Private Sub cboTabOrig_CloseUp()
  chkContaClientes.Enabled = True
  chkContaClientes.Value = vbUnchecked
  If cboTabDest.Text = cboTabOrig.Text And cboTabOrig.Text <> "" Then
    chkContaClientes.Value = vbChecked
  Else
    chkContaClientes.Enabled = False
  End If
End Sub

Private Sub chkContaClientes_Click()
  If chkContaClientes.Value = vbChecked Then
    If Not frmGerente.gbSenhaGerente Then
      chkContaClientes.Value = vbUnchecked
      Exit Sub
    End If
  End If
End Sub

Private Sub Combo_Classe_CloseUp()
 Combo_Classe.Text = Combo_Classe.Columns(1).Text
 Combo_Classe_LostFocus
End Sub

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado o suporte a transação com tratamento a erro
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Calcula_Click()
  Dim Produto As Variant
  Dim Preço As Variant
  Dim Copiados As Long
  Dim Aux As Integer
  Dim i As Integer
  Dim nTempCopiados As Long
  
  Dim Str_Arredonda As String
  Dim Novo_Preço As Double
  
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler
  
  Copiados = 0
  Produto = 0

  Call StatusMsg("")
  
  If IsNull(cboTabOrig.Text) Or cboTabOrig.Text = "" Then
    DisplayMsg "Tabela de Origem inválida !"
    cboTabOrig.SetFocus
    Exit Sub
  End If

  If IsNull(cboTabDest.Text) Or cboTabDest.Text = "" Then
    DisplayMsg "Tabela Destino inválida !"
    cboTabDest.SetFocus
    Exit Sub
  End If

  cboTabDest.Text = Trim(cboTabDest.Text)

  If IsNull(Multiplicador.Text) Then
    DisplayMsg "Digite o valor."
    Multiplicador.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(Multiplicador.Text) Then
    DisplayMsg "Digite o valor."
    Multiplicador.SetFocus
    Exit Sub
  End If
  
  If cboTabDest.Text = cboTabOrig.Text Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja efetuar as alterações na mesma tabela de preços?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
       DisplayMsg "Tabela não alterada."
       Exit Sub
    End If
    Sobre.Value = 1
  End If
    
  '28/02/2005 - Daniel
  '
  'Solicitação: Consultora Marineida
  '
  'Criado validação para não prosseguir caso não seja o gerente
  If Trim(cboTabOrig.Text) = Trim(cboTabDest.Text) Then
    If Not frmGerente.gbSenhaGerente Then Exit Sub
  End If
  '------------------------------------------------------------
  
  Screen.MousePointer = vbHourglass
  ws.BeginTrans
  blnOnTransaction = True
  
  If IsNull(Nome_Classe.Caption) Or Nome_Classe.Caption = "" Then Combo_Classe.Text = 0
  If IsNull(Nome_Sub.Caption) Or Nome_Sub.Caption = "" Then Combo_Sub.Text = 0


  Rem Começa a copiar as tabelas
  rsProdutos.Index = "Código"
  rsPreços.Index = "Tabela"
  rsPreços2.Index = "Tabela"
  
  

Lp1:
  If nTempCopiados <> Copiados Then
    nTempCopiados = Copiados
    Call StatusMsg("Foram copiados " & Copiados & " registros.")
  End If
  rsPreços.Seek ">", cboTabOrig, Produto
  If rsPreços.NoMatch Then
      Aux = 1
      GoTo Fim
  End If
  If rsPreços("Tabela") <> cboTabOrig.Text Then
      Aux = 2
      GoTo Fim
  End If

  Produto = rsPreços("Produto")

  rsProdutos.Seek "=", Produto
  If rsProdutos.NoMatch Then GoTo Lp1


  Rem Verifica se e' da classe desejada
  If Val(Combo_Classe.Text) <> 0 Then
     If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then GoTo Lp1
  End If

  Rem Verifica se e' da sub classe desejada
  If Val(Combo_Sub.Text) <> 0 Then
     If rsProdutos("Sub Classe") <> Val(Combo_Sub.Text) Then GoTo Lp1
  End If


  Novo_Preço = rsPreços("Preço") + CDbl(Multiplicador.Text)
  

  rsPreços2.Seek "=", cboTabDest.Text, rsPreços("Produto")
  If Not rsPreços2.NoMatch Then
    If Sobre.Value = 0 Then
      GoTo Lp1
    End If
  
    rsPreços2.Edit
    rsPreços2("Preço") = Format(Novo_Preço, "#############0.00")
    rsPreços2("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsPreços2.Update
    
    If chkContaClientes.Value = vbChecked Then
      Call UpdateContaClientes(cboTabDest.Text, rsPreços2("Produto").Value, Novo_Preço)
    End If
  
    'Atualiza o sincronismo para o produto WEB alterado
    Call WEB_SynchronizeProduct(rsPreços("Produto").Value)
    
    Copiados = Copiados + 1
    GoTo Lp1
  End If


  rsPreços2.AddNew
  
  rsPreços2("Tabela") = cboTabDest.Text
  rsPreços2("Produto") = rsPreços("Produto")
  rsPreços2("Preço") = Format(Novo_Preço, "############0.00")
  rsPreços2("Data Alteração") = Format(Date, "dd/mm/yyyy")
  
  rsPreços2.Update
  
  If chkContaClientes.Value = vbChecked Then
    Call UpdateContaClientes(cboTabDest.Text, rsPreços2("Produto").Value, Novo_Preço)
  End If
  
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(rsPreços("Produto").Value)

  Copiados = Copiados + 1

  GoTo Lp1

Fim:
 
  'Cria configuração da tabela
  Call CheckConfigTablePrice(cboTabDest.Text)
  
  ws.CommitTrans
  blnOnTransaction = False
  
  datPrecos.Refresh
  cboTabOrig.Refresh
  cboTabDest.Refresh
  
  cboTabDest.Text = ""
  
  Screen.MousePointer = vbDefault
  
  DisplayMsg "Final de processo. Copiados " & Copiados & " registros."

  Call StatusMsg("")

  Exit Sub

ErrHandler:
  Screen.MousePointer = vbDefault
  If blnOnTransaction Then ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Combo_Classe_LostFocus()
  Nome_Classe.Caption = ""
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub

  rsClasses.Index = "Código"
  rsClasses.Seek "=", Combo_Classe.Text
  If Not rsClasses.NoMatch Then
     Nome_Classe.Caption = rsClasses("Nome")
  Else
     Combo_Classe.Text = 0
  End If

End Sub

Private Sub Combo_Sub_CloseUp()
 Combo_Sub.Text = Combo_Sub.Columns(1).Text
 Combo_Sub_LostFocus

End Sub

Private Sub Combo_Sub_LostFocus()
  Nome_Sub.Caption = ""
  If IsNull(Combo_Sub.Text) Then Exit Sub
  If Not IsNumeric(Combo_Sub.Text) Then Exit Sub

  rsSub_Classes.Index = "Código"
  rsSub_Classes.Seek "=", Combo_Sub.Text
  If Not rsSub_Classes.NoMatch Then
     Nome_Sub.Caption = rsSub_Classes("Nome")
  Else
     Combo_Sub.Text = 0
  End If

End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsPreços2 = db.OpenRecordset("Preços")
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSub_Classes = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Preços")
  Set rsConta_Cli = db.OpenRecordset("SELECT * FROM [Conta Cliente]", dbOpenDynaset)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datPrecos.DatabaseName = gsQuickDBFileName

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsPreços.Close
  rsPreços2.Close
  rsClasses.Close
  rsSub_Classes.Close
  rsProdutos.Close
  rsTabelas.Close
  rsConta_Cli.Close
  Set rsPreços = Nothing
  Set rsPreços2 = Nothing
  Set rsClasses = Nothing
  Set rsSub_Classes = Nothing
  Set rsProdutos = Nothing
  Set rsTabelas = Nothing
  Set rsConta_Cli = Nothing
End Sub

Private Sub Multiplicador_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub cboTabDest_KeyPress(KeyAscii As Integer)
  KeyAscii = gnLimitKeyPress(cboTabDest, 15, KeyAscii)
  If KeyAscii <> 0 Then
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub cboTabDest_LostFocus()
  If IsNull(cboTabDest.Text) Then Exit Sub
  cboTabDest.Text = UCase$(cboTabDest.Text)
'  If cboTabDest.Text = cboTabOrig.Text And Len(cboTabOrig.Text) > 0 Then
'    DisplayMsg "Aviso: As alterações serão realizadas na mesma tabela e não existe um desfaz automático."
'  End If
End Sub
