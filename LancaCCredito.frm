VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLancaCCredito 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lançamentos/Manutenção de Cartão de Crédito"
   ClientHeight    =   4140
   ClientLeft      =   2400
   ClientTop       =   1785
   ClientWidth     =   8160
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
   HelpContextID   =   1580
   Icon            =   "LancaCCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   8160
   Begin VB.CheckBox Recebido 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Recebido"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3315
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Cartão 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2040
      Width           =   3105
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Filial"
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
      Left            =   6030
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Cliente"
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
      Left            =   6030
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data3 
      Appearance      =   0  'Flat
      Caption         =   "Admin"
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
      Left            =   6030
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cartões"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Valor_Liq 
      Height          =   375
      Left            =   1470
      TabIndex        =   7
      Top             =   3270
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   375
      Left            =   1470
      TabIndex        =   6
      Top             =   2865
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   15066597
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
   Begin MSMask.MaskEdBox Bom_Para 
      Height          =   375
      Left            =   1470
      TabIndex        =   5
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2460
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo cboCartao 
      Bindings        =   "LancaCCredito.frx":4E95A
      DataSource      =   "Data3"
      Height          =   375
      Left            =   1470
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      DataFieldList   =   "Nome"
      MaxDropDownItems=   16
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
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   8705
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1905
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   2566
      Columns(2).Caption=   "Dias Pagar"
      Columns(2).Name =   "Dias Pagar"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Dias Pagar"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      Columns(3).Width=   2619
      Columns(3).Caption=   "Taxa"
      Columns(3).Name =   "Taxa"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Taxa"
      Columns(3).DataType=   4
      Columns(3).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
      DataFieldToDisplay=   "Nome"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "LancaCCredito.frx":4E96E
      DataSource      =   "Data2"
      Height          =   375
      Left            =   1470
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
      DataFieldList   =   "Nome"
      MaxDropDownItems=   16
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9366
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1693
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
      DataFieldToDisplay=   "Nome"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "LancaCCredito.frx":4E982
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1470
      TabIndex        =   0
      Top             =   120
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
      Columns(0).Width=   9816
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1667
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   -75
      Top             =   3675
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
      Bands           =   "LancaCCredito.frx":4E996
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Líquido"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   3330
      Width           =   1305
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2730
      TabIndex        =   17
      Top             =   120
      Width           =   5340
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seqüência"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   15
      Top             =   1125
      Width           =   1095
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2730
      TabIndex        =   14
      Top             =   1080
      Width           =   5340
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administradora"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   13
      Top             =   1620
      Width           =   1290
   End
   Begin VB.Label lblNomeCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2730
      TabIndex        =   12
      Top             =   1560
      Width           =   5340
   End
   Begin VB.Label label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cartão"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   2070
      Width           =   855
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Vencimento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   2490
      Width           =   1590
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   2895
      Width           =   975
   End
End
Attribute VB_Name = "frmLancaCCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim Ordem As Variant
Dim rsAdministradoras As Recordset
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsCR As Recordset

Private gnDiasPagar As Integer
Private gnTaxa As Single

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

  gsWhere = ""
  
  Call Combo_Empresa_LostFocus
  Call cboCartao_LostFocus
  
  Select Case ActiveBar1.Tools("miOpOrdem").CBListIndex
    
    Case -1, 0  '"Por Adminstradora, Cartão, Seqüência"
      If lblNomeCartao.Caption = "" Then
        cboCartao.Text = "0"
      End If
      If Cartão.Text = "" Then
        Cartão.Text = "0"
      End If
      gsWhere = "WHERE Tipo = 'O' AND Administradora >= " & cboCartao.Text & " AND Cartão >= '" & Cartão.Text & "'"
    Case 1  '"Por Filial, Data Vencimento"
      If Nome_Empresa.Caption = "" Then
        Combo_Empresa.Text = "0"
      End If
      If Not IsDate(Bom_Para.Text) Then
        Bom_Para.Text = Date - 3
      End If
      gsWhere = "WHERE Tipo = 'O' AND Filial >= " & Combo_Empresa.Text & " AND Vencimento >= #" & Format(Bom_Para.Text, "mm/dd/yyyy") & "#"
  End Select
  
  Set rsCR = db.OpenRecordset(gsSql & " " & gsWhere & " " & gsOrder, dbOpenDynaset)
  If Not rsCR.EOF Then
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
  
  If IsNull(Num_Registro) Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "Não existe registro para apagar !"
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Resposta = MsgBox(("Deseja realmente apagar este lançamento ?"), 20, "ATENÇÃO!!")
  If Resposta = 6 Then
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
      "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
      "frmLancaCCredito_DeleteRecord", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
    rsCR.Delete
    Num_Registro = Null
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  Dim sTexto As String
  Dim intRepeatUpdateLocked As Integer
  Dim blnInTransaction As Boolean
  
  
  On Error GoTo Trata_Erro:
  
  Call StatusMsg("")

  Rem Verifica Empresa
  If Nome_Empresa.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Empresa inválida, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Combo_Empresa.SetFocus
    Exit Sub
  End If
  
  If IsNull(Sequência.Text) Then Sequência.Text = 0
  If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
  If Val(Sequência.Text) < 0 Then Sequência.Text = 0
  
  If Nome_Cliente.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Cliente inválido, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Combo_Cliente.SetFocus
    Exit Sub
  End If
  
  If lblNomeCartao.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Administradora inválida, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    cboCartao.SetFocus
    Exit Sub
  End If
  
  If IsNull(Cartão.Text) Or Cartão.Text = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Numero do cartão inválido, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Cartão.SetFocus
    Exit Sub
  End If
  
  If IsNull(Bom_Para.Text) Or Bom_Para.Text = "" Or Not IsDate(Bom_Para.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = "Data incorreta, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Bom_Para.SetFocus
    Exit Sub
  End If
  
  If IsNull(Valor.Text) Then Valor.Text = 0
  If Valor.Text = "" Then Valor.Text = 0
  If Not IsNumeric(Valor.Text) Then Valor.Text = 0
  If CDbl(Valor.Text <= 0) Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor incorreto, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Valor.SetFocus
    Exit Sub
  End If
  
  If IsNull(Valor_Liq.Text) Then Valor_Liq.Text = 0
  If Valor_Liq.Text = "" Then Valor_Liq.Text = 0
  If Not IsNumeric(Valor_Liq.Text) Then Valor_Liq.Text = 0
  If CDbl(Valor_Liq.Text <= 0) Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor incorreto, verifique."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Valor_Liq.SetFocus
    Exit Sub
  End If
  
  ws.BeginTrans
  blnInTransaction = True
  
  With rsCR
    If IsNull(Num_Registro) Then
      .AddNew
      Num_Registro = !Contador
      sTexto = "Lançamento efetuado."
    Else
      .LockEdits = True
      .Edit
      sTexto = "Lançamento alterado."
    End If
    
    .Fields("Tipo") = "O"
    .Fields("Filial") = Combo_Empresa.Text
    .Fields("Sequência") = Sequência.Text
    .Fields("Cliente") = Combo_Cliente.Text
    .Fields("Administradora") = cboCartao.Text
    .Fields("Cartão") = Cartão.Text
    .Fields("Vencimento") = Bom_Para.Text
    .Fields("Valor Cartão") = CDbl(Valor.Text)
    .Fields("Valor") = CDbl(Valor_Liq.Text)
    .Fields("Processado") = Recebido.Value
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
    If Recebido.Value = True Then
      .Fields("Valor Recebido") = CDbl(Valor_Liq.Text)
    Else
      .Fields("Valor Recebido") = 0
    End If
    
    'LOG *****************
    Dim sSQL_Log As String
    sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
    sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & rsCR("Sequência") & " Venc:" & rsCR("Vencimento") & " VrAtu:" & rsCR("Valor"), 80) & "', 'ATU CARTOES')"
    db.Execute sSQL_Log, dbFailOnError
    'fim *******************
    
'''    '10/09/2007 - Anderson
'''    'Gera arquivo log do sistema
'''    If g_bolSystemLog Then
'''      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
'''      "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
'''      "frmLancaCCredito_UpdateRecord", _
'''      "Contas a Receber", g_strArquivoSystemLog
'''    End If
    
    .Update
    .Bookmark = .LastModified
    
    ws.CommitTrans
    blnInTransaction = False
    
  Call StatusMsg("")
    
  End With
  
  Exit Sub
  
Trata_Erro:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Manutenção - Contas a receber")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Exit Sub
        Case 3 'Encerrar
          End
      End Select
  End Select
End Sub

Public Sub ClearScreen()
  Call StatusMsg("")
  Combo_Empresa.Text = ""
  Nome_Empresa.Caption = ""
  Sequência.Text = ""
  Combo_Cliente.Text = ""
  Nome_Cliente.Caption = ""
  cboCartao.Text = ""
  lblNomeCartao.Caption = ""
  Bom_Para.Mask = ""
  Bom_Para.Text = ""
  Bom_Para.Mask = "##/##/####"
  Valor.Text = 0
  Valor_Liq.Text = 0
  Cartão.Text = ""
  Recebido.Value = False
  
  If Not rsCR.EOF Then
    On Error Resume Next
    rsCR.MoveFirst
    rsCR.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  Combo_Empresa.SetFocus
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsCR
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
  With rsCR
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
  With rsCR
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
  With rsCR
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
        Case -1, 0 '"Por Adminstradora, Cartão, Seqüência"
          gsOrder = "ORDER BY Administradora, Cartão, Sequência"
        Case 1 '"Por Filial, Data Vencimento"
          gsOrder = "ORDER BY Filial, Vencimento, Sequência"
      End Select
  End Select
End Sub

Private Sub Bom_Para_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Bom_Para.Text = frmCalendario.gsDateCalender(Bom_Para.Text)
  End Select
End Sub


Private Sub Bom_Para_LostFocus()
  Bom_Para.Text = Ajusta_Data(Bom_Para.Text)
End Sub

Private Sub cboCartao_CloseUp()
  Dim bm As Variant
  bm = cboCartao.GetBookmark(0)
  cboCartao.Text = cboCartao.Columns(1).CellText(bm)
  If Len(Trim(cboCartao.Text)) > 0 Then
    gnDiasPagar = cboCartao.Columns("Dias Pagar").CellText(bm)
    Bom_Para.Text = Date + gnDiasPagar
    gnTaxa = cboCartao.Columns("Taxa").CellText(bm)
  End If
  cboCartao_LostFocus
End Sub

Private Sub cboCartao_LostFocus()
  lblNomeCartao.Caption = ""
  If IsNull(cboCartao.Text) Then Exit Sub
  If Not IsNumeric(cboCartao.Text) Then Exit Sub
  If Val(cboCartao.Text) < 0 Or Val(cboCartao.Text) > 9999 Then Exit Sub

  rsAdministradoras.Index = "Código"
  rsAdministradoras.Seek "=", Val(cboCartao.Text)
  If rsAdministradoras.NoMatch Then Exit Sub
  lblNomeCartao.Caption = rsAdministradoras("Nome")
  gnDiasPagar = rsAdministradoras("Dias Pagar")
  gnTaxa = rsAdministradoras("Taxa")
  '12/11/2004 - Daniel
  'Solicitado pelo Suporte Infopar para mostrar o valor
  'real independente da taxa ter sido alterada ou não
  'Comentamos a linha da chamada da Private WriteLiquidValue
  'Call WriteLiquidValue
End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 0 Or Val(Combo_Cliente.Text) > 99999999 Then Exit Sub

  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  Nome_Cliente.Caption = rsClientes("Nome")

End Sub

Private Sub Combo_Empresa_CloseUp()
 Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
 Combo_Empresa_LostFocus
End Sub

Private Sub Combo_Empresa_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Empresa.Text) Then Exit Sub
  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
  If Val(Combo_Empresa.Text) < 0 Or Val(Combo_Empresa.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Empresa.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

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

  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Adminstradora, Cartão, Seqüência"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Filial, Data Vencimento"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  gsSql = "SELECT * FROM [Contas a Receber] "
  gsOrder = "ORDER BY Administradora, Cartão, Sequência"
  Set rsCR = db.OpenRecordset(gsSql & " WHERE Tipo = 'O' " & gsOrder, dbOpenDynaset)
  
  Set rsAdministradoras = db.OpenRecordset("Cartões", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  
  Me.Show
  DoEvents
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName

  Ordem = Null
  
  Call ActiveBarLoadToolTips(Me)
  
  Call ClearScreen
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub ShowRecord()
  Combo_Empresa.Text = rsCR("Filial")
  Sequência.Text = rsCR("Sequência")
  Combo_Cliente.Text = rsCR("Cliente")
  cboCartao.Text = rsCR("Administradora")
  Cartão.Text = rsCR("Cartão")
  Bom_Para.Text = Format(rsCR("Vencimento"), "dd/mm/yyyy")
  Valor.Text = rsCR("Valor Cartão")
  Valor_Liq.Text = rsCR("Valor")
  Recebido = -rsCR("Processado")
  Combo_Empresa_LostFocus
  Combo_Cliente_LostFocus
  cboCartao_LostFocus
  Num_Registro = rsCR("Contador")
  'Num_Registro = Null
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub Valor_LostFocus()
  Call WriteLiquidValue
End Sub

Private Sub WriteLiquidValue()
  Valor_Liq.Text = Round(CSng(gsHandleNull(Valor.Text)) - CSng(gsHandleNull(Valor.Text)) * gnTaxa / 100, 2)
End Sub
