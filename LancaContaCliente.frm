VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmLancaContaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lançamentos/Manutenção de Conta de Clientes"
   ClientHeight    =   3960
   ClientLeft      =   2775
   ClientTop       =   1980
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LancaContaCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   8265
   Begin VB.TextBox Descrição 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1425
      MaxLength       =   70
      TabIndex        =   3
      Top             =   1470
      Width           =   6540
   End
   Begin MSMask.MaskEdBox Teste 
      Height          =   360
      Left            =   6690
      TabIndex        =   9
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2805
      Width           =   1260
      _ExtentX        =   2223
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
      Left            =   1710
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3285
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Data Data3 
      Appearance      =   0  'Flat
      Caption         =   "Data3"
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
      Left            =   3465
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cli_For"
      Top             =   3285
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.TextBox Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1425
      MaxLength       =   9
      TabIndex        =   6
      Top             =   2805
      Width           =   1260
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Left            =   2970
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   3465
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSMask.MaskEdBox Valor_Pago 
      Height          =   360
      Left            =   6690
      TabIndex        =   8
      Top             =   2340
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
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
   Begin MSMask.MaskEdBox Qtde 
      Height          =   360
      Left            =   1425
      TabIndex        =   4
      Top             =   1905
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   360
      Left            =   1425
      TabIndex        =   5
      Top             =   2355
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto 
      Bindings        =   "LancaContaCliente.frx":4E95A
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1425
      TabIndex        =   2
      Top             =   1065
      Width           =   1905
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
      Columns(0).Width=   8625
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3360
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "LancaContaCliente.frx":4E96E
      DataSource      =   "Data3"
      Height          =   360
      Left            =   1425
      TabIndex        =   1
      ToolTipText     =   "Escolha o Cliente com Conta na Filial"
      Top             =   585
      Width           =   1260
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
      Columns(0).Width=   8096
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1958
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2222
      _ExtentY        =   635
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "LancaContaCliente.frx":4E982
      DataSource      =   "Data1"
      Height          =   360
      Left            =   1425
      TabIndex        =   0
      Top             =   150
      Width           =   1260
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   6826
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
      _ExtentX        =   2222
      _ExtentY        =   635
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   360
      Left            =   6690
      TabIndex        =   7
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   1905
      Width           =   1260
      _ExtentX        =   2223
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
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   450
      Top             =   3330
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
      Bands           =   "LancaContaCliente.frx":4E996
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   225
      Left            =   240
      TabIndex        =   22
      Top             =   1530
      Width           =   960
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2745
      TabIndex        =   21
      Top             =   150
      Width           =   5220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2745
      TabIndex        =   19
      Top             =   585
      Width           =   5220
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data da Venda"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5235
      TabIndex        =   17
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seqüência"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Recebido: (=)"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4950
      TabIndex        =   13
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Recebimento"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4860
      TabIndex        =   12
      Top             =   2850
      Width           =   1590
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Nome_Produto 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3420
      TabIndex        =   10
      Top             =   1035
      Width           =   4545
   End
End
Attribute VB_Name = "frmLancaContaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsContas_Cliente As Recordset
Dim rsProdutos As Recordset
Dim Contador As Long

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
  
  Select Case ActiveBar1.Tools("miOpOrdem").CBListIndex
    
    Case -1, 0  '"Por Filial, Cliente, Data, Produto"
      If Len(Trim(Combo_Empresa.Text)) = 0 Then
        Combo_Empresa.Text = "0"
      End If
      If Len(Trim(Combo_Cliente.Text)) = 0 Then
        Combo_Cliente.Text = "0"
      End If
      If Len(Trim(Combo_Produto.Text)) = 0 Then
        Combo_Produto.Text = "0"
      End If
      If Not IsDate(Data.Text) Then
        Data.Text = Date - 3
      End If
      gsWhere = "WHERE Filial >= " & Combo_Empresa.Text & " AND Cliente >= " & Combo_Cliente.Text
      gsWhere = gsWhere & " AND Data >= #" & Format(Data.Text, "mm/dd/yyyy") & "# "
      gsWhere = gsWhere & " AND Produto >= '" & Combo_Produto.Text & "'"
    Case 1  '"Por Filial, Sequência"
      If Len(Trim(Combo_Empresa.Text)) = 0 Then
        Combo_Empresa.Text = "0"
      End If
      If Len(Trim(Sequência.Text)) = 0 Then
        Sequência.Text = "0"
      End If
      gsWhere = "WHERE Filial >= " & Combo_Empresa.Text & " AND Sequência >= " & Sequência.Text
  End Select
  
  Set rsContas_Cliente = db.OpenRecordset(gsSql & " " & gsWhere & " " & gsOrder, dbOpenDynaset)
  If Not rsContas_Cliente.EOF Then
    Call ShowRecord
  Else
    gsTitle = LoadResString(201)
    gsMsg = "Nenhum registro encontrado em função dos dados fornecidos."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  
End Sub

Sub ShowRecord()

  Combo_Empresa.Text = rsContas_Cliente("Filial")
  Combo_Empresa_LostFocus
  Combo_Cliente.Text = rsContas_Cliente("Cliente")
  Combo_Cliente_LostFocus
  Combo_Produto.Text = rsContas_Cliente("Produto")
  Combo_Produto_LostFocus
  
  Qtde.Text = rsContas_Cliente("Qtde") & ""
  Valor.Text = rsContas_Cliente("Valor") & ""
  Sequência.Text = rsContas_Cliente("Sequência")
  
  Descrição.Text = rsContas_Cliente("Descrição") & ""
  
  Data.Mask = ""
  Data.Text = ""
  Data.Mask = "##/##/####"
  If IsDate(rsContas_Cliente("Data")) Then Data.Text = Format(rsContas_Cliente("Data"), "dd/mm/yyyy")
  
  Valor_Pago.Text = rsContas_Cliente("Valor Pago")
  
  Teste.Mask = ""
  Teste.Text = ""
  Teste.Mask = "##/##/####"
  If IsDate(rsContas_Cliente("Data Pagamento")) Then
    Teste.Text = Format(rsContas_Cliente("Data Pagamento"), "dd/mm/yyyy")
  End If
  
  Num_Registro = rsContas_Cliente.Bookmark
  
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If
  
  Resposta = MsgBox(("Deseja realmente apagar esta conta ?"), 20, "ATENÇÃO!!")
  If Resposta = 6 Then
    rsContas_Cliente.Delete
    
    'LOG *****************
    Dim sSQL_Log As String
    sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
    sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & Sequência.Text & " Cli:" & Combo_Cliente.Text & " Prd:" & Combo_Produto.Text & " Vr:" & Valor.Text & " vPg:" & Valor_Pago.Text & " Dt:" & Data.Text, 80) & "', 'CNT_REC: exc cnt-cli')"
    db.Execute sSQL_Log, dbFailOnError
    'fim *******************
  
    Num_Registro = Null
    Call ClearScreen
  End If
  
End Sub

Private Sub UpdateRecord()
  Dim Aux_Filial As Integer
  Dim Aux_Sequência As Long
  Dim Erro As Integer
  Dim sTexto As String

  Call StatusMsg("")
  
  If Nome_Empresa.Caption = "" Then
    DisplayMsg "Filial incorreta, verifique."
    Combo_Empresa.SetFocus
    Exit Sub
  End If
  
  If Nome_Cliente.Caption = "" Then
    DisplayMsg "Cliente incorreto, verifique."
    Combo_Cliente.SetFocus
    Exit Sub
  End If
  
  If Nome_Produto.Caption = "" Then
    DisplayMsg "Produto incorreto, verifique."
    Combo_Produto.SetFocus
    Exit Sub
  End If

  Erro = False
  If IsNull(Qtde.Text) Then Erro = True
  If Erro = False Then If Qtde.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Qtde.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Quantidade incorreta, verifique."
    Qtde.SetFocus
    Exit Sub
  End If

  If Not IsDate(Data.Text) Then
    DisplayMsg "Data incorreta, verifique."
    Data.SetFocus
    Exit Sub
  End If
  
  If Teste.Text <> "  /  /    " And Not IsDate(Teste.Text) Then
    DisplayMsg "Data de pagamento incorreta, verifique."
    Teste.SetFocus
    Exit Sub
  End If
  
  If IsNull(Valor_Pago.Text) Then Valor_Pago.Text = 0
  If Not IsNumeric(Valor_Pago.Text) Then Valor_Pago.Text = 0
  If CDbl(Valor_Pago.Text) < 0 Then Valor_Pago.Text = 0
  
  If CDbl(Valor_Pago.Text) > CDbl(Valor.Text) Then
    DisplayMsg "Valor pago não pode ser superior ao valor da conta."
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  Call StatusMsg("Gravando ...")
  
  With rsContas_Cliente
  
    If IsNull(Num_Registro) Then
      .AddNew
      sTexto = "Lançamento Efetuado."
    Else
      .Edit
      sTexto = "Lançamento Alterado."
    End If
  
    Contador = .Fields("Contador")
    .Fields("Filial") = Combo_Empresa.Text
    .Fields("Cliente") = Combo_Cliente.Text
    .Fields("Produto") = Combo_Produto.Text
    .Fields("Qtde") = Qtde.Text
    .Fields("Valor") = CDbl(Valor.Text)
    .Fields("Descrição") = Descrição.Text
    
    If IsNull(Sequência.Text) Then Sequência.Text = 0
    If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
    If Sequência.Text = "" Then Sequência.Text = 0
    
    .Fields("Sequência") = Sequência.Text
    .Fields("Data") = Data.Text
    
    If IsNull(Valor_Pago.Text) Then Valor_Pago.Text = 0
    If Not IsNumeric(Valor_Pago.Text) Then Valor_Pago.Text = 0
    If Valor_Pago.Text = "" Then Valor_Pago.Text = 0
    
    .Fields("Valor Pago") = CDbl(Valor_Pago.Text)
    
    If Not IsDate(Teste.Text) Then
      .Fields("Data Pagamento") = Null
    Else
      .Fields("Data Pagamento") = Teste.Text
    End If
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
  
    Aux_Filial = .Fields("Filial")
    Aux_Sequência = .Fields("Sequência")
  
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
      
  End With
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Seq:" & Sequência.Text & " Cli:" & Combo_Cliente.Text & " Prd:" & Combo_Produto.Text & " Vr:" & Valor.Text & " vPg:" & Valor_Pago.Text & " Dt:" & Data.Text, 80) & "', 'CNT_REC: atu cnt-cli')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault

End Sub

Private Sub ClearScreen()
 
  Call StatusMsg("")
  
  Combo_Empresa.Text = ""
  Nome_Empresa.Caption = ""
  Combo_Cliente.Text = ""
  Nome_Cliente.Caption = ""
  Combo_Produto.Text = ""
  Nome_Produto.Caption = ""
  Descrição.Text = ""
  Qtde.Text = 0
  Valor.Text = 0
  Sequência.Text = ""
  Data.Mask = ""
  Data.Text = ""
  Data.Mask = "##/##/####"
  Valor_Pago.Text = 0
  
  Teste.Mask = ""
  Teste.Text = ""
  Teste.Mask = "##/##/####"
  
  If Not rsContas_Cliente.EOF Then
    On Error Resume Next
    rsContas_Cliente.MoveFirst
    rsContas_Cliente.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null

End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsContas_Cliente
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
  With rsContas_Cliente
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
  With rsContas_Cliente
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
  With rsContas_Cliente
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
        Case 0 '"Por Filial, Cliente, Data, Produto"
          gsOrder = "ORDER BY Filial, Cliente, Data, Produto"
        Case 1 '"Por Filial, Sequência"
          gsOrder = "ORDER BY Filial, Sequência"
      End Select
  End Select
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

Private Sub Combo_Produto_CloseUp()
  Combo_Produto.Text = Combo_Produto.Columns(1).Text
  Combo_Produto_LostFocus
End Sub

Private Sub Combo_Produto_LostFocus()

  Nome_Produto.Caption = ""
  If IsNull(Combo_Produto.Text) Then Exit Sub
  
  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Combo_Produto.Text
  If rsProdutos.NoMatch Then Exit Sub
  Nome_Produto.Caption = rsProdutos("Nome") & ""
  
End Sub

Private Sub Data_LostFocus()
  Data.Text = Ajusta_Data(Data.Text)
End Sub

Private Sub Data_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data.Text = frmCalendario.gsDateCalender(Data.Text)
  End Select
End Sub

Private Sub Descrição_GotFocus()
  If IsNull(Descrição.Text) Then Descrição.Text = ""
  If Descrição.Text <> "" Then Exit Sub
  Descrição.Text = Nome_Produto.Caption & ""
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
  Dim sSql As String

  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  ActiveBar1.Tools("miOpOrdem").CBList.Clear
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 0, "Por Filial, Cliente, Data, Produto"
  ActiveBar1.Tools("miOpOrdem").CBList.InsertItem 1, "Por Filial, Seqüência"
  ActiveBar1.Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  
  gsSql = "SELECT * FROM [Conta Cliente]"
  gsOrder = "ORDER BY Filial, Cliente, Data, Produto"
  Set rsContas_Cliente = db.OpenRecordset(gsSql & " " & gsOrder, dbOpenDynaset)
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)

  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  
  sSql = "SELECT Nome, Código, [Tem Conta] From Cli_For "
  sSql = sSql & " Where (Tipo = 'C' Or Tipo = 'R' Or Tipo = 'O') "
  sSql = sSql & " And (Inativo = False) And ([Tem Conta] = True) ORDER BY Nome"
  Set Data3.Recordset = db.OpenRecordset(sSql, dbOpenDynaset)
  Data3.Refresh
  
  Num_Registro = Null
  
  Call ActiveBarLoadToolTips(Me)
  
  Me.Show
  DoEvents
  
  Call ClearScreen
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsClientes.Close
  rsContas_Cliente.Close
  rsProdutos.Close
  Set rsParametros = Nothing
  Set rsClientes = Nothing
  Set rsContas_Cliente = Nothing
  Set rsProdutos = Nothing
End Sub

Private Sub Sequência_KeyPress(KeyAscii As Integer)
  KeyAscii = Verifica_Tecla_Integer(KeyAscii)
End Sub

Private Sub Teste_LostFocus()
  Teste.Text = Ajusta_Data(Teste.Text)
End Sub

Private Sub Teste_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Teste.Text = frmCalendario.gsDateCalender(Teste.Text)
  End Select
End Sub
