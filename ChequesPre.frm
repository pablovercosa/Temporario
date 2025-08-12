VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmChequesPre 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lançamentos/Manutenção de Cheques Pré-Datados"
   ClientHeight    =   4920
   ClientLeft      =   1320
   ClientTop       =   1245
   ClientWidth     =   7575
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
   Icon            =   "ChequesPre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4920
   ScaleWidth      =   7575
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   1350
      MaxLength       =   40
      TabIndex        =   5
      Top             =   2095
      Width           =   6030
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
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
      Left            =   5670
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   5325
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Depositado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Depositado"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3705
      Width           =   1320
   End
   Begin VB.CheckBox Devolvido 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Cheque devolvido, não apagar."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4065
      Width           =   2760
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
      Left            =   3780
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Banco"
      Top             =   5340
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1935
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   5340
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   45
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   5325
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Cheque 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   1335
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1724
      Width           =   1455
   End
   Begin VB.TextBox Sequência 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   1335
      MaxLength       =   9
      TabIndex        =   1
      Top             =   611
      Width           =   1455
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "ChequesPre.frx":4E95A
      DataSource      =   "Data4"
      Height          =   345
      Left            =   1335
      TabIndex        =   8
      Top             =   3210
      Width           =   1035
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
      Columns.Count   =   3
      Columns(0).Width=   6720
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2699
      Columns(1).Caption=   "Apelido"
      Columns(1).Name =   "Apelido"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Apelido"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1799
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   1826
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   350
      Left            =   1335
      TabIndex        =   7
      Top             =   2837
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   18
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
      Height          =   350
      Left            =   1335
      TabIndex        =   6
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   2466
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
      Bindings        =   "ChequesPre.frx":4E96E
      DataSource      =   "Data3"
      Height          =   345
      Left            =   1335
      TabIndex        =   3
      Top             =   1350
      Width           =   1035
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
      Columns(0).Width=   10001
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1826
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "ChequesPre.frx":4E982
      DataSource      =   "Data2"
      Height          =   345
      Left            =   1335
      TabIndex        =   2
      Top             =   975
      Width           =   1035
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
      Columns(0).Width=   9393
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1640
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1826
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "ChequesPre.frx":4E996
      DataSource      =   "Data1"
      Height          =   345
      Left            =   1335
      TabIndex        =   0
      Top             =   240
      Width           =   1035
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
      Columns(0).Width=   7594
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1773
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1826
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2130
      Width           =   855
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   3960
      Top             =   3930
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
      Bands           =   "ChequesPre.frx":4E9AA
   End
   Begin VB.Label Nome_Vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2385
      TabIndex        =   22
      Top             =   3210
      Width           =   5010
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3225
      Width           =   975
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bom para"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2505
      Width           =   855
   End
   Begin VB.Label label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Nome_Banco 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2385
      TabIndex        =   17
      Top             =   1350
      Width           =   5010
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1365
      Width           =   735
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2385
      TabIndex        =   15
      Top             =   990
      Width           =   5010
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seqüência"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2385
      TabIndex        =   12
      Top             =   240
      Width           =   5010
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   270
      Width           =   855
   End
End
Attribute VB_Name = "frmChequesPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sChequeOriginal As String
Dim sValorOriginal As String
Dim sBomParaOriginal As String

Dim Num_Registro As Variant
Dim Ordem As Variant
Dim rsBancos As Recordset
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsCR As Recordset
Dim rsFuncionarios As Recordset

Private gsSql As String
Private gsWhere As String
Private gsOrder As String

Private Sub DeleteRecord()
  Dim Resposta As Integer
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Não existe registro para apagar !"
    Exit Sub
  End If
  
  If Devolvido.Value = True Then
    DisplayMsg "Cheque devolvido não pode ser apagado."
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente apagar este cheque?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
      "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
      "frmChequesPre_DeleteRecord", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
    rsCR.Delete
    Call MovePrevious
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  Rem Verifica Empresa
  If Nome_Empresa.Caption = "" Then
     DisplayMsg "Empresa inválida, verifique."
     Combo_Empresa.SetFocus
     Exit Sub
  End If
  
  If IsNull(Sequência.Text) Then Sequência.Text = 0
  If Not IsNumeric(Sequência.Text) Then Sequência.Text = 0
  If Val(Sequência.Text) < 0 Then Sequência.Text = 0
  
  If Nome_Cliente.Caption = "" Then
     DisplayMsg "Cliente inválido, verifique."
     Combo_Cliente.SetFocus
     Exit Sub
  End If
  
  If Nome_Banco.Caption = "" Then
     DisplayMsg "Banco inválido, verifique."
     Combo_Banco.SetFocus
     Exit Sub
  End If
  
  If IsNull(Cheque.Text) Or Cheque.Text = "" Then
     DisplayMsg "Numero do cheque inválido, verifique."
     Cheque.SetFocus
     Exit Sub
  End If
  
  If IsNull(Bom_Para.Text) Or Bom_Para.Text = "" Or Not IsDate(Bom_Para.Text) Then
     DisplayMsg "Data incorreta, verifique."
     Bom_Para.SetFocus
     Exit Sub
  End If
  
  
  If IsNull(Valor.Text) Then Valor.Text = 0
  If Valor.Text = "" Then Valor.Text = 0
  If Not IsNumeric(Valor.Text) Then Valor.Text = 0
  If CDbl(Valor.Text <= 0) Then
     DisplayMsg "Valor incorreto, verifique."
     Valor.SetFocus
     Exit Sub
  End If
  
  If Nome_Vendedor.Caption = "" Then Combo_Vendedor.Text = 0
  
  
  Call StatusMsg("Gravando ...")
  
  With rsCR
    If IsNull(Num_Registro) Then
       .AddNew
      .Fields("Data Emissão") = Format(Date, "dd/mm/yyyy")
    Else
       .Edit
    End If
    
    .Fields("Tipo") = "C"
    
    .Fields("Filial") = Combo_Empresa.Text
    .Fields("Sequência") = Sequência.Text
    .Fields("Cliente") = Combo_Cliente.Text
    .Fields("Banco") = Combo_Banco.Text
    .Fields("Cheque") = Cheque.Text
    .Fields("Descrição") = txtDescricao.Text
    .Fields("Vencimento") = Bom_Para.Text
    .Fields("Valor") = CDbl(Valor.Text)
    .Fields("Devolvido") = Devolvido.Value
    .Fields("Vendedor") = Val(Combo_Vendedor.Text)
    .Fields("Processado") = Depositado.Value
    .Fields("Data Alteração") = Format(Date, "dd/mm/yyyy")
    .Fields("Valor Recebido") = 0
    If Depositado.Value = True And Devolvido.Value = False Then
      .Fields("Valor Recebido") = CDbl(Valor.Text)
    End If
    
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
      "Cli:" & rsCR("Cliente") & "- Seq:" & rsCR("Sequência") & "- NF:" & rsCR("Nota") & "- Venc:" & rsCR("Vencimento") & "- Valor:" & rsCR("Valor"), _
      "frmChequesPre_UpdateRecord", _
      "Contas a Receber", g_strArquivoSystemLog
    End If
      
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  End With
  
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " ChOr:" & sChequeOriginal & " ChAt:" & Cheque.Text & " VrOr:" & sValorOriginal & " VrAtu:" & Valor.Text & " DtBom:" & sBomParaOriginal & " DtBomAt:" & Bom_Para.Text, 80) & "', 'CNT_REC: atu cheq-pr')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  Call StatusMsg("")
  
  Exit Sub
  
Deu_Erro1:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar registro."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub

Private Sub ClearScreen()

  Call StatusMsg("")
  
  Combo_Empresa.Text = ""
  Nome_Empresa.Caption = ""
  Sequência.Text = ""
  Combo_Cliente.Text = ""
  Nome_Cliente.Caption = ""
  Combo_Banco.Text = ""
  Nome_Banco.Caption = ""
  txtDescricao.Text = ""
  Bom_Para.Mask = ""
  Bom_Para.Text = ""
  Bom_Para.Mask = "##/##/####"
  Valor.Text = 0
  Cheque.Text = ""
  Combo_Vendedor.Text = 0
  Combo_Vendedor_LostFocus
  Devolvido.Value = False
  Depositado.Value = False
  
  If Not rsCR.EOF And Not rsCR.BOF Then
    rsCR.MoveFirst
    If Not rsCR.BOF Then
      rsCR.MovePrevious
    End If
  End If
  
  Ordem = Null
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
        Case 0 '"Por Filial, Bom Para"
          gsOrder = "ORDER BY Filial, Vencimento, Cliente"
        
        Case 1 '"Por Banco, Cheque"
          gsOrder = "ORDER BY Banco, Cheque"
        
        Case 2 '"Por Cliente, Bom Para"
          gsOrder = "ORDER BY Cliente, Vencimento"
        
        '24/08/2004 - mpdea
        'Incluído pesquisa por cheque
        Case 3 'Por Cheque
          gsOrder = "ORDER BY Cheque"
      
      End Select
  End Select
End Sub

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
    
    Case -1, 0  '"Por Filial, Bom Para"
      If Not IsDate(Bom_Para.Text) Then
        Bom_Para.Text = Date - 3
      End If
      If Len(Trim(Combo_Empresa.Text)) = 0 Then
        Combo_Empresa.Text = "0"
      End If
      gsWhere = "WHERE Tipo = 'C' AND Filial >= " & Combo_Empresa.Text & " AND Vencimento >= #" & Format(Bom_Para.Text, "mm/dd/yyyy") & "#"
    
    Case 1  '"Por Banco, Cheque"
      If Len(Trim(Combo_Banco.Text)) = 0 Then
        Combo_Banco.Text = "0"
      End If
      gsWhere = "WHERE Tipo = 'C' AND Banco >= " & Combo_Banco.Text & " AND Cheque >= '" & Cheque.Text & "'"
    
    Case 2  '"Por Cliente, Bom Para"
      If Not IsDate(Bom_Para.Text) Then
        Bom_Para.Text = Date - 3
      End If
      If Len(Trim(Combo_Cliente.Text)) = 0 Then
        Combo_Cliente.Text = "0"
      End If
      gsWhere = "WHERE Tipo = 'C' AND Cliente >= " & Combo_Cliente.Text & " AND Vencimento >= #" & Format(Bom_Para.Text, "mm/dd/yyyy") & "#"
    
    '24/08/2004 - mpdea
    'Incluído pesquisa por cheque
    Case 3 'Por cheque
      gsWhere = "WHERE Tipo = 'C' AND Cheque >= '" & Cheque.Text & "'"
    
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

Private Sub Bom_Para_LostFocus()
  Bom_Para.Text = Ajusta_Data(Bom_Para.Text)
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
  Nome_Banco.Caption = rsBancos("Nome")

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

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_LostFocus()

  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) < 0 Or Val(Combo_Vendedor.Text) > 9999 Then Exit Sub

  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Vendedor.Text)
  If rsFuncionarios.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsFuncionarios("Apelido")

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
  
  With ActiveBar1.Tools("miOpOrdem")
    With .CBList
      .Clear
      .InsertItem 0, "Por Filial, Bom Para"
      .InsertItem 1, "Por Banco, Cheque"
      .InsertItem 2, "Por Cliente, Bom Para"
    
      '24/08/2004 - mpdea
      'Incluído pesquisa por cheque
      .InsertItem 3, "Por Cheque"
      
    End With
    .Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
  End With
  
  gsSql = "SELECT * FROM [Contas a Receber] "
  gsOrder = "ORDER BY Filial, Vencimento, Cliente"
  Set rsCR = db.OpenRecordset(gsSql & " WHERE Tipo = 'C' " & gsOrder, dbOpenDynaset)
  
  Set rsBancos = db.OpenRecordset("Bancos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  Data4.DatabaseName = gsQuickDBFileName

  Call ActiveBarLoadToolTips(Me)
  
  Me.Show
  DoEvents
  
  Call ClearScreen

  Screen.MousePointer = vbDefault
  
End Sub

Private Sub ShowRecord()
  Combo_Empresa.Text = rsCR("Filial")
  Sequência.Text = rsCR("Sequência")
  Combo_Cliente.Text = rsCR("Cliente")
  Combo_Banco.Text = rsCR("Banco")
  Combo_Vendedor.Text = rsCR("Vendedor")
  Cheque.Text = rsCR("Cheque") & ""
  txtDescricao.Text = rsCR("Descrição") & ""
  Bom_Para.Text = Format(rsCR("Vencimento"), "dd/mm/yyyy")
  Valor.Text = rsCR("Valor")
  Devolvido.Value = -rsCR("Devolvido")
  Depositado.Value = -rsCR("Processado")
  Combo_Empresa_LostFocus
  Combo_Cliente_LostFocus
  Combo_Banco_LostFocus
  Combo_Vendedor_LostFocus
  Num_Registro = rsCR.Bookmark

  'Guarda valor original para o log
  sChequeOriginal = Cheque.Text
  sValorOriginal = Valor.Text
  sBomParaOriginal = Bom_Para.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsCR.Close
  rsBancos.Close
  rsParametros.Close
  rsClientes.Close
  rsFuncionarios.Close
  
  Set rsCR = Nothing
  Set rsBancos = Nothing
  Set rsParametros = Nothing
  Set rsClientes = Nothing
  Set rsFuncionarios = Nothing
End Sub
