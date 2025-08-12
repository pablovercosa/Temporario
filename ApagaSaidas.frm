VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmApagaSaidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apaga Sa�das"
   ClientHeight    =   3825
   ClientLeft      =   1005
   ClientTop       =   1515
   ClientWidth     =   7035
   Icon            =   "ApagaSaidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   7035
   Begin VB.Frame fraResumos 
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   1930
      Width           =   6735
      Begin VB.CheckBox chkApagar 
         Caption         =   "Apagar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Caption         =   "Marque o campo abaixo se Deseja Apagar as informa��es de Resumo Di�rio, Resumo Di�rio Financeiro e Resumo Clientes:"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Per�odo"
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   1140
      Width           =   6735
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3555
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   285
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2700
         TabIndex        =   11
         Top             =   345
         Width           =   810
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.CommandButton B_Apaga 
      Caption         =   "Apagar"
      Height          =   400
      Left            =   5520
      TabIndex        =   4
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Op_Sa�da"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Opera��o 
      Bindings        =   "ApagaSaidas.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1695
      TabIndex        =   0
      Top             =   720
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
      Columns(0).Width=   8440
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1693
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Opera��o 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2655
      TabIndex        =   8
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Opera��o :"
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1695
      TabIndex        =   6
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Filial :"
      Height          =   255
      Left            =   135
      TabIndex        =   5
      Top             =   270
      Width           =   495
   End
End
Attribute VB_Name = "frmApagaSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOp_Sa�da As Recordset
Dim rsSaidas As Recordset
Dim rsParametros As Recordset
Dim rsSaidas_Prod As Recordset
Dim rsMov_Cheques As Recordset
Dim rsMov_Parcelas As Recordset
Dim rsSaidas_Serv As Recordset

Private Sub B_Apaga_Click()
Dim Erro As Integer
Dim Seq As Long

  Call StatusMsg("")

  On Error GoTo Processa_Erro
  
  Rem Verifica Datas
  Erro = False
  If IsNull(Data_Ini.Text) Then Erro = True
  If Erro = False Then If Not IsDate(Data_Ini.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data inicial inv�lida, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Data_Fim.Text) Then Erro = True
  If Erro = False Then If Not IsDate(Data_Fim.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data final inv�lida, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser menor ou igual � data final."
    Data_Ini.SetFocus
    Exit Sub
  End If

  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  
  '25/10/2004 - Daniel
  'Adicionado rotina para apagar registros
  'nas tabelas de Resumo do Quick Store
  '
  'Tables: [Resumo Di�rio Financeiro], [Resumo Di�rio]
  '        [Resumo Clientes]
  '
  '28/12/2004 - Daniel
  'Implementado crit�rio se o usu�rio deseja apagar ou n�o os Resumos
  If chkApagar.Value = vbChecked Then
    Call StatusMsg("Apagando registros das tabelas de resumos...")
    Call ApagarRegistrosResumos
  End If
  '--------------------------------------------------------------
  
  Seq = 0
  rsSaidas.Index = "Sequ�ncia"
  rsSaidas_Prod.Index = "Sequ�ncia"
  rsMov_Cheques.Index = "Ordem"
  rsMov_Parcelas.Index = "Ordem"
Lp1:
  rsSaidas.Seek ">", gnCodFilial, Seq
  If rsSaidas.NoMatch Then GoTo Fim
  If rsSaidas("Filial") <> gnCodFilial Then GoTo Fim
  Seq = rsSaidas("Sequ�ncia")
  
  If Nome_Opera��o.Caption <> "" Then
   If rsSaidas("Opera��o") <> Val(Combo_Opera��o.Text) Then GoTo Lp1
  End If
  
  If CDate(rsSaidas("Data")) < CDate(Data_Ini.Text) Then GoTo Lp1
  If CDate(rsSaidas("Data")) > CDate(Data_Fim.Text) Then GoTo Lp1
  
  Call StatusMsg("Apagando movimenta��o " + str(Seq))
  DoEvents
    
  rsSaidas.Delete
  
  
Lp_Prod:
  rsSaidas_Prod.Seek ">", gnCodFilial, Seq, 0
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Serv
  If rsSaidas_Prod("Filial") <> gnCodFilial Then GoTo Ve_Serv
  If rsSaidas_Prod("Sequ�ncia") <> Seq Then GoTo Ve_Serv
  
  rsSaidas_Prod.Delete
  GoTo Lp_Prod
  
Ve_Serv:
  rsSaidas_Serv.Index = "Sequ�ncia"
  
Lp_Serv:
  rsSaidas_Serv.Seek ">", gnCodFilial, Seq, 0
  If rsSaidas_Serv.NoMatch Then GoTo Ve_Cheques
  If rsSaidas_Serv("Filial") <> gnCodFilial Then GoTo Ve_Cheques
  If rsSaidas_Serv("Sequ�ncia") <> Seq Then GoTo Ve_Cheques
  
  rsSaidas_Serv.Delete
  GoTo Lp_Serv
  
  
Ve_Cheques:
  rsMov_Cheques.Seek ">", gnCodFilial, Seq, 0
  If rsMov_Cheques.NoMatch Then GoTo Ve_Parcelas
  If rsMov_Cheques("Filial") <> gnCodFilial Then GoTo Ve_Parcelas
  If rsMov_Cheques("Sequ�ncia") <> Seq Then GoTo Ve_Parcelas

  rsMov_Cheques.Delete
  GoTo Ve_Cheques
  
Ve_Parcelas:
  rsMov_Parcelas.Seek ">", gnCodFilial, Seq, 0
  If rsMov_Parcelas.NoMatch Then GoTo Lp1
  If rsMov_Parcelas("Filial") <> gnCodFilial Then GoTo Lp1
  If rsMov_Parcelas("Sequ�ncia") <> Seq Then GoTo Lp1

  rsMov_Parcelas.Delete
  GoTo Ve_Parcelas
  
  
Fim:

  DisplayMsg "Movimenta��es apagadas."
  
  Exit Sub
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Apagar Sa�das")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Sub
    Case 3 'Encerrar
      End
  End Select
  
End Sub


Private Sub Combo_Opera��o_CloseUp()
 Combo_Opera��o.Text = Combo_Opera��o.Columns(1).Text
 Combo_Opera��o_LostFocus
 
End Sub

Private Sub Combo_Opera��o_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Opera��o_LostFocus()
  Call StatusMsg("")
  Nome_Opera��o.Caption = ""
  If IsNull(Combo_Opera��o.Text) Then Exit Sub
  If Combo_Opera��o.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Opera��o.Text) Then Exit Sub
  If Val(Combo_Opera��o.Text) < 1 Then Exit Sub
  If Val(Combo_Opera��o.Text) > 999 Then Exit Sub
  
  rsOp_Sa�da.Index = "C�digo"
  rsOp_Sa�da.Seek "=", Val(Combo_Opera��o.Text)
  If rsOp_Sa�da.NoMatch Then Exit Sub
  Nome_Opera��o.Caption = rsOp_Sa�da("Nome")
  
End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub


Private Sub Form_Load()

  Call CenterForm(Me)
  
  gsTitle = LoadResString(201)
  gsMsg = LoadResString(208)
  gnStyle = vbOKOnly + vbInformation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  Data1.DatabaseName = gsQuickDBFileName
  Set rsOp_Sa�da = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsSaidas = db.OpenRecordset("Sa�das")
  Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos")
  Set rsMov_Cheques = db.OpenRecordset("Movimento - Cheques")
  Set rsMov_Parcelas = db.OpenRecordset("Movimento - Parcelas")
  Set rsSaidas_Serv = db.OpenRecordset("Sa�das - Servi�os")
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub
  Nome_Filial.Caption = str(gnCodFilial) + " - " + rsParametros("Nome")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsOp_Sa�da.Close
  rsParametros.Close
  rsSaidas.Close
  rsSaidas_Prod.Close
  rsMov_Cheques.Close
  rsMov_Parcelas.Close
  rsSaidas_Serv.Close
  
  Set rsOp_Sa�da = Nothing
  Set rsParametros = Nothing
  Set rsSaidas = Nothing
  Set rsSaidas_Prod = Nothing
  Set rsMov_Cheques = Nothing
  Set rsMov_Parcelas = Nothing
  Set rsSaidas_Serv = Nothing
End Sub

Private Sub ApagarRegistrosResumos()
  '25/10/2004 - Daniel
  'Adicionado rotina para apagar registros
  'nas tabelas de Resumo do Quick Store
  '
  'Tables: [Resumo Di�rio Financeiro], [Resumo Di�rio]
  '        [Resumo Clientes]
  Dim strSQL As String

  strSQL = "DELETE * FROM [Resumo Di�rio Financeiro]"
  strSQL = strSQL & " WHERE Data >= #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#"

  db.Execute strSQL

  strSQL = ""
  strSQL = "DELETE * FROM [Resumo Di�rio]"
  strSQL = strSQL & " WHERE Data >= #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#"

  db.Execute strSQL

  strSQL = ""
  strSQL = "DELETE * FROM [Resumo Clientes]"
  strSQL = strSQL & " WHERE Dia >= #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Dia <= #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#"

  db.Execute strSQL

End Sub

