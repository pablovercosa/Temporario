VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmVerificaDatas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificação de Datas dos Próximos Acertos"
   ClientHeight    =   5790
   ClientLeft      =   2445
   ClientTop       =   1635
   ClientWidth     =   7485
   Icon            =   "VerificaDatas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   7485
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      Height          =   795
      Left            =   195
      TabIndex        =   6
      Top             =   60
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3675
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para Calendário"
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
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   7
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   6015
      TabIndex        =   3
      Top             =   525
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6030
      Visible         =   0   'False
      Width           =   2745
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Bindings        =   "VerificaDatas.frx":058A
      Height          =   4605
      Left            =   195
      TabIndex        =   4
      Top             =   1020
      Width           =   7155
      _Version        =   196617
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   12621
      _ExtentY        =   8123
      _StockProps     =   79
      Caption         =   "Acertos"
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
   Begin VB.CommandButton B_Mostra 
      Caption         =   "Exibir"
      Height          =   400
      Left            =   6030
      TabIndex        =   2
      Top             =   75
      Width           =   1335
   End
   Begin VB.Label Tipo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   360
      Left            =   630
      TabIndex        =   5
      Top             =   5385
      Visible         =   0   'False
      Width           =   2325
   End
End
Attribute VB_Name = "frmVerificaDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmp_Entradas As Recordset
Dim rsEmp_Saídas As Recordset
Dim rsCliFor As Recordset

Sub Mostra_Entrada()

  Dim sSql As String
  Dim Rec_Contas As Recordset
  
  sSql = "SELECT [Data Cobrança] , Fornecedor, Cli_for.Nome FROM [Consignação Entrada]"
  sSql = sSql + " INNER JOIN [Cli_For] ON [Consignação Entrada].Fornecedor = Cli_For.Código"
  sSql = sSql + " WHERE Concluído = False" + " AND [Data Cobrança] >= #" + Data_Ini.Text + "#"
  sSql = sSql + " And [Data Cobrança] <= #" + Data_Fim.Text + "#"
  sSql = sSql + " ORDER BY [Data Cobrança]"
  
  Call StatusMsg("Aguarde, montando tabela.")
  DoEvents
  
  Set Rec_Contas = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade1.DataMode = 1
  
  Set Data3.Recordset = Rec_Contas

  Grade1.DataMode = ssDataModeBound
  Grade1.Columns(0).Width = 1200
  Grade1.Columns(1).Width = 1200
  Grade1.Columns(2).Width = 4000
  
  Call StatusMsg("")

End Sub

Sub Mostra_Saída()

  Dim sSql As String
  Dim Rec_Contas As Recordset
  
  sSql = "SELECT [Data Cobrança] , Cliente, Cli_for.Nome FROM [Consignação Saída]"
  sSql = sSql + " INNER JOIN [Cli_For] ON [Consignação Saída].Cliente = Cli_For.Código"
  sSql = sSql + " WHERE Concluído = False" + " AND [Data Cobrança] >= #" + Data_Ini.Text + "#"
  sSql = sSql + " And [Data Cobrança] <= #" + Data_Fim.Text + "#"
  sSql = sSql + " ORDER BY [Data Cobrança]"
  
  Call StatusMsg("Aguarde, montando tabela...")
  Set Rec_Contas = db.OpenRecordset(sSql, dbOpenDynaset)

  Grade1.DataMode = 1
  
  Set Data3.Recordset = Rec_Contas

  Grade1.DataMode = ssDataModeBound
  Grade1.Columns(0).Width = 1200
  Grade1.Columns(1).Width = 1200
  Grade1.Columns(2).Width = 4000
  
  Call StatusMsg("")

End Sub

Private Sub B_Imprime_Click()
  Grade1.PrintData ssPrintAllRows, True, True
End Sub

Private Sub B_Mostra_Click()

  Dim Erro As Boolean
  
  Call StatusMsg("")
  
  If Not IsDate(Data_Ini.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(53)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(53)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(54)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Data_Ini.SetFocus
    Exit Sub
  End If

  If Tipo.Caption = "ENTRADA" Then
    Mostra_Entrada
  ElseIf Tipo.Caption = "SAÍDA" Then
    Mostra_Saída
  End If
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
  Set rsEmp_Entradas = db.OpenRecordset("Consignação Entrada", , dbReadOnly)
  Set rsEmp_Saídas = db.OpenRecordset("Consignação Saída", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Data3.DatabaseName = gsQuickDBFileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsEmp_Entradas.Close
  rsEmp_Saídas.Close
  rsCliFor.Close
  Set rsEmp_Entradas = Nothing
  Set rsEmp_Saídas = Nothing
  Set rsCliFor = Nothing
End Sub
