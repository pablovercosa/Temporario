VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmVerificaOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificação de O.S."
   ClientHeight    =   6315
   ClientLeft      =   525
   ClientTop       =   825
   ClientWidth     =   10725
   Icon            =   "VerificaOS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   10725
   Begin VB.CommandButton B_Limpa 
      Caption         =   "&Limpar"
      Height          =   400
      Left            =   5265
      TabIndex        =   10
      Top             =   5730
      Width           =   1335
   End
   Begin VB.CommandButton B_Anterior 
      Caption         =   "&Anterior"
      Height          =   400
      Left            =   7470
      TabIndex        =   11
      Top             =   5730
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   7140
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.TextBox Referência 
      Height          =   315
      Left            =   6465
      MaxLength       =   10
      TabIndex        =   3
      Top             =   540
      Width           =   1485
   End
   Begin VB.CommandButton B_Próximo 
      Caption         =   "&Próximo"
      Height          =   400
      Left            =   9045
      TabIndex        =   12
      Top             =   5730
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4650
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7140
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7140
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   8010
      TabIndex        =   18
      Top             =   -15
      Width           =   2655
      Begin VB.OptionButton O_Sequência 
         Caption         =   "Seqüência"
         Height          =   225
         Left            =   1395
         TabIndex        =   7
         Top             =   525
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton O_Referência 
         Caption         =   "Ref. Interna"
         Height          =   225
         Left            =   1395
         TabIndex        =   6
         Top             =   270
         Width           =   1170
      End
      Begin VB.OptionButton O_Data 
         Caption         =   "Data Entrada"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   1275
      End
      Begin VB.OptionButton O_Cliente 
         Caption         =   "Cliente"
         Height          =   225
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   1170
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "VerificaOS.frx":058A
      DataSource      =   "Data3"
      Height          =   315
      Left            =   3945
      TabIndex        =   0
      Top             =   120
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
      Columns(0).Width=   9287
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2355
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
   Begin SSDataWidgets_B.SSDBGrid Grade2 
      Bindings        =   "VerificaOS.frx":059E
      Height          =   1800
      Left            =   105
      TabIndex        =   9
      Top             =   2700
      Width           =   10545
      _Version        =   196617
      AllowUpdate     =   0   'False
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
      BackColorOdd    =   8438015
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   18600
      _ExtentY        =   3175
      _StockProps     =   79
      Caption         =   "Serviços"
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
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Bindings        =   "VerificaOS.frx":05B2
      Height          =   1680
      Left            =   90
      TabIndex        =   8
      Top             =   930
      Width           =   10545
      _Version        =   196617
      AllowUpdate     =   0   'False
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
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   18600
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Produtos"
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
   Begin MSMask.MaskEdBox Data_Entrada 
      Height          =   315
      Left            =   3945
      TabIndex        =   2
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   540
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Sequência 
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      Top             =   555
      Width           =   1170
   End
   Begin VB.Label Técnico 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1590
      TabIndex        =   33
      Top             =   5640
      Width           =   3165
   End
   Begin VB.Label Label14 
      Caption         =   "Técnico :"
      Height          =   225
      Left            =   225
      TabIndex        =   32
      Top             =   5655
      Width           =   1170
   End
   Begin VB.Label Aprovado 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6945
      TabIndex        =   31
      Top             =   5250
      Width           =   3480
   End
   Begin VB.Label Label11 
      Caption         =   "Orçamento aprovado :"
      Height          =   285
      Left            =   5055
      TabIndex        =   30
      Top             =   5325
      Width           =   1695
   End
   Begin VB.Label Prometido 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1590
      TabIndex        =   29
      Top             =   5280
      Width           =   3165
   End
   Begin VB.Label Label9 
      Caption         =   "Prometido para :"
      Height          =   285
      Left            =   225
      TabIndex        =   28
      Top             =   5310
      Width           =   1170
   End
   Begin VB.Label Observações 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1590
      TabIndex        =   27
      Top             =   4920
      Width           =   8835
   End
   Begin VB.Label Label12 
      Caption         =   "Observações :"
      Height          =   225
      Left            =   225
      TabIndex        =   26
      Top             =   4950
      Width           =   1170
   End
   Begin VB.Label Total_Nota 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   9150
      TabIndex        =   25
      Top             =   4575
      Width           =   1275
   End
   Begin VB.Label Total_Serviços 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5055
      TabIndex        =   24
      Top             =   4575
      Width           =   1275
   End
   Begin VB.Label Total_Produtos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1590
      TabIndex        =   23
      Top             =   4575
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "Total Nota :"
      Height          =   285
      Left            =   7785
      TabIndex        =   22
      Top             =   4575
      Width           =   960
   End
   Begin VB.Label Label7 
      Caption         =   "Total Serviços :"
      Height          =   285
      Left            =   3765
      TabIndex        =   21
      Top             =   4620
      Width           =   1170
   End
   Begin VB.Label Label6 
      Caption         =   "Total Produtos :"
      Height          =   285
      Left            =   225
      TabIndex        =   20
      Top             =   4605
      Width           =   1275
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4890
      TabIndex        =   19
      Top             =   120
      Width           =   3060
   End
   Begin VB.Label Label5 
      Caption         =   "Cliente :"
      Height          =   285
      Left            =   3225
      TabIndex        =   17
      Top             =   165
      Width           =   630
   End
   Begin VB.Label Label4 
      Caption         =   "Ref. Interna :"
      Height          =   225
      Left            =   5430
      TabIndex        =   16
      Top             =   615
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Entrada :"
      Height          =   225
      Left            =   3210
      TabIndex        =   15
      Top             =   615
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Seqüência :"
      Height          =   285
      Left            =   105
      TabIndex        =   14
      Top             =   615
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Verificação de O.S."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   0
      TabIndex        =   13
      Top             =   75
      Width           =   3015
   End
End
Attribute VB_Name = "frmVerificaOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSaidas As Recordset
Dim rsClientes As Recordset
Dim rsFuncionarios As Recordset

Sub ShowRecord()

  Dim sSql As String
  Dim Rec_Produtos As Recordset
  Dim Rec_Serviços As Recordset
  
  sSql = "SELECT [Saídas - Produtos].Código, Produtos.Nome, Qtde, Preço, [Preço Final] From [Saídas - Produtos]"
  sSql = sSql + " Inner Join Produtos ON [Saídas - Produtos].[Código sem Grade] = Produtos.Código"
  sSql = sSql + " Where Filial = " + str(gnCodFilial)
  sSql = sSql + " AND Sequência = " + str(rsSaidas("Sequência"))
    
  Set Rec_Produtos = db.OpenRecordset(sSql, dbOpenDynaset)
  
  Grade1.DataMode = 1
  Set Data1.Recordset = Rec_Produtos
  Grade1.Visible = False
  Grade1.DataMode = 0
  Grade1.ReBind
  Grade1.Columns(0).Width = 2000
  Grade1.Columns(1).Width = 5000
  Grade1.Columns(2).Width = 1000
  Grade1.Columns(3).Width = 1000
  Grade1.Columns(3).NumberFormat = "###,###,##0.00"
  Grade1.Columns(4).Width = 1000
  Grade1.Columns(4).NumberFormat = "###,###,##0.00"
  
  Grade1.Visible = True
  

  sSql = "SELECT Código, Descrição, Tempo, Preço, Completo From [Saídas - Serviços]"
  sSql = sSql + " Where Filial = " + str(gnCodFilial)
  sSql = sSql + " AND Sequência = " + str(rsSaidas("Sequência"))
    
  Set Rec_Serviços = db.OpenRecordset(sSql, dbOpenDynaset)
  
  Grade2.DataMode = 1
  Set Data2.Recordset = Rec_Serviços
  Grade2.Visible = False
  Grade2.DataMode = 0
  Grade2.ReBind
  Grade2.Columns(0).Width = 2000
  Grade2.Columns(1).Width = 5000
  Grade2.Columns(2).Width = 1000
  Grade2.Columns(3).Width = 1000
  Grade2.Columns(3).NumberFormat = "###,###,##0.00"
  Grade2.Columns(4).Width = 1000
  Grade2.Columns(4).Style = ssStyleCheckBox
  
  
  Grade2.Visible = True


  Total_Serviços.Caption = Format(rsSaidas("Serviços"), "###,###,##0.00")
  Total_Produtos.Caption = Format(rsSaidas("Produtos"), "###,###,##0.00")
  Total_Nota.Caption = Format(rsSaidas("Total"), "###,###,##0.00")
  Observações.Caption = rsSaidas("Observações") & ""
  Prometido.Caption = rsSaidas("Prometido Para") & ""
  Aprovado.Caption = rsSaidas("Orçamento Aprovado") & ""

  Combo_Cliente.Text = rsSaidas("Cliente")
  Combo_Cliente_LostFocus
  
  Referência.Text = rsSaidas("Referência") & ""
  
  Data_Entrada.Text = Format(rsSaidas("Data"), "dd/mm/yyyy")
  Sequência.Text = rsSaidas("Sequência")

  Técnico.Caption = rsSaidas("Técnico") & ""
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", rsSaidas("Técnico")
  If Not rsFuncionarios.NoMatch Then
    Técnico.Caption = Técnico.Caption + " - " + rsFuncionarios("Apelido")
  End If
  
  



End Sub

Private Sub B_Anterior_Click()
Dim Aux_Seq As Variant
Dim Aux_Ref As Variant
Dim Aux_Data As Variant
Dim Aux_Cliente As Variant

Dim Erro As Boolean

Aux_Seq = Sequência.Text
If IsNull(Aux_Seq) Then Aux_Seq = 0
If Not IsNumeric(Aux_Seq) Then Aux_Seq = 0

Aux_Cliente = Combo_Cliente.Text
If IsNull(Aux_Cliente) Then Aux_Cliente = 0
If Not IsNumeric(Aux_Cliente) Then Aux_Cliente = 0

Aux_Ref = Referência.Text
If IsNull(Aux_Ref) Then Aux_Ref = ""

Aux_Data = Data_Entrada.Text
If IsNull(Aux_Data) Then Aux_Data = CDate("01/01/1980")
If Not IsDate(Aux_Data) Then Aux_Data = CDate("01/01/1980")




Erro = False

  If O_Sequência.Value = True Then
    If Aux_Seq = 0 Then Aux_Seq = 9999999
    rsSaidas.Index = "Sequência"
    rsSaidas.Seek "<", gnCodFilial, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  If O_Referência.Value = True Then
    rsSaidas.Index = "Referência"
    rsSaidas.Seek "<", gnCodFilial, Aux_Ref, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  If O_Data.Value = True Then
    rsSaidas.Index = "Data"
    rsSaidas.Seek "<", gnCodFilial, Aux_Data, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  If O_Cliente.Value = True Then
    rsSaidas.Index = "Cliente"
    rsSaidas.Seek "<", gnCodFilial, Aux_Cliente, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  
End Sub

Private Sub B_Limpa_Click()

 Sequência.Text = ""
 Combo_Cliente.Text = ""
 Combo_Cliente_LostFocus
 Referência.Text = ""
 
 Total_Serviços.Caption = ""
 Total_Produtos.Caption = ""
 Total_Nota.Caption = ""
 
 Observações.Caption = ""
 Prometido.Caption = ""
 Aprovado.Caption = ""
 Técnico.Caption = ""
 
 Grade1.Visible = False
 Grade2.Visible = False
 
 Data_Entrada.Mask = ""
 Data_Entrada.Text = ""
 Data_Entrada.Mask = "##/##/####"
 


End Sub

Private Sub B_Próximo_Click()
Dim Aux_Seq As Variant
Dim Aux_Ref As Variant
Dim Aux_Data As Variant
Dim Aux_Cliente As Variant

Dim Erro As Boolean

Aux_Seq = Sequência.Text
If IsNull(Aux_Seq) Then Aux_Seq = 0
If Not IsNumeric(Aux_Seq) Then Aux_Seq = 0

Aux_Cliente = Combo_Cliente.Text
If IsNull(Aux_Cliente) Then Aux_Cliente = 0
If Not IsNumeric(Aux_Cliente) Then Aux_Cliente = 0

Aux_Ref = Referência.Text
If IsNull(Aux_Ref) Then Aux_Ref = ""

Aux_Data = Data_Entrada.Text
If IsNull(Aux_Data) Then Aux_Data = CDate("01/01/1980")
If Not IsDate(Aux_Data) Then Aux_Data = CDate("01/01/1980")




Erro = False

  If O_Sequência.Value = True Then
    rsSaidas.Index = "Sequência"
    rsSaidas.Seek ">", gnCodFilial, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  If O_Referência.Value = True Then
    rsSaidas.Index = "Referência"
    rsSaidas.Seek ">", gnCodFilial, Aux_Ref, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  If O_Data.Value = True Then
    rsSaidas.Index = "Data"
    rsSaidas.Seek ">", gnCodFilial, Aux_Data, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  If O_Cliente.Value = True Then
    rsSaidas.Index = "Cliente"
    rsSaidas.Seek ">", gnCodFilial, Aux_Cliente, Aux_Seq
    If rsSaidas.NoMatch Then Erro = True
    If Erro = False Then If rsSaidas("Filial") <> gnCodFilial Then Erro = True
    If Erro = True Then
      Beep
      Exit Sub
    End If
    
    ShowRecord
  End If
  
  
End Sub

Private Sub Combo_Cliente_CloseUp()

 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus

End Sub

Private Sub Combo_Cliente_LostFocus()

 Nome_Cliente.Caption = ""
 If IsNull(Combo_Cliente.Text) Then Exit Sub
 If Combo_Cliente.Text = "" Then Exit Sub
 If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
 If Val(Combo_Cliente.Text) < 1 Then Exit Sub

 rsClientes.Index = "Código"
 rsClientes.Seek "=", Val(Combo_Cliente.Text)
 If rsClientes.NoMatch Then Exit Sub
 
 Nome_Cliente.Caption = rsClientes("Nome") & ""
 
End Sub

Private Sub Data_Entrada_LostFocus()
  Data_Entrada.Text = Ajusta_Data(Data_Entrada.Text)
End Sub

Private Sub Data_Entrada_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Entrada.Text = frmCalendario.gsDateCalender(Data_Entrada.Text)
  End Select
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
 Set rsSaidas = db.OpenRecordset("Saídas", , dbReadOnly)
 Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
 Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
 
 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName
 Data3.DatabaseName = gsQuickDBFileName
End Sub

Private Sub Referência_LostFocus()
 If IsNull(Referência.Text) Then Exit Sub
 
 Referência.Text = UCase(Referência.Text)
End Sub

Private Sub Sequência_KeyPress(KeyAscii As Integer)

 KeyAscii = Verifica_Tecla_Integer(KeyAscii)

End Sub

Private Sub Sequência_LostFocus()

  If IsNull(Sequência.Text) Then Exit Sub
  If Sequência.Text = "" Then Exit Sub
  If Not IsNumeric(Sequência.Text) Then Exit Sub
  
  rsSaidas.Index = "Sequência"
  rsSaidas.Seek "=", gnCodFilial, Sequência.Text
  If rsSaidas.NoMatch Then Exit Sub
  
  ShowRecord
  
End Sub
