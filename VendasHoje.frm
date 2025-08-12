VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVendasHoje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Vendas de Hoje"
   ClientHeight    =   4860
   ClientLeft      =   1635
   ClientTop       =   3615
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VendasHoje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4860
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPeriodo 
      Caption         =   "Período"
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmd_calendarioNFCeAbaCinza 
      Height          =   420
      Left            =   5580
      Picture         =   "VendasHoje.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   -15
      Width           =   495
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   15
      TabIndex        =   0
      Top             =   420
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdProcura 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   2805
   End
   Begin VB.ComboBox cboOrdem 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "VendasHoje.frx":4F23C
      Left            =   15
      List            =   "VendasHoje.frx":4F246
      TabIndex        =   3
      Text            =   "1 - Crescente"
      Top             =   15
      Width           =   1605
   End
   Begin VB.CheckBox chkSomenteNaoEfetivadas 
      Appearance      =   0  'Flat
      Caption         =   "Somente as não efetivadas"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1740
      TabIndex        =   2
      Top             =   60
      Width           =   2325
   End
   Begin MSMask.MaskEdBox msk_dataDiaNFCe 
      Height          =   285
      Left            =   4260
      TabIndex        =   6
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   45
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Clique no NÚMERO"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   4650
      Width           =   1515
   End
End
Attribute VB_Name = "frmVendasHoje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public bo_acaoSeleciona_e_fecha As Boolean

Private Sub chkPeriodo_Click()
  If chkPeriodo.Value = 0 Then
    msk_dataDiaNFCe.Mask = "##/##/####"
    msk_dataDiaNFCe.Format = "dd/mm/yyyy"
    msk_dataDiaNFCe.Text = Format(Data_Atual, "dd/MM/yyyy")
    cmd_calendarioNFCeAbaCinza.Visible = True
  ElseIf chkPeriodo.Value = 1 Then
    msk_dataDiaNFCe.Format = ""
    msk_dataDiaNFCe.Mask = "##"
    msk_dataDiaNFCe.Text = "30"
    cmd_calendarioNFCeAbaCinza.Visible = False
  End If
  
  Call ListaVendas
End Sub

Private Sub cmd_calendarioNFCeAbaCinza_Click()
    msk_dataDiaNFCe.Text = frmCalendario.gsDateCalender(msk_dataDiaNFCe.Text)
End Sub

Private Sub cmdProcura_Click()

  If msk_dataDiaNFCe.Text = "  /  /    " Then
      MsgBox "Informe o dia de pesquisa", vbInformation, "Atenção"
      msk_dataDiaNFCe.SetFocus
      Exit Sub
  ElseIf Trim(msk_dataDiaNFCe.Text) = "" Then
      MsgBox "Informe a quantidade de dias da pesquisa", vbInformation, "Atenção"
      msk_dataDiaNFCe.SetFocus
      Exit Sub
  End If

  ListaVendas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Dim strSetting As String
  
  msk_dataDiaNFCe.Text = Format(Data_Atual, "dd/MM/yyyy")
  
  '07/05/2003 - mpdea
  'Settings de preferências da tela
  cboOrdem.Text = GetSetting("QuickStore", "Acha Venda", "Ordem", cboOrdem.List(0))
  chkSomenteNaoEfetivadas.Value = GetSetting("QuickStore", "Acha Venda", "NaoEfetivada", vbChecked)
  
  Call ListaVendas
  
  If gbUsuarioAcessoApenasTelaVendaRapida = True Then
      chkSomenteNaoEfetivadas.Value = 1
      chkSomenteNaoEfetivadas.Enabled = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '07/05/2003 - mpdea
  'Settings de preferências da tela
  SaveSetting "QuickStore", "Acha Venda", "Ordem", cboOrdem.Text
  SaveSetting "QuickStore", "Acha Venda", "NaoEfetivada", chkSomenteNaoEfetivadas.Value
  
  gOrigemTelaSaidasChamadorDaTelaAcharVendaHoje = False
  
  '18/01/2006 - mpdea
  'Descarrega form
  Set frmVendasHoje = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ListItem)
  Dim nMoviment As Long
  Static nLastItem As Long
  
  nMoviment = CLng(Item.Text)
  If nMoviment <> nLastItem Then
    nLastItem = nMoviment
    If gbHasSaidasServicos(gnCodFilial, nMoviment) Then
      DisplayMsg "Esta movimentação contém serviço(s). Para ser visualizada utilize a tela de Saídas."
      Exit Sub
    End If
    Call StatusMsg("Aguarde...")
    
    If gOrigemTelaSaidasChamadorDaTelaAcharVendaHoje = True Then
        Call frmSaidas.SearchRecord_peloNumSeq(nMoviment)
    Else
        Call g_frmVendaRapida.Mostra_Mov(nMoviment)
    End If
    Call StatusMsg("")
  End If
  
  If bo_acaoSeleciona_e_fecha = True Then
      Unload Me
  End If

  
End Sub

Private Sub ListaVendas()
  ' 30/04/2003 - Maikel
  ' Substituido todo o código antigo de preenchimento da grid de vendas pelos poucos códigos abaixo
  
  Dim rsSaidas      As Recordset
  Dim rsParametros  As Recordset
  Dim rsCliFor      As Recordset
  
  Dim strSQL        As String
  Dim Nome_cli      As String
  Dim Recebimento   As String
  Dim itmX          As ListItem
  Dim clmX          As ColumnHeader
  
  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  
  Set clmX = ListView1.ColumnHeaders.Add(, , "Número", 550)
  Set clmX = ListView1.ColumnHeaders.Add(, , "Cliente", 4000)
  Set clmX = ListView1.ColumnHeaders.Add(, , "Ref. Interna", 2000)
  Set clmX = ListView1.ColumnHeaders.Add(, , "Valor Total", 800, lvwColumnRight)
  Set clmX = ListView1.ColumnHeaders.Add(, , "Recb.", 450)
  Set clmX = ListView1.ColumnHeaders.Add(, , "Comanda", 800)
  
  '18/01/2006 - mpdea
  'Alterado tipo de abertura do recordset (dbOpenSnapshot -> dbOpenDynaset, dbReadOnly)
  Set rsParametros = db.OpenRecordset("SELECT [VR Código Operação], VR_OcultaOrc FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset, dbReadOnly)
  
  strSQL = "SELECT Saídas.Sequência, Saídas.Cliente, Saídas.Referência, Saídas.Total, Saídas.Efetivada, Saídas.Recebimento, SaidasComandas.CodComanda" & _
           " FROM (Saídas LEFT JOIN SaidasComandas ON (Saídas.Sequência = SaidasComandas.CodSaida) AND (Saídas.Filial = SaidasComandas.Filial)" & _
           " ) INNER JOIN [Operações Saída] ON [Operações Saída].Código = Saídas.Operação " & _
           " WHERE (((Saídas.Filial)=" & gnCodFilial & ")"
  
  If chkPeriodo.Value = 0 Then
    strSQL = strSQL & " AND ((Saídas.Data)=#" & Format(msk_dataDiaNFCe.Text, "mm/dd/yyyy") & "#))"
  ElseIf chkPeriodo.Value = 1 Then
    Dim periodo As Date
    periodo = DateAdd("d", -1 * CInt(msk_dataDiaNFCe.Text), Date)
    strSQL = strSQL & " AND ((Saídas.Data)>=#" & Format(periodo, "mm/dd/yyyy") & "#))"
  End If
  
  If gOrigemTelaSaidasChamadorDaTelaAcharVendaHoje = False Then
    If Not rsParametros.EOF Then strSQL = strSQL & " AND Saídas.Operação = " & rsParametros("VR Código Operação")
  End If
  
  '07/05/2003 - mpdea
  'Corrigido check para não efetivadas
  If chkSomenteNaoEfetivadas.Value = vbChecked Then
    strSQL = strSQL & " AND NOT Saídas.Efetivada "
  End If
  
  If rsParametros("VR_OcultaOrc").Value Then
    strSQL = strSQL & " AND [Operações Saída].Tipo <> 'O' "
  End If
  
  '07/05/2003 - mpdea
  'Corrigido campo da ordenação
  strSQL = strSQL & " ORDER BY Saídas.Sequência "
  
  If GetCodigoCombos(cboOrdem.Text) = 2 Then strSQL = strSQL & " DESC "
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rsSaidas
    If Not (.BOF And .EOF) Then .MoveFirst
    
    Do While Not .EOF
      Set rsCliFor = db.OpenRecordset(" SELECT Nome FROM Cli_For " & _
                                      " WHERE Código = " & .Fields("Cliente"), dbOpenSnapshot)
      
      '07/05/2003 - mpdea
      'Corrigido referência das propriedades do recordset rsCliFor
      If (rsCliFor.BOF And rsCliFor.EOF) Then
        Nome_cli = "Cliente não encontrado"
      Else
        Nome_cli = rsCliFor.Fields("Nome") & ""
      End If
      
      Recebimento = IIf(.Fields("Recebimento"), "Sim", "Não")
      
      Set itmX = ListView1.ListItems.Add()
      itmX.Text = .Fields("Sequência")
      itmX.SubItems(1) = Nome_cli
      itmX.SubItems(2) = .Fields("Referência") & ""
      itmX.SubItems(3) = Format(.Fields("Total"), FORMAT_VALUE)
      itmX.SubItems(4) = Recebimento
      itmX.SubItems(5) = .Fields("CodComanda") & ""
      
      .MoveNext
    Loop
  End With
End Sub
