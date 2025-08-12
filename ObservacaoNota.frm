VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmObsNota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Observa��es"
   ClientHeight    =   5685
   ClientLeft      =   2250
   ClientTop       =   2130
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ObservacaoNota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   10140
   Begin VB.Frame Frame3 
      Caption         =   "Transportadora "
      Height          =   765
      Left            =   30
      TabIndex        =   17
      Top             =   1980
      Width           =   10065
      Begin VB.ComboBox cbo_frete 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "ObservacaoNota.frx":4E95A
         Left            =   1290
         List            =   "ObservacaoNota.frx":4E970
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   248
         Width           =   3675
      End
      Begin SSDataWidgets_B.SSDBCombo cboTransp 
         Bindings        =   "ObservacaoNota.frx":4EA9D
         Height          =   360
         Left            =   5460
         TabIndex        =   8
         Top             =   240
         Width           =   4545
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
         BackColorOdd    =   8438015
         RowHeight       =   423
         Columns.Count   =   6
         Columns(0).Width=   6535
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Municipio"
         Columns(1).Name =   "Municipio"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Cidade"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   688
         Columns(2).Caption=   "UF"
         Columns(2).Name =   "UF"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "Estado"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3863
         Columns(3).Caption=   "CNPJ"
         Columns(3).Name =   "CNPJ"
         Columns(3).CaptionAlignment=   0
         Columns(3).DataField=   "CGC"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3519
         Columns(4).Caption=   "IE"
         Columns(4).Name =   "IE"
         Columns(4).CaptionAlignment=   0
         Columns(4).DataField=   "Inscri��o"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   4657
         Columns(5).Caption=   "Endere�o"
         Columns(5).Name =   "Ender"
         Columns(5).CaptionAlignment=   0
         Columns(5).DataField=   "Endere�o"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         _ExtentX        =   8017
         _ExtentY        =   635
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label1 
         Caption         =   "Frete pago por"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label10 
         Caption         =   "Nome"
         Height          =   225
         Left            =   4965
         TabIndex        =   18
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.Data datTransp 
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
      Height          =   390
      Left            =   8775
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Transportadoras"
      Top             =   6345
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4575
      Width           =   10065
   End
   Begin VB.Frame Frame4 
      Caption         =   "Informa��es Adicionais"
      Height          =   1185
      Left            =   30
      TabIndex        =   19
      Top             =   2775
      Width           =   10065
      Begin VB.TextBox UfrmPlaca 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6915
         MaxLength       =   2
         TabIndex        =   29
         Top             =   690
         Width           =   390
      End
      Begin VB.TextBox Placa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   8430
         MaxLength       =   8
         TabIndex        =   28
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox Qtde 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Esp�cie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   4695
         MaxLength       =   10
         TabIndex        =   10
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox Marca 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   8430
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Bruto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   12
         Top             =   690
         Width           =   1815
      End
      Begin VB.TextBox L�quido 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   4695
         MaxLength       =   10
         TabIndex        =   13
         Top             =   690
         Width           =   1515
      End
      Begin VB.Label Label12 
         Caption         =   "UF"
         Height          =   225
         Left            =   6615
         TabIndex        =   31
         Top             =   765
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Placa"
         Height          =   225
         Left            =   7950
         TabIndex        =   30
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label13 
         Caption         =   "Qtde"
         Height          =   225
         Left            =   165
         TabIndex        =   24
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label14 
         Caption         =   "Esp�cie"
         Height          =   225
         Left            =   4050
         TabIndex        =   23
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Marca"
         Height          =   225
         Left            =   7875
         TabIndex        =   22
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label16 
         Caption         =   "Peso Bruto"
         Height          =   225
         Left            =   165
         TabIndex        =   21
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label17 
         Caption         =   "Peso L�quido"
         Height          =   225
         Left            =   3630
         TabIndex        =   20
         Top             =   765
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Observa��es "
      Height          =   1950
      Left            =   30
      TabIndex        =   16
      Top             =   0
      Width           =   10065
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   9915
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   9405
         MaxLength       =   30
         TabIndex        =   3
         Top             =   645
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   2
         Top             =   645
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   9915
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1455
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   9405
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1050
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1050
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox Obs 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   9405
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1455
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5115
      Width           =   10065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4035
      Width           =   10065
   End
End
Attribute VB_Name = "frmObsNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTransportadoras As Recordset
Public gsCliente      As String
Public lngSequencia   As Long
Public bytTipoTabela  As Byte      ' 1 - Sa�das, 2 - Entradas
'27/04/2004 - Daniel
'Vars para C�lculo autom�tico dos Pesos Bruto e L�quido
Dim m_sngPesoLiquido As Single
Dim m_sngPesoBruto   As Single

Private Sub cboTransp_Click()
  Dim bm As Variant
  bm = cboTransp.GetBookmark(0)
  gsTransportadora = cboTransp.Columns("Nome").CellText(bm) & ""
  gsCNPJTransportadora = cboTransp.Columns("CNPJ").CellText(bm) & ""
  gsIETransportadora = cboTransp.Columns("IE").CellText(bm) & ""
  gsEnderTransportadora = cboTransp.Columns("Ender").CellText(bm) & ""
  gsMunicipioTransportadora = cboTransp.Columns("Municipio").CellText(bm) & ""
  gsUFTransportadora = cboTransp.Columns("UF").CellText(bm) & ""
End Sub

Private Sub cboTransp_KeyPress(KeyAscii As Integer)
  If Len(cboTransp.Text) >= rsTransportadoras("Nome").Size Then
    If KeyAscii <> vbKeyBack Then
      Beep
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub cmdClear_Click()
  Dim nI As Integer
  'For nI = 0 To 7
  '  Obs(nI) = ""
  'Next nI
  For nI = 0 To 1
    Obs(nI) = ""
  Next nI
  cboTransp.Text = ""
  gsTransportadora = ""
  gsCNPJTransportadora = ""
  gsIETransportadora = ""
  gsEnderTransportadora = ""
  gsMunicipioTransportadora = ""
  gsUFTransportadora = ""
  Placa.Text = ""
  UfrmPlaca.Text = ""
  Qtde.Text = ""
  Esp�cie.Text = ""
  Marca.Text = ""
  Bruto.Text = ""
  L�quido.Text = ""
  'O_Emitente.Value = False
End Sub

Private Sub Command1_Click()
  Dim nI As Integer
  
  'For nI = 0 To 7
  '  gsObsDoc(nI) = Obs(nI).Text & ""
  'Next nI
  For nI = 0 To 1
    Obs(nI).Text = Replace(Obs(nI).Text, vbCrLf, "")
    gsObsDoc(nI) = RemoveCaracteresEspeciaisParaNFE(Obs(nI).Text) & ""
  Next nI
  
  gsPlaca = Placa.Text & ""
  gsUfrmPlaca = UfrmPlaca.Text & ""

  If Trim(Qtde.Text) <> "" Then
    If IsNumeric(Qtde.Text) Then
        gsQtdeTrans = Qtde.Text & ""
    Else
        MsgBox "Entre com um valor num�rico para Quantidade", vbInformation, "Aten��o"
        Exit Sub
    End If
  End If
  
  gsMarcaTrans = RemoveCaracteresEspeciaisParaNFE(Marca.Text) & ""
  gsEspecieTrans = Esp�cie.Text & ""
  
  If Trim(Bruto.Text) <> "" Then
    If IsNumeric(Bruto.Text) Then
        gsPesoBruto = Bruto.Text & ""
    Else
        MsgBox "Entre com um valor num�rico para Peso Bruto", vbInformation, "Aten��o"
        Exit Sub
    End If
  End If
  
  If Trim(L�quido.Text) <> "" Then
    If IsNumeric(L�quido.Text) Then
        gsPesoLiquido = L�quido.Text & ""
    Else
        MsgBox "Entre com um valor num�rico para Peso L�quido", vbInformation, "Aten��o"
        Exit Sub
    End If
  End If
   
'  If O_Destinat�rio.Value = True Then gsFretePago = "2"
'  If O_Destinat�rio.Value = False Then gsFretePago = "1"
'  If O_SemFrete.Value = True Then gsFretePago = "9"

  gsFretePago = cbo_frete.ItemData(cbo_frete.ListIndex)
  
  gsRetornoDoc = "OK"
  Unload Me
End Sub

Private Sub Command2_Click()
  gsRetornoDoc = "N�O"
  Unload Me
End Sub

Private Sub Form_Activate()
  If Val("0" & gsCliente) = 0 Then Exit Sub
  
  With rsTransportadoras
    .FindFirst "C�digo = " & CInt(gsCliente)
    If rsTransportadoras.NoMatch Then
      Exit Sub
    End If
    cboTransp.Text = .Fields("Nome") & ""
    gsTransportadora = cboTransp.Text & ""
    gsCNPJTransportadora = .Fields("CGC") & ""
    gsIETransportadora = .Fields("Inscri��o") & ""
    gsEnderTransportadora = .Fields("Endere�o") & ""
    gsMunicipioTransportadora = .Fields("Cidade") & ""
    gsUFTransportadora = .Fields("Estado") & ""
  End With
End Sub

'13/08/2004 - mpdea
'Inclu�do tratamento de erro
Private Sub Form_Load()
  
  On Error GoTo ErrHandler
  
  Call CenterForm(Me)
    
  datTransp.DatabaseName = gsQuickDBFileName
  
  '13/09/2022 - Pablo
  'Inclu�do para n�o pr�-selecionar nenhuma transportadora no combobox
  Dim query As String
  query = "Select 0 as C�digo, '' as Nome, null as Endere�o, null as Cidade, null as Estado, null as CGC, null as Inscri��o, null as Telefone, null as Contatos, null as 'Data Altera��o' from transportadoras where c�digo = 1 union all select * from transportadoras;"
  'Set rsTransportadoras = db.OpenRecordset("SELECT * FROM Transportadoras ORDER BY Nome", dbOpenDynaset)
  Set rsTransportadoras = db.OpenRecordset(query, dbOpenDynaset)
  
  gsQtdeTrans = ""
  gsTransportadora = ""
  gsCNPJTransportadora = ""
  gsIETransportadora = ""
  gsEnderTransportadora = ""
  gsMunicipioTransportadora = ""
  gsUFTransportadora = ""
  
  '15/05/2007 - Anderson
  gsRetornoDoc = ""
  
  '20/07/2005 - Daniel
  'Case: Arte Telhas (Barro Queimado)
  'Solicitou para que a tela de observa��es da nota
  'viesse limpa ao ser exibida
  'If Not CheckSerialCaseMod("QS39215-718", "QS39240-574") Then Call GetSettings
  
  Call GetSettings
  
  '27/04/2004 - Daniel
  'C�lculo autom�tico dos Pesos
  Bruto.Text = m_sngPesoBruto
  L�quido.Text = m_sngPesoLiquido
  
  '27/09/2004 - mpdea
  'CASE: Embalavi
  'Calcula o total de volumagem (inteiro) por quantidade na movimenta��o
  If CheckSerialCaseMod("QS31306-629", "QS31571-867", "QS31572-951", "QS31581-959", "QS33016-722", "QS33458-286", "QS37456-162") Then
    Call CalculaQtde
  End If
  
  cbo_frete.ListIndex = 5
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao carregar tela: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim nI As Integer
  
  '15/05/2007 - Anderson
  'Se for clicado no bot�o cancelar, n�o armazenar as informa��es
  If gsRetornoDoc = "OK" Then
    For nI = 0 To 1
      Call SaveSetting("QuickStore", "ObsNota", "Obs" & CInt(nI), Obs(nI))
    Next nI
    'For nI = 0 To 7
    '  Call SaveSetting("QuickStore", "ObsNota", "Obs" & CInt(nI), Obs(nI))
    'Next nI
    SaveSetting _
      "QuickStore", "ObsNota", "NomeTransp", cboTransp.Text
    SaveSetting _
      "QuickStore", "ObsNota", "Placa", Placa.Text
    SaveSetting _
      "QuickStore", "ObsNota", "UF", UfrmPlaca.Text
    SaveSetting _
      "QuickStore", "ObsNota", "Especie", Esp�cie.Text
    SaveSetting _
      "QuickStore", "ObsNota", "Marca", Marca.Text
    'SaveSetting _
    '  "QuickStore", "ObsNota", "Emitente", O_Emitente.Value
  End If
End Sub

Private Sub GetSettings()
  Dim rstSaidas As Recordset
  Dim strSQL    As String
  Dim nI        As Integer
  Dim blnSaida  As Boolean '27/04/2004 - Daniel
  
  Dim rstParametros As Recordset '15/05/2007 - Anderson
  Dim bolManterInformacaoNotaFiscal As Boolean '15/05/2007 - Anderson
  
  '-------------------------------------------------------------------------------------------
  '15/05/2007 - Anderson
  'Indica se o Quick Store deve manter as observa��es impressas na �ltima Nota Fiscal
  
  bolManterInformacaoNotaFiscal = False
  
  Set rstParametros = db.OpenRecordset("SELECT * FROM [Par�metros Filial] WHERE Filial=" & gnCodFilial)
  
  If Not rstParametros.EOF Then
    bolManterInformacaoNotaFiscal = -rstParametros("MantemInformacaoUltimaNotaFiscal")
  End If
  
  rstParametros.Close
  
  Set rstParametros = Nothing
  '-------------------------------------------------------------------------------------------
  '15/05/2007 - Anderson
  'Altera��o realizada para incluir o campo Nota Impressa e verificar se existe algum dado registrado na nota fiscal para exibi��o
  'strSQL = " SELECT obs_Obs1, obs_Obs2, obs_Obs3, obs_Obs4, obs_Obs5, obs_Obs6, obs_Obs7, obs_Obs8, " & _
  '         " obs_Transportadora, obs_Placa, obs_Uf, obs_Qtde, obs_Especie, obs_Marca, obs_PesoLiquido, obs_PesoBruto, obs_FretePago "
  
  'strSQL = " SELECT [Nota Impressa], obs_Obs1, obs_Obs2, obs_Obs3, obs_Obs4, obs_Obs5, obs_Obs6, obs_Obs7, obs_Obs8, "
  strSQL = " SELECT [Nota Impressa], obs_infCpl1, obs_infCpl2, " & _
           " obs_Transportadora, obs_Placa, obs_Uf, obs_Qtde, obs_Especie, obs_Marca, obs_PesoLiquido, obs_PesoBruto, obs_FretePago "
  
  If bytTipoTabela = 1 Then
    strSQL = strSQL & " FROM Sa�das "
    blnSaida = True
  Else
    strSQL = strSQL & " FROM Entradas "
    blnSaida = False
  End If
  
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial & " AND Sequ�ncia = " & lngSequencia
  
  '27/04/2004 - Daniel
  If blnSaida Then
    m_sngPesoBruto = 0
    m_sngPesoLiquido = 0
    
    Call CalculaPesos(lngSequencia)
  End If
  
  Set rstSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstSaidas
    If Not (.BOF And .EOF) Then
    
      '15/05/2007 - Anderson
      'If IsNull(.Fields("obs_Obs1").Value) Then
      If .Fields("Nota Impressa").Value = 0 And bolManterInformacaoNotaFiscal Then
      
        'For nI = 0 To 7
        '  Obs(nI).Text = GetSetting("QuickStore", "ObsNota", "Obs" & CInt(nI), "")
        'Next nI
        For nI = 0 To 1
          Obs(nI).Text = GetSetting("QuickStore", "ObsNota", "Obs" & CInt(nI), "")
        Next nI
        cboTransp.Text = GetSetting("QuickStore", "ObsNota", "NomeTransp", "")
        Placa.Text = GetSetting("QuickStore", "ObsNota", "Placa", "")
        UfrmPlaca.Text = GetSetting("QuickStore", "ObsNota", "UF", "")
        Esp�cie.Text = GetSetting("QuickStore", "ObsNota", "Especie", "")
        Marca.Text = GetSetting("QuickStore", "ObsNota", "Marca", "")
  '      Bruto.Text = CStr(gsHandleNull(Format(gnPesoBruto, "#.000")))
  '      L�quido.Text = CStr(gsHandleNull(Format(gnPesoLiquido, "#.000")))
        'O_Emitente.Value = GetSetting("QuickStore", "ObsNota", "Emitente", "0")
      
      Else
    
        
        For nI = 0 To 1
          Obs(nI).Text = .Fields("obs_infCpl" & nI + 1).Value & ""
        Next nI
        'For nI = 0 To 7
        '  Obs(nI).Text = .Fields("obs_Obs" & nI + 1).Value & ""
        'Next nI
        cboTransp.Text = .Fields("obs_Transportadora") & ""
        Placa.Text = .Fields("obs_Placa") & ""
        UfrmPlaca.Text = .Fields("obs_Uf") & ""
        Qtde.Text = .Fields("obs_Qtde") & ""
        Esp�cie.Text = .Fields("obs_Especie") & ""
        Marca.Text = .Fields("obs_Marca") & ""
        'C�lculo autom�tico dos Pesos
        'Bruto.Text = .Fields("obs_PesoBruto") & ""
        'L�quido.Text = .Fields("obs_PesoLiquido") & ""
        
        'O_Emitente.Value = (.Fields("obs_FretePago") & "" = 1)
      
      End If
      
    End If
    
    .Close
    Set rstSaidas = Nothing
  End With
  
  
  '13/08/2004 - mpdea
  'Posiciona dados da Transportadora
  With rsTransportadoras
    .FindFirst "Nome = '" & cboTransp.Text & "'"
    If Not rsTransportadoras.NoMatch Then
      gsTransportadora = cboTransp.Text & ""
      gsCNPJTransportadora = .Fields("CGC") & ""
      gsIETransportadora = .Fields("Inscri��o") & ""
      gsEnderTransportadora = .Fields("Endere�o") & ""
      gsMunicipioTransportadora = .Fields("Cidade") & ""
      gsUFTransportadora = .Fields("Estado") & ""
    End If
  End With
  
End Sub

Private Function CalculaPesos(ByVal Seq As Long) As String
  '27/04/2004 - Daniel
  'Criado Rotina para C�lculo de Pesos Autom�ticos
  Dim rstSaidasProdutos As Recordset
  Dim strCodProduto     As String
  Dim sngQtde           As Single
  
  Set rstSaidasProdutos = db.OpenRecordset("SELECT C�digo, Qtde FROM [Sa�das - Produtos] WHERE Sequ�ncia =" & Seq & " AND Filial =" & gnCodFilial, dbOpenDynaset)
  
  With rstSaidasProdutos
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        strCodProduto = .Fields("C�digo").Value
        sngQtde = .Fields("Qtde").Value
        
        Call LocalizarProdutos(strCodProduto, sngQtde)
      
      .MoveNext
      Loop
    End If
    .Close
  End With

  Set rstSaidasProdutos = Nothing

End Function

Private Function LocalizarProdutos(ByVal CodProduto As String, Qtde As Single) As String
  '27/04/2004 - Daniel
  Dim rstProdutos    As Recordset
  
  Dim sngPesoLiquido As Single
  Dim sngPesoBruto As Single
  
  
  Set rstProdutos = db.OpenRecordset("SELECT PesoLiquido, PesoBruto FROM Produtos WHERE C�digo ='" & CodProduto & "'", dbOpenDynaset)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        '16/06/2004 - mpdea
        'Corrigido RT-94
        Call IsDataType(dtSingle, .Fields("PesoLiquido").Value, sngPesoLiquido)
        Call IsDataType(dtSingle, .Fields("PesoBruto").Value, sngPesoBruto)
        
        'C�lculo
        m_sngPesoBruto = (Qtde * sngPesoBruto) + m_sngPesoBruto
        m_sngPesoLiquido = (Qtde * sngPesoLiquido) + m_sngPesoLiquido
      
        .MoveNext
      Loop
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing


End Function

'27/09/2004 - mpdea
'Calcula o total de volumagem (inteiro) por quantidade na movimenta��o
Private Sub CalculaQtde()
  Dim rstSaidasProdutos As Recordset
  Dim strSQL As String
  Dim strCodProduto As String
  Dim sngQtde As Single
  Dim intVolumagem As Integer
  Dim intVolumagemQtdeTotal As Integer
  
  
  strSQL = "SELECT C�digo, Qtde "
  strSQL = strSQL & "FROM [Sa�das - Produtos] "
  strSQL = strSQL & "WHERE Sequ�ncia = " & lngSequencia
  strSQL = strSQL & " AND Filial = " & gnCodFilial
  
  Set rstSaidasProdutos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstSaidasProdutos
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        strCodProduto = .Fields("C�digo").Value
        sngQtde = .Fields("Qtde").Value
        intVolumagem = m_intGetProdutoVolumagem(strCodProduto)
        
        'Totalizador
        If intVolumagem > 0 Then
          intVolumagemQtdeTotal = intVolumagemQtdeTotal + (sngQtde \ intVolumagem)
        End If
      
      .MoveNext
      Loop
    End If
    .Close
  End With
  Set rstSaidasProdutos = Nothing
  
  'Exibe o total
  Qtde.Text = intVolumagemQtdeTotal

End Sub

'27/09/2004 - mpdea
'Obt�m a volumagem do produto
Private Function m_intGetProdutoVolumagem(ByVal strCodigo As String) As Integer
  Dim rstGet As Recordset
  Dim strSQL As String
  Dim intVolumagem As Integer
  
  strSQL = "SELECT Volumagem FROM Produtos WHERE C�digo = '" & strCodigo & "'"
  Set rstGet = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstGet
    If Not (.BOF And .EOF) Then
      Call IsDataType(dtInteger, .Fields("Volumagem").Value, intVolumagem)
    End If
    .Close
  End With
  Set rstGet = Nothing
  
  'Retorna a volumagem
  m_intGetProdutoVolumagem = intVolumagem
  
End Function

