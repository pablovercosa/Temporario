VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEstoquePreco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Estoque das Filiais e Preço (Personalizado)"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelEstoquePreco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   7890
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtrar relatório por (Classe e Sub Classe):"
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   2760
      Width           =   7575
      Begin VB.Data datSubClasses 
         Caption         =   "datSubClasses"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5280
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "con_Sub_Classe"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data datClasses 
         Caption         =   "datClasses"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1560
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "con_Classe"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtSubclasseNome 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtClasseNome 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmRelEstoquePreco.frx":058A
         DataSource      =   "datClasses"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
         DataFieldList   =   "Código"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   3
         Columns(0).FieldLen=   256
         Columns(1).Width=   7064
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmRelEstoquePreco.frx":05A3
         DataSource      =   "datSubClasses"
         Height          =   315
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   855
         DataFieldList   =   "Código"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Código"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Código"
         Columns(0).DataType=   3
         Columns(0).FieldLen=   256
         Columns(1).Width=   7064
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame fraSaida 
      Caption         =   "Saída"
      Height          =   735
      Left            =   3960
      TabIndex        =   23
      Top             =   3720
      Width           =   3735
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraOrdem 
      Caption         =   "Ordem"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   3735
      Begin VB.OptionButton optOrdemCodigo 
         Caption         =   "Código"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optOrdemNome 
         Caption         =   "Nome"
         Height          =   225
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport rptReport 
      Left            =   120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame fraTabelasPrecos 
      Caption         =   "Exibir preço para as Tabelas de Preços selecionadas abaixo:"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   7575
      Begin VB.ComboBox cboTabelaPreco 
         Height          =   315
         Index           =   1
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cboTabelaPreco 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblTabelaPreco 
         AutoSize        =   -1  'True
         Caption         =   "2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   21
         Top             =   420
         Width           =   180
      End
      Begin VB.Label lblTabelaPreco 
         AutoSize        =   -1  'True
         Caption         =   "1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   420
         Width           =   180
      End
   End
   Begin VB.Frame fraFiliais 
      Caption         =   "Exibir estoque para as Filiais selecionadas abaixo:"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7575
      Begin VB.Data datFiliais 
         Caption         =   "datFiliais"
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
         Height          =   375
         Left            =   5160
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
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
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5895
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
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   5895
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
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelEstoquePreco.frx":05BF
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   360
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Filial"
         Columns(0).Name =   "Filial"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Filial"
         Columns(0).DataType=   2
         Columns(0).FieldLen=   256
         Columns(1).Width=   8811
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelEstoquePreco.frx":05D8
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   1
         Top             =   720
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Filial"
         Columns(0).Name =   "Filial"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Filial"
         Columns(0).DataType=   2
         Columns(0).FieldLen=   256
         Columns(1).Width=   8811
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelEstoquePreco.frx":05F1
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Top             =   1080
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   1773
         Columns(0).Caption=   "Filial"
         Columns(0).Name =   "Filial"
         Columns(0).Alignment=   1
         Columns(0).CaptionAlignment=   1
         Columns(0).DataField=   "Filial"
         Columns(0).DataType=   2
         Columns(0).FieldLen=   256
         Columns(1).Width=   8811
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Nome"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label lblFilial 
         AutoSize        =   -1  'True
         Caption         =   "3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lblFilial 
         AutoSize        =   -1  'True
         Caption         =   "2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblFilial 
         AutoSize        =   -1  'True
         Caption         =   "1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   420
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   4680
      Width           =   1695
   End
End
Attribute VB_Name = "frmRelEstoquePreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'19/01/2006 - mpdea
'CASE: Kilouça (QS71271-970)
'Implementado form para relatório personalizado

'24/01/2006 - mpdea
'Adicionado filtro para Classe e Subclasse

Private rstPrecos As Recordset

Private Sub cboClasse_CloseUp()
  Call cboClasse_LostFocus
End Sub

Private Sub cboClasse_LostFocus()
  Dim intClasse As Integer
  
  
  On Error GoTo ErrHandler
  
  
  txtClasseNome.Text = ""
  
  If cboClasse.Text <> "" Then
    If Not IsDataType(dtInteger, cboClasse.Text, intClasse) Then
      DisplayMsg "Classe inválida."
      cboClasse.Text = ""
      Exit Sub
    End If
    
    If intClasse < 1 Or intClasse > 9999 Then
      DisplayMsg "Classe inválida."
      cboClasse.Text = ""
      Exit Sub
    End If
    
    With datClasses.Recordset
      .FindFirst "Código = " & intClasse
      If Not .NoMatch Then
        txtClasseNome.Text = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboFilial_CloseUp(Index As Integer)
  Call cboFilial_LostFocus(Index)
End Sub

Private Sub cboFilial_LostFocus(Index As Integer)
  Dim bytFillial As Byte
  
  
  On Error GoTo ErrHandler
  
  
  txtNomeFilial(Index).Text = ""
  
  If cboFilial(Index).Text <> "" Then
    If Not IsDataType(dtByte, cboFilial(Index).Text, bytFillial) Then
      DisplayMsg "Filial inválida."
      cboFilial(Index).Text = ""
      Exit Sub
    End If
    
    If bytFillial < 1 Or bytFillial > 99 Then
      DisplayMsg "Filial inválida."
      cboFilial(Index).Text = ""
      Exit Sub
    End If
    
    If Filial_Liberada <> 0 Then
      If bytFillial <> Filial_Liberada Then
        DisplayMsg "Funcionário não tem acesso a esta filial."
        cboFilial(Index).Text = ""
        Exit Sub
      End If
    End If
    
    With datFiliais.Recordset
      .FindFirst "Filial = " & bytFillial
      If Not .NoMatch Then
        txtNomeFilial(Index).Text = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub cboSubClasse_CloseUp()
  Call cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim intSubClasse As Integer
  
  
  On Error GoTo ErrHandler
  
  
  txtSubclasseNome.Text = ""
  
  If cboSubClasse.Text <> "" Then
    If Not IsDataType(dtInteger, cboSubClasse.Text, intSubClasse) Then
      DisplayMsg "Sub Classe inválida."
      cboSubClasse.Text = ""
      Exit Sub
    End If
    
    If intSubClasse < 1 Or intSubClasse > 9999 Then
      DisplayMsg "Sub Classe inválida."
      cboSubClasse.Text = ""
      Exit Sub
    End If
    
    With datSubClasses.Recordset
      .FindFirst "Código = " & intSubClasse
      If Not .NoMatch Then
        txtSubclasseNome.Text = .Fields("Nome").Value & ""
      End If
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmdImprimir_Click()
  
  On Error GoTo ErrHandler
  
  'Status
  Screen.MousePointer = vbHourglass
  cmdImprimir.Enabled = False
  Call StatusMsg("Obtendo informações...")

  'Preenche a tabela temporária
  Call FillTempData
  
  'Status
  Call StatusMsg("Imprimindo relatório...")
  
  'Prepara Relatório
  With rptReport
    .ReportFileName = gsReportPath & "rptEstoquePreco.rpt"
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 rptReport
    
    'Seta localização das bases de dados
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    
    'Ordem
    If optOrdemCodigo.Value Then
      .SortFields(0) = "+{tblRelEstoquePrecos.codigo_ordenacao}"
    Else
      .SortFields(0) = "+{tblRelEstoquePrecos.nome}"
    End If
    
    'Saída
    If optSaidaVideo.Value Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    
    'Nome da empresa
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", rptReport)

   'Exibe relatório
    .Action = 1
  End With

  'Status
  Call StatusMsg("Pronto")
  cmdImprimir.Enabled = True
  Screen.MousePointer = vbDefault

  Exit Sub
  
ErrHandler:
  'Status
  Call StatusMsg("Erro")
  cmdImprimir.Enabled = True
  Screen.MousePointer = vbDefault
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub FillTempData()
  Dim strSQL As String
  Dim rstReport As Recordset
  Dim rstProdutos As Recordset
  
  Dim intX As Integer
  
  'Array com as filiais informadas
  Dim intEstoqueFilial(2) As Integer
  'Array com as tabela de preçoc informadas
  Dim strTabelaPreco(1) As String
  
  Dim dblEstoque As Double
  Dim intErro As Integer
  
  Dim intClasse As Integer
  Dim intSubClasse As Integer
  
  
  On Error GoTo ErrHandler
  
  
  'Obtém filiais
  For intX = 0 To 2
    If txtNomeFilial(intX).Text <> "" Then
      intEstoqueFilial(intX) = CInt(cboFilial(intX).Text)
    End If
  Next intX
  
  'Obtém tabelas de preços
  For intX = 0 To 1
    If cboTabelaPreco(intX).Text <> "" Then
      strTabelaPreco(intX) = cboTabelaPreco(intX).Text
    End If
  Next intX
  
  'Obtém filtro
  'Classe
  If txtClasseNome.Text <> "" Then
    intClasse = CInt(cboClasse.Text)
  End If
  'Subclasse
  If txtSubclasseNome.Text <> "" Then
    intSubClasse = CInt(cboSubClasse.Text)
  End If
  
  'Cabeçalhos
  With rptReport
    'Estoque das Filiais
    For intX = 0 To 2
      If intEstoqueFilial(intX) > 0 Then
        .Formulas(intX + 1) = "estoque_filial_" & intX + 1 & _
          " = '" & intEstoqueFilial(intX) & "'"
      Else
        .Formulas(intX + 1) = "estoque_filial_" & intX + 1 & " = ''"
      End If
    Next intX
    'Tabelas de Preços
    For intX = 0 To 1
      .Formulas(intX + 4) = "tabela_preco_" & intX + 1 & _
        " = '" & strTabelaPreco(intX) & "'"
    Next intX
  End With
  
  
  'Limpa tabela temporária
  dbTemp.Execute "DELETE FROM tblRelEstoquePrecos;", dbFailOnError
  
  'Índice da tabela de Preços
  rstPrecos.Index = "Tabela"
  
  'Abre tabela temporária
  Set rstReport = dbTemp.OpenRecordset("tblRelEstoquePrecos", dbOpenDynaset)
  
  'Seleciona produtos
  strSQL = "SELECT Código, Nome, [Código Ordenação], "
  strSQL = strSQL & "[Código do Fornecedor], Classe, [Sub Classe] "
  strSQL = strSQL & "FROM Produtos "
  strSQL = strSQL & "WHERE Tipo = 'N' AND NOT Desativado"
  
  'Filtro
  If intClasse > 0 Then
    strSQL = strSQL & " AND Classe = " & intClasse
  End If
  If intSubClasse > 0 Then
    strSQL = strSQL & " AND [Sub Classe] = " & intSubClasse
  End If
  
  Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstProdutos
    Do Until .EOF
      
      With rstReport
        .AddNew
        .Fields("codigo").Value = rstProdutos.Fields("Código").Value
        .Fields("codigo_ordenacao").Value = rstProdutos.Fields("Código Ordenação").Value
        .Fields("nome").Value = rstProdutos.Fields("Nome").Value
        .Fields("codigo_fornecedor").Value = rstProdutos.Fields("Código do Fornecedor").Value
        
        'Estoque
        For intX = 0 To 2
          If intEstoqueFilial(intX) > 0 Then
            dblEstoque = Acha_Estoque(intEstoqueFilial(intX), .Fields("codigo").Value, 0, 0, 0, intErro)
            If intErro = 0 Then
              .Fields("qtde_filial_" & intX + 1).Value = dblEstoque
            Else
              .Fields("qtde_filial_" & intX + 1).Value = Null
            End If
          Else
            .Fields("qtde_filial_" & intX + 1).Value = Null
          End If
        Next intX
        
        'Preço
        For intX = 0 To 1
          If strTabelaPreco(intX) <> "" Then
            rstPrecos.Seek "=", strTabelaPreco(intX), .Fields("codigo").Value
            If rstPrecos.NoMatch Then
              .Fields("preco_" & intX + 1).Value = Null
            Else
              .Fields("preco_" & intX + 1).Value = rstPrecos("Preço")
            End If
          Else
            .Fields("preco_" & intX + 1).Value = 0
          End If
        Next intX
        
        .Fields("classe").Value = rstProdutos.Fields("Classe").Value
        .Fields("subclasse").Value = rstProdutos.Fields("Sub Classe").Value
        .Update
      End With
      
      .MoveNext
    Loop
    .Close
  End With
  Set rstProdutos = Nothing
  
  rstReport.Close
  Set rstReport = Nothing
  
  Exit Sub
  
ErrHandler:
  'Fecha tabelas
  If Not rstProdutos Is Nothing Then
    rstProdutos.Close
    Set rstProdutos = Nothing
  End If
  '
  If Not rstReport Is Nothing Then
    rstReport.Close
    Set rstReport = Nothing
  End If
  
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

Private Sub Form_Load()
  Dim rstTabelasPrecos As Recordset
  
  On Error GoTo ErrHandler

  Call StatusMsg("")
  
  Call CenterForm(Me)
  
  'Tabela de Preços
  Set rstPrecos = db.OpenRecordset("Preços", , dbReadOnly)
  
  'Seta Banco de dados para controles Data
  datFiliais.DatabaseName = gsQuickDBFileName
  datClasses.DatabaseName = gsQuickDBFileName
  datSubClasses.DatabaseName = gsQuickDBFileName
  
  'Preenche com as tabelas de preços existentes
  cboTabelaPreco(0).Clear
  cboTabelaPreco(1).Clear
  Set rstTabelasPrecos = db.OpenRecordset("SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela", dbOpenDynaset, dbReadOnly)
  With rstTabelasPrecos
    If Not (.BOF And .EOF) Then
      Do While Not .EOF
        cboTabelaPreco(0).AddItem .Fields("Tabela") & ""
        cboTabelaPreco(1).AddItem .Fields("Tabela") & ""
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rstTabelasPrecos = Nothing
  
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not rstPrecos Is Nothing Then
    rstPrecos.Close
    Set rstPrecos = Nothing
  End If
  
  Call StatusMsg("")
End Sub
