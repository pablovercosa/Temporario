VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelAutorizacaoPublicidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Geral"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmRelAutorizacaoPublicidade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7575
   Begin Crystal.CrystalReport rptAutorizacao 
      Left            =   6960
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3495
      Width           =   1815
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   3495
      Begin MSMask.MaskEdBox mskPeriodoFin 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para obter calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskPeriodoIni 
         Height          =   315
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para obter calendário"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   18
         Top             =   420
         Width           =   90
      End
   End
   Begin VB.Frame fraSaida 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3720
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   260
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   520
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1490
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   7335
      Begin VB.TextBox txtVendedor 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   3495
      End
      Begin VB.Data datTipoComercial 
         Caption         =   "datTipoComercial"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Descricao FROM TipoComercial ORDER BY Código"
         Top             =   960
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.TextBox txtTipoComercial 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Data datRadio 
         Caption         =   "datRadio"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Radio ORDER BY Código"
         Top             =   600
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.TextBox txtRadio 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   660
         Width           =   3495
      End
      Begin VB.Data datVendedor 
         Caption         =   "datVendedor"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Funcionários ORDER BY Código"
         Top             =   120
         Visible         =   0   'False
         Width           =   1980
      End
      Begin SSDataWidgets_B.SSDBCombo cboRadio 
         Bindings        =   "frmRelAutorizacaoPublicidade.frx":058A
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   667
         Width           =   2295
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboTipoComercial 
         Bindings        =   "frmRelAutorizacaoPublicidade.frx":05A1
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelAutorizacaoPublicidade.frx":05C0
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   255
         Width           =   2295
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label4 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Comercial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rádio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   727
         Width           =   405
      End
   End
   Begin VB.Frame fraDica 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   -120
      TabIndex        =   9
      Top             =   -120
      Width           =   8175
      Begin VB.Label lblDica 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelAutorizacaoPublicidade.frx":05DA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Autorização de Publicidades"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmRelAutorizacaoPublicidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'27/01/2004 - Daniel
'Case: STC de Caxias do Sul - RS
'14/04/2004 - Manutenido conforme novas implementações

Private Sub cboRadio_CloseUp()
  cboRadio.Text = cboRadio.Columns(0).Text
  cboRadio_LostFocus
End Sub

Private Sub cboRadio_LostFocus()
  Dim rstRadio As Recordset
  
  txtRadio.Text = ""
  
  If Not IsNumeric(cboRadio.Text) Then Exit Sub
  
  Set rstRadio = db.OpenRecordset("SELECT Código, Nome FROM Radio WHERE Código = " & CInt(cboRadio.Text), dbOpenDynaset)
  
  With rstRadio
    If Not (.BOF And .EOF) Then
      txtRadio.Text = .Fields("Nome") & ""
    End If
  End With
  
  rstRadio.Close
  Set rstRadio = Nothing

End Sub

Private Sub cboTipoComercial_CloseUp()
  cboTipoComercial.Text = cboTipoComercial.Columns(0).Text
  cboTipoComercial_LostFocus
End Sub

Private Sub cboTipoComercial_LostFocus()
  Dim rstTipoComercial As Recordset
  
  txtTipoComercial.Text = ""
  
  If Not IsNumeric(cboTipoComercial.Text) Then Exit Sub
  
  Set rstTipoComercial = db.OpenRecordset("SELECT Código, Descricao FROM TipoComercial WHERE Código = " & CInt(cboTipoComercial.Text), dbOpenDynaset)
  
  With rstTipoComercial
    If Not (.BOF And .EOF) Then
      txtTipoComercial.Text = .Fields("Descricao") & ""
    End If
  End With
  
  rstTipoComercial.Close
  Set rstTipoComercial = Nothing

End Sub

Private Sub cboVendedor_CloseUp()
  cboVendedor.Text = cboVendedor.Columns(0).Text
  cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_LostFocus()
  Dim rstFuncionarios As Recordset
  
  txtVendedor.Text = ""
  
  If Not IsNumeric(cboVendedor.Text) Then Exit Sub
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Código, Nome FROM Funcionários WHERE Código = " & CInt(cboVendedor.Text), dbOpenDynaset)
  
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      txtVendedor.Text = .Fields("Nome") & ""
    End If
  End With
  
  rstFuncionarios.Close
  Set rstFuncionarios = Nothing
  
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim blnErro         As Boolean
  Dim strReport       As String
  Dim strSQL          As String
  Dim strDataIni      As String
  Dim strDataFin      As String
  
  Call StatusMsg("")
  'Verificação dos campos datas
  If Not IsDate(mskPeriodoIni.Text) Then
    MsgBox "Período Inicial inválido, verifique.", vbExclamation, "Quick Store"
    mskPeriodoIni.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskPeriodoFin.Text) Then
    MsgBox "Período Final inválido, verifique.", vbExclamation, "Quick Store"
    
    Exit Sub
  End If
  
  If CDate(mskPeriodoIni.Text) > CDate(mskPeriodoFin.Text) Then
    MsgBox "Período Final menor que o Período Inicial, verifique.", vbExclamation, "Quick Store"
    Exit Sub
  End If
  '--------------------------------------------------------------------------------
    
    strSQL = " {Programacao.Num Autorizacao} = {Contrato.Num Autorizacao} "
    
    If (Len(Trim(txtVendedor.Text)) > 0) Then  'Selecionou algum vendedor
      strSQL = strSQL & " AND {Contrato.Cod Vendedor} = " & CInt(cboVendedor.Text)
    End If
    
    If (Len(Trim(txtRadio.Text)) > 0) Then   'Selecionou alguma rádio
      strSQL = strSQL & " AND {Contrato.Cod Radio} = " & CInt(cboRadio.Text)
    End If
      
    If (Len(Trim(txtTipoComercial)) > 0) Then  'Selecionou algum tipo
      strSQL = strSQL & " AND {Contrato.Cod TipoComercial} = " & CInt(cboTipoComercial.Text)
    End If
    
    'Tratamento para as datas
    strDataIni = "Date" + Format$(mskPeriodoIni.Text, "(yyyy,mm,dd)")
    strDataFin = "Date" + Format$(mskPeriodoFin.Text, "(yyyy,mm,dd)")
 
    strSQL = strSQL + " AND {Contrato.Data Assinatura} >=" + strDataIni
    strSQL = strSQL + " AND {Contrato.Data Assinatura} <=" + strDataFin
    'Fim do Tratamento para as datas
    
    'Mostrar o que foi faturado (confirmado o recebimento)
    strSQL = strSQL + " AND {Programacao.Faturado} "
    
  '--------------------------------------------------------------------------------
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptAutorizacoesPublicidade.rpt"
  MousePointer = vbHourglass
  
  With rptAutorizacao
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 rptAutorizacao
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsQuickDBFileName
    .DataFiles(5) = gsQuickDBFileName
    .DataFiles(6) = gsTempDBFileName
    .DataFiles(7) = gsQuickDBFileName
        
    .SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    '.SortFields(0) = "+{Contrato.Data Assinatura}" 'Ordenação
    
    .WindowState = crptMaximized
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", rptAutorizacao)
  
    .Action = 1
  End With
  
  MousePointer = vbDefault
  
  Call StatusMsg("")
End Sub

Private Sub Form_Load()

  datVendedor.DatabaseName = gsQuickDBFileName
  datRadio.DatabaseName = gsQuickDBFileName
  datTipoComercial.DatabaseName = gsQuickDBFileName
  
  Call CenterForm(Me)
  
End Sub

Private Sub mskPeriodoFin_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla está pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskPeriodoFin.Text = frmCalendario.gsDateCalender(mskPeriodoFin.Text)
  End If
End Sub

Private Sub mskPeriodoFin_LostFocus()
  mskPeriodoFin.Text = Ajusta_Data(mskPeriodoFin.Text)
End Sub

Private Sub mskPeriodoIni_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla está pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskPeriodoIni.Text = frmCalendario.gsDateCalender(mskPeriodoIni.Text)
  End If
End Sub

Private Sub mskPeriodoIni_LostFocus()
  mskPeriodoIni.Text = Ajusta_Data(mskPeriodoIni.Text)
End Sub
