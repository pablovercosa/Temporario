VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Produtos"
   ClientHeight    =   7620
   ClientLeft      =   1170
   ClientTop       =   1830
   ClientWidth     =   15765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1480
   Icon            =   "RelProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7620
   ScaleWidth      =   15765
   Begin VB.Frame frm_tipo 
      Caption         =   "Tipo de busca de dados"
      Height          =   765
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15705
      Begin VB.OptionButton opt_02 
         Caption         =   "Customizado"
         Height          =   255
         Left            =   9480
         TabIndex        =   2
         Top             =   330
         Width           =   1275
      End
      Begin VB.OptionButton opt_01 
         Caption         =   "Padrão"
         Height          =   255
         Left            =   3330
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame frm_01 
      Height          =   6735
      Left            =   30
      TabIndex        =   3
      Top             =   810
      Width           =   15705
      Begin VB.Frame Frame2 
         Caption         =   "Detalhes"
         Height          =   1185
         Left            =   90
         TabIndex        =   25
         Top             =   2520
         Width           =   6015
         Begin VB.OptionButton O_M_Detalhado 
            Appearance      =   0  'Flat
            Caption         =   "Muito detalhado"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   30
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton O_Detalhado 
            Appearance      =   0  'Flat
            Caption         =   "Detalhado"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   29
            Top             =   690
            Width           =   1335
         End
         Begin VB.OptionButton O_Normal 
            Appearance      =   0  'Flat
            Caption         =   "Normal"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2370
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton O_P_Detalhado 
            Appearance      =   0  'Flat
            Caption         =   "Pouco detalhado"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4050
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton O_Simples 
            Appearance      =   0  'Flat
            Caption         =   "Simples"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4050
            TabIndex        =   26
            Top             =   690
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem"
         Height          =   1185
         Left            =   6240
         TabIndex        =   21
         Top             =   2520
         Width           =   4125
         Begin VB.OptionButton O_Código 
            Appearance      =   0  'Flat
            Caption         =   "Código"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   420
            TabIndex        =   24
            Top             =   510
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton O_Nome 
            Appearance      =   0  'Flat
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1590
            TabIndex        =   23
            Top             =   510
            Width           =   855
         End
         Begin VB.OptionButton O_Classe 
            Appearance      =   0  'Flat
            Caption         =   "Classe"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2760
            TabIndex        =   22
            Top             =   510
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saída"
         Height          =   1170
         Left            =   10500
         TabIndex        =   18
         Top             =   2520
         Width           =   5055
         Begin VB.OptionButton optVideo 
            Appearance      =   0  'Flat
            Caption         =   "Vídeo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   510
            TabIndex        =   20
            Top             =   510
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optPrinter 
            Appearance      =   0  'Flat
            Caption         =   "Impressora"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2010
            TabIndex        =   19
            Top             =   510
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gerar Relatório"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3840
         Width           =   15480
      End
      Begin VB.Data datProduto 
         Caption         =   "Produto"
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
         Left            =   4350
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Produto"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Data datSubClasse 
         Caption         =   "SubClasse"
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
         Left            =   2490
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Sub_Classe"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Data datClasse 
         Caption         =   "Classe"
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
         Left            =   630
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Con_Classe"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Frame Frame5 
         Caption         =   "Opções"
         Height          =   2235
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   15480
         Begin VB.OptionButton optOption 
            Appearance      =   0  'Flat
            Caption         =   "&Todos os Produtos"
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   2250
         End
         Begin VB.OptionButton optOption 
            Appearance      =   0  'Flat
            Caption         =   "&Um produto"
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   6990
            TabIndex        =   6
            Top             =   270
            Width           =   1350
         End
         Begin VB.CheckBox chkInativos 
            Appearance      =   0  'Flat
            Caption         =   "Considerar também os produtos Inativos"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   150
            TabIndex        =   5
            Top             =   1710
            Width           =   3330
         End
         Begin SSDataWidgets_B.SSDBCombo cboProduto 
            Bindings        =   "RelProdutos.frx":4E95A
            DataSource      =   "datProduto"
            Height          =   315
            Left            =   7830
            TabIndex        =   8
            Top             =   645
            Width           =   1815
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
            Columns(0).Width=   8096
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Nome"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3651
            Columns(1).Caption=   "Código"
            Columns(1).Name =   "Codigo"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Código"
            Columns(1).DataType=   5
            Columns(1).FieldLen=   256
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
            Bindings        =   "RelProdutos.frx":4E973
            DataSource      =   "datSubClasse"
            Height          =   315
            Left            =   1425
            TabIndex        =   9
            Top             =   1020
            Width           =   1095
            DataFieldList   =   "Código"
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
            Columns(0).Width=   7064
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Nome"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1931
            Columns(1).Caption=   "Código"
            Columns(1).Name =   "Codigo"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Código"
            Columns(1).DataType=   3
            Columns(1).FieldLen=   256
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B.SSDBCombo cboClasse 
            Bindings        =   "RelProdutos.frx":4E98E
            DataSource      =   "datClasse"
            Height          =   315
            Left            =   1425
            TabIndex        =   10
            Top             =   645
            Width           =   1095
            DataFieldList   =   "Código"
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
            Columns(0).Width=   7117
            Columns(0).Caption=   "Nome"
            Columns(0).Name =   "Nome"
            Columns(0).CaptionAlignment=   0
            Columns(0).DataField=   "Nome"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1588
            Columns(1).Caption=   "Código"
            Columns(1).Name =   "Codigo"
            Columns(1).Alignment=   1
            Columns(1).CaptionAlignment=   1
            Columns(1).DataField=   "Código"
            Columns(1).DataType=   3
            Columns(1).FieldLen=   256
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.Label lblNomeProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9720
            TabIndex        =   16
            Top             =   645
            Width           =   5640
         End
         Begin VB.Label Label1 
            Caption         =   "Código"
            Height          =   225
            Left            =   7245
            TabIndex        =   15
            Top             =   690
            Width           =   675
         End
         Begin VB.Label lblNomeSubClasse 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2550
            TabIndex        =   14
            Top             =   1020
            Width           =   3960
         End
         Begin VB.Label Label4 
            Caption         =   "Sub Classe"
            Height          =   255
            Left            =   450
            TabIndex        =   13
            Top             =   1050
            Width           =   930
         End
         Begin VB.Label lblNomeClasse 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2550
            TabIndex        =   12
            Top             =   645
            Width           =   3960
         End
         Begin VB.Label Label2 
            Caption         =   "Classe"
            Height          =   255
            Left            =   450
            TabIndex        =   11
            Top             =   675
            Width           =   645
         End
      End
      Begin Crystal.CrystalReport rptReport 
         Left            =   105
         Top             =   6210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Frame frm_02 
      Height          =   6735
      Left            =   30
      TabIndex        =   31
      Top             =   810
      Visible         =   0   'False
      Width           =   15705
      Begin VB.CommandButton cmd_imprimirOperDet 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   13950
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6300
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid gridOperacoesDetalhe 
         Height          =   3105
         Left            =   60
         TabIndex        =   33
         Top             =   3150
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   5477
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedCols       =   0
         BackColor       =   15066597
         BackColorFixed  =   8454143
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16250871
         AllowBigSelection=   0   'False
         SelectionMode   =   1
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
      End
   End
End
Attribute VB_Name = "frmRelProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboClasse_DropDown()
  cboClasse.DataFieldList = "Código"
End Sub

Private Sub cboClasse_KeyPress(KeyAscii As Integer)
  If cboClasse.DroppedDown Then
    cboClasse.DataFieldList = "Nome"
  End If
  If Len(cboClasse.Text) >= 4 Then
    If KeyAscii <> vbKeyBack And Not cboClasse.DroppedDown Then
      Beep
      KeyAscii = 0
      Exit Sub
    End If
  End If
'  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub cboClasse_Validate(Cancel As Boolean)
'  If cboClasse.Text <> "" And cboClasse.Text <> "0" Then
'    If IsNumeric(cboClasse.Text) Then
'      If Val(cboClasse.Text) > 0 And Val(cboClasse.Text) < 10000 Then
'        lblNomeClasse.Caption = gsGetNameClasse(cboClasse.Text)
'      Else
'        lblNomeClasse.Caption = ""
'      End If
'    Else
'      lblNomeClasse.Caption = ""
'    End If
'  Else
'    lblNomeClasse.Caption = ""
'  End If
  
  If cboClasse.DroppedDown Then
    cboClasse.DroppedDown = False
  End If
  If cboClasse.Text = "" Or cboClasse.Text = "0" Then
    lblNomeClasse.Caption = ""
  Else
    If IsNumeric(cboClasse.Text) Then
      datClasse.Recordset.FindFirst "Código = " & CInt(cboClasse.Text)
      If Not datClasse.Recordset.NoMatch Then
        lblNomeClasse.Caption = datClasse.Recordset(0).Value
      Else
        Beep
        cboClasse.Text = ""
        lblNomeClasse.Caption = ""
        Cancel = True
      End If
    Else
      Beep
      cboClasse.Text = ""
      lblNomeClasse.Caption = ""
      Cancel = True
    End If
  End If
End Sub

Private Sub cboSubClasse_DropDown()
  cboSubClasse.DataFieldList = "Código"
End Sub

Private Sub cboSubClasse_KeyPress(KeyAscii As Integer)
  If cboSubClasse.DroppedDown Then
    cboSubClasse.DataFieldList = "Nome"
  End If
  If Len(cboSubClasse.Text) >= 4 Then
    If KeyAscii <> vbKeyBack And Not cboSubClasse.DroppedDown Then
      Beep
      KeyAscii = 0
      Exit Sub
    End If
  End If
'  KeyAscii = gnSomenteNumero(KeyAscii)
End Sub

Private Sub cboSubClasse_Validate(Cancel As Boolean)
'  If cboSubClasse.Text <> "" And cboSubClasse.Text <> "0" Then
'    If IsNumeric(cboSubClasse.Text) Then
'      If Val(cboSubClasse.Text) > 0 And Val(cboSubClasse.Text) < 10000 Then
'        lblNomeSubClasse.Caption = gsGetNameSubClasse(cboSubClasse.Text)
'      Else
'        lblNomeSubClasse.Caption = ""
'      End If
'    Else
'      lblNomeSubClasse.Caption = ""
'    End If
'  Else
'    lblNomeSubClasse.Caption = ""
'  End If

  If cboSubClasse.DroppedDown Then
    cboSubClasse.DroppedDown = False
  End If
  If cboSubClasse.Text = "" Or cboSubClasse.Text = "0" Then
    lblNomeSubClasse.Caption = ""
  Else
    If IsNumeric(cboSubClasse.Text) Then
      datSubClasse.Recordset.FindFirst "Código = " & CInt(cboSubClasse.Text)
      If Not datSubClasse.Recordset.NoMatch Then
        lblNomeSubClasse.Caption = datSubClasse.Recordset(0).Value
      Else
        Beep
        cboSubClasse.Text = ""
        lblNomeSubClasse.Caption = ""
        Cancel = True
      End If
    Else
      Beep
      cboSubClasse.Text = ""
      lblNomeSubClasse.Caption = ""
      Cancel = True
    End If
  End If
End Sub

Private Sub cmdImprimir_Click()
  Dim sReport As String
  Dim sSql As String
  
  Call StatusMsg("")
  'Seleção
  If optOption(1).Value Then
    If lblNomeProduto.Caption <> "" Then
      sSql = "{Produtos.Código} = '" & cboProduto.Text & "'"
    Else
      DisplayMsg "Escolha um produto."
      Exit Sub
    End If
  Else
    sSql = "{Produtos.Código} <> '0'"
    If lblNomeClasse.Caption <> "" Then
      sSql = sSql & " And {Produtos.Classe} = " & cboClasse.Text
    End If
    If lblNomeSubClasse.Caption <> "" Then
      sSql = sSql & " And {Produtos.Sub Classe} = " & cboSubClasse.Text
    End If
    If chkInativos.Value = vbUnchecked Then
      sSql = sSql & " AND {Produtos.Desativado} = False"   '" AND NOT {Produtos.Desativado}"
    End If
  End If
  
  'Nome do arquivo .rpt
  If O_M_Detalhado.Value Then
    sReport = gsReportPath & "PRODUTO1.RPT"
  ElseIf O_Detalhado.Value Then
    sReport = gsReportPath & "PRODUTO2.RPT"
  ElseIf O_Normal.Value Then
    sReport = gsReportPath & "PRODUTO3.RPT"
  ElseIf O_P_Detalhado.Value Then
    sReport = gsReportPath & "PRODUTO4.RPT"
  ElseIf O_Simples.Value Then
    sReport = gsReportPath & "PRODUTO5.RPT"
  End If
  
  MousePointer = vbHourglass
  With rptReport
    .Reset
    .ReportFileName = sReport
    .DataFiles(0) = gsQuickDBFileName
    .SelectionFormula = sSql
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    If O_Código.Value Then
      .SortFields(0) = "+{Produtos.Código Ordenação}"
    ElseIf O_Nome.Value = True Then
      .SortFields(0) = "+{Produtos.Nome}"
    ElseIf O_Classe.Value Then
      .SortFields(0) = "+{Classes.Nome}"
      .SortFields(1) = "+{Produtos.Nome}"
    End If
    .WindowState = crptMaximized
    .Destination = IIf(optVideo.Value, crptToWindow, crptToPrinter)
    
    '25/01/2006 - mpdea
    'Exibe botão para configurar impressora
    .WindowShowPrintSetupBtn = True
    
    Call StatusMsg("Aguarde, imprimindo...")
  
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 rptReport
  
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", rptReport)
  
  
    .Action = 1
  End With
  MousePointer = vbDefault
  
  Call StatusMsg("")

End Sub

Private Sub cboClasse_CloseUp()
  cboClasse.DataFieldList = "Código"
  cboClasse.Text = cboClasse.Columns("Código").Text
  Call cboClasse_Validate(True)
End Sub

Private Sub cboClasse_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboClasse_LostFocus()
  Call StatusMsg("")
End Sub

Private Sub cboProduto_CloseUp()
  cboProduto.Text = cboProduto.Columns("Codigo").Text
  cboProduto_LostFocus
End Sub

Private Sub cboProduto_LostFocus()
  Call StatusMsg("")
  If cboProduto.Text <> "" Then
    lblNomeProduto.Caption = gsGetNameProduto(cboProduto.Text)
  Else
    lblNomeProduto.Caption = ""
  End If
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.DataFieldList = "Código"
  cboSubClasse.Text = cboSubClasse.Columns("Código").Text
  Call cboSubClasse_Validate(True)
End Sub

Private Sub cboSubClasse_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboSubClasse_LostFocus()
  Call StatusMsg("")
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  datProduto.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
End Sub

Private Sub opt_01_Click()
    If opt_01.Value = True Then
        frm_01.Visible = True
        frm_02.Visible = False
    Else
        frm_01.Visible = False
        frm_02.Visible = True
    End If
End Sub

Private Sub opt_02_Click()
    If opt_02.Value = True Then
        frm_02.Visible = True
        frm_01.Visible = False
    Else
        frm_02.Visible = False
        frm_01.Visible = True
    End If

End Sub

Private Sub optOption_Click(Index As Integer)
  If Index = 0 Then
    cboProduto.Enabled = False
    cboClasse.Enabled = True
    cboSubClasse.Enabled = True
    chkInativos.Enabled = True
  Else
    cboProduto.Enabled = True
    cboClasse.Enabled = False
    cboSubClasse.Enabled = False
    chkInativos.Enabled = False
  End If
End Sub
