VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmEtiquetaEnderecamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiqueta de endereçamento"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEtiquetaEnderecamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   7095
   Begin VB.Data datFiliais 
      Caption         =   "datFiliais"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin SSDataWidgets_B.SSDBCombo cboCodigoFilial 
      Bindings        =   "frmEtiquetaEnderecamento.frx":058A
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
      DataFieldList   =   "Filial"
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Filial"
   End
   Begin Crystal.CrystalReport crtView 
      Left            =   120
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
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
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   -120
      TabIndex        =   17
      Top             =   -120
      Width           =   8175
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Use os campos abaixo como filtro para as etiquetas a serem emitidas pelo sistema."
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão de etiqueta de endereçamento"
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
         TabIndex        =   18
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sequência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      TabIndex        =   15
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox txtSequenciaFim 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtSequenciaInicio 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox txtNotaFim 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNotaInicio 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Fim"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3375
      Begin MSMask.MaskEdBox mskDataFim 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataInicio 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Fim"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblNomeFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   21
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Filial"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "frmEtiquetaEnderecamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCodigoFilial_CloseUp()
  cboCodigoFilial.Text = cboCodigoFilial.Columns(0).Text
  cboCodigoFilial_LostFocus
End Sub

Private Sub cboCodigoFilial_LostFocus()
  lblNomeFilial.Caption = ""
  With cboCodigoFilial
    If IsNumeric(.Text) Then
      datFiliais.Recordset.FindFirst " Filial = " & .Text
      If Not datFiliais.Recordset.NoMatch Then
        lblNomeFilial.Caption = datFiliais.Recordset.Fields("Nome") & ""
      End If
    End If
  End With
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim strSelection As String
  
  With crtView
    .Reset
    .ReportFileName = gsReportPath & "EtiquetaEnderecamento.rpt"
    .DataFiles(0) = gsQuickDBFileName
    '18/02/2004 - Daniel
    'Acrescentado o .DataFiles(1)
    .DataFiles(1) = gsQuickDBFileName
    
    strSelection = " {Saídas.Cliente} <> 0 "
    
    If IsDate(mskDataInicio.Text) Then
      strSelection = strSelection & " AND {Saídas.Data} >= Date" & Format(mskDataInicio.Text, "(yyyy,mm,dd)")
    End If
    
    If IsDate(mskDataFim.Text) Then
      strSelection = strSelection & " AND {Saídas.Data} <= Date" & Format(mskDataFim.Text, "(yyyy,mm,dd)")
    End If
    
    If IsNumeric(txtNotaInicio.Text) Then
      strSelection = strSelection & " AND {Saídas.Nota Impressa} >= " & txtNotaInicio.Text
    End If
    
    If IsNumeric(txtNotaFim.Text) Then
      strSelection = strSelection & " AND {Saídas.Nota Impressa} <= " & txtNotaFim.Text
    End If
    
    If IsNumeric(txtSequenciaInicio.Text) Then
      strSelection = strSelection & " AND {Saídas.Sequência} >= " & txtSequenciaInicio.Text
    End If
    
    If IsNumeric(txtSequenciaFim.Text) Then
      strSelection = strSelection & " AND {Saídas.Sequência} >= " & txtSequenciaFim.Text
    End If
    
    .SelectionFormula = strSelection
    .Action = 1
  End With
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  datFiliais.DatabaseName = gsQuickDBFileName
End Sub

Private Sub mskDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then mskDataFim.Text = frmCalendario.gsDateCalender(mskDataFim.Text)
End Sub

Private Sub mskDataFim_LostFocus()
  mskDataFim.Text = Ajusta_Data(mskDataFim.Text)
End Sub

Private Sub mskDataInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then mskDataInicio.Text = frmCalendario.gsDateCalender(mskDataInicio.Text)
End Sub

Private Sub mskDataInicio_LostFocus()
  mskDataInicio.Text = Ajusta_Data(mskDataInicio.Text)
End Sub
