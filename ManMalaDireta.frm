VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManMalaDireta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção da Mala Direta"
   ClientHeight    =   5550
   ClientLeft      =   1530
   ClientTop       =   1665
   ClientWidth     =   9375
   HelpContextID   =   1330
   Icon            =   "ManMalaDireta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   9375
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   285
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6195
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Monta 
      Caption         =   "&Pesquisar"
      Height          =   400
      Left            =   7950
      TabIndex        =   2
      Top             =   570
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   5985
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome do Cliente"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   525
         Width           =   1575
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código do Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Bindings        =   "ManMalaDireta.frx":058A
      Height          =   4335
      Left            =   135
      TabIndex        =   3
      Top             =   1110
      Width           =   9165
      _Version        =   196617
      AllowDelete     =   -1  'True
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      RowHeight       =   423
      Columns(0).Width=   3200
      UseDefaults     =   0   'False
      _ExtentX        =   16166
      _ExtentY        =   7646
      _StockProps     =   79
      Caption         =   "Etiquetas a Serem Emitidas. Cada Linha Representa uma Etiqueta."
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
End
Attribute VB_Name = "frmManMalaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rec_Mala As Recordset

Private Sub B_Monta_Click()
 Dim sSql As String
 Dim i As Integer
 Dim Registros As Long
 Dim Mensa As String
 
 
  sSql = "SELECT Ordem, Cliente, Cli_For.Nome, [Mala Direta - Tempo].Nome FROM [Mala Direta - Tempo]"
  sSql = sSql + " INNER JOIN Cli_For ON [Mala Direta - Tempo].Cliente = Cli_For.Código"
  
  If O_Código.Value = True Then sSql = sSql + " ORDER BY Cliente"
  If O_Nome.Value = True Then sSql = sSql + " ORDER BY Cli_For.Nome"
  
  
  
  Call StatusMsg("Aguarde, montando tabela...")
  DoEvents
  
  
  Set Rec_Mala = db.OpenRecordset(sSql, dbOpenDynaset)
  On Error GoTo Prossegue
  Registros = 0
  Rec_Mala.MoveLast
  
  Registros = Rec_Mala.RecordCount

Prossegue:
  On Error GoTo 0
  Grade1.DataMode = 1
  Set Data1.Recordset = Rec_Mala

  Grade1.Visible = False

  'Colunas
  '0 Ordem
  '1 Cliente
  '2 Nome Cliente
  '3 Contato
  
  Grade1.DataMode = 0
  Grade1.ReBind
    Grade1.Columns(0).Visible = False
    
    Grade1.Columns(1).Width = 900
    Grade1.Columns(1).Locked = True
    Grade1.Columns(2).Width = 3840
    Grade1.Columns(2).Locked = True
    'Grade1.Columns(3).Header = "Nome"
    Grade1.Columns(3).Width = 3500
    
  Grade1.Visible = True
  Call StatusMsg("")
  Mensa = "Total de Etiquetas :" + str(Registros)
  DisplayMsg Mensa

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub Grade1_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If Not bGridBeforeDelete() = True Then
    Cancel = True
  End If
End Sub
