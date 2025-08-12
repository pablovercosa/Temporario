VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEdicoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Edições"
   ClientHeight    =   5295
   ClientLeft      =   3675
   ClientTop       =   2355
   ClientWidth     =   4605
   Icon            =   "Edicoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   750
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Edições"
      Top             =   4860
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton B_Retorna 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1635
   End
   Begin SSDataWidgets_B.SSDBGrid Grade1 
      Bindings        =   "Edicoes.frx":4E95A
      Height          =   4215
      Left            =   105
      TabIndex        =   0
      Top             =   465
      Width           =   4425
      _Version        =   196617
      AllowAddNew     =   -1  'True
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Produto"
      Columns(0).Name =   "Produto"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Produto"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1482
      Columns(1).Caption=   "Edição"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   5
      Columns(2).Width=   5106
      Columns(2).Caption=   "Nome"
      Columns(2).Name =   "Nome"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Nome"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   7805
      _ExtentY        =   7435
      _StockProps     =   79
      Caption         =   "Edições"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   75
      Width           =   1650
   End
End
Attribute VB_Name = "frmEdicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Retorna_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Dim Rec_Edições As Recordset
  Dim sSql As String
  
  
  sSql = "SELECT Produto, Código, Nome FROM Edições"
  sSql = sSql + " WHERE Produto = '" + Trim(str(Código.Caption)) + "'"
    
  Set Rec_Edições = db.OpenRecordset(sSql)

 ' Grade1.DataMode = 1

  Set Data1.Recordset = Rec_Edições
   
  
 ' Grade1.DataMode = 0

  Grade1.ReBind
  Grade1.Columns(0).Visible = False
  Grade1.Columns(2).Width = 2800

End Sub

Private Sub Form_Load()
  Data1.DatabaseName = gsQuickDBFileName

End Sub


Private Sub Grade1_BeforeUpdate(Cancel As Integer)
 Dim Aux As Variant
 Dim Erro As Boolean
 
 Erro = False
 Aux = Grade1.Columns(1).Text
 If IsNull(Aux) Then Erro = True
 If Erro = False Then If Aux = "" Then Erro = True
 If Erro = False Then If Not IsNumeric(Aux) Then Erro = True
 If Erro = False Then If Val(Aux) < 1 Then Erro = True
 
 If Erro = True Then
   Cancel = True
   Exit Sub
 End If
 

 Grade1.Columns(0).Text = Código.Caption
 
 
End Sub

Private Sub Grade1_KeyPress(KeyAscii As Integer)
 Dim Tam As Integer

 If Grade1.Col = 1 Then
    Tam = Len(Grade1.Columns(1).Text)
    If Tam >= 5 Then
      KeyAscii = 0
      Exit Sub
    End If
    
 
    KeyAscii = Verifica_Tecla_Integer(KeyAscii)
 End If
 

End Sub


