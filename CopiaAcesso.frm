VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCopiaAcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cópia de Permissões "
   ClientHeight    =   3240
   ClientLeft      =   1260
   ClientTop       =   2400
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CopiaAcesso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3240
   ScaleWidth      =   7395
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
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
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2730
      Width           =   7125
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   7125
   End
   Begin VB.Data Data1 
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
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Funcionários"
      Top             =   990
      Visible         =   0   'False
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Novo 
      Bindings        =   "CopiaAcesso.frx":4E95A
      DataSource      =   "Data1"
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   1380
      Width           =   1095
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
      BackColorOdd    =   16777152
      Columns(0).Width=   3200
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Nome_Novo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1320
      TabIndex        =   6
      Top             =   1380
      Width           =   5985
   End
   Begin VB.Label Label2 
      Caption         =   "Funcionário do qual os acessos devem ser copiados"
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1110
      Width           =   4005
   End
   Begin VB.Label Nome_Func 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1320
      TabIndex        =   4
      Top             =   390
      Width           =   5985
   End
   Begin VB.Label Cod_Func 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   180
      TabIndex        =   3
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Funcionário Atual"
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmCopiaAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsFuncionarios As Recordset

Private Sub cmdCopy_Click()
  Dim sSql As String
  Dim rsAcessos As Recordset
  Dim rsAcessos2 As Recordset
  Dim sNum As String
  Dim sProg As String
  Dim bGravar As Boolean
  Dim bApagar As Boolean
  
  If Nome_Novo.Caption = "" Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(245)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Combo_Novo.SetFocus
    Exit Sub
  End If
  
  If Val(Cod_Func.Caption) = Val(Combo_Novo.Text) Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(246)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Call StatusMsg(LoadResString(244))
  
  
  sSql = "Delete * From Acessos Where Usuário =" + Cod_Func.Caption
  Call db.Execute(sSql, dbFailOnError)
  
  sSql = "SELECT * FROM Acessos WHERE Usuário = " & Val(Combo_Novo.Text) & " ORDER BY Usuário, Numero"
  Set rsAcessos = db.OpenRecordset(sSql, dbOpenDynaset)
  
  Set rsAcessos2 = db.OpenRecordset("SELECT * FROM Acessos", dbOpenDynaset)
  
  Do While Not rsAcessos.EOF
    DoEvents
    sNum = rsAcessos("Numero").Value
    sProg = rsAcessos("Programa").Value
    bGravar = rsAcessos("Gravar").Value
    bApagar = rsAcessos("Apagar").Value
    With rsAcessos2
      .AddNew
      .Fields("Usuário") = Val(Cod_Func.Caption)
      .Fields("Numero") = sNum
      .Fields("Programa") = sProg
      .Fields("Gravar") = bGravar
      .Fields("Apagar") = bApagar
      .Update
    End With
    Call StatusMsg(sProg)
    rsAcessos.MoveNext
  Loop
  
  rsAcessos.Close
  Set rsAcessos = Nothing
  rsAcessos2.Close
  Set rsAcessos2 = Nothing
  
  gsTitle = LoadResString(201)
  gsMsg = LoadResString(243)
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  gbCopyPermissoes = True
  Unload Me
  
End Sub

Private Sub cmdCancel_Click()
  gbCopyPermissoes = False
  Unload Me
End Sub

Private Sub Combo_Novo_CloseUp()
  Combo_Novo.Text = Combo_Novo.Columns(2).Text
  Combo_Novo_LostFocus
End Sub

Private Sub Combo_Novo_LostFocus()
  Dim Aux As Variant
  
  Call StatusMsg("")
  Nome_Novo.Caption = ""
  
  Aux = Combo_Novo.Text
  If IsNull(Aux) Then Exit Sub
  If Aux = "" Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux < 1) Then Exit Sub
  If Val(Aux > 9999) Then Exit Sub
  
  rsFuncionarios.FindFirst "Código = " & CInt(Aux)
  If rsFuncionarios.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(242)
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  Nome_Novo.Caption = rsFuncionarios("Nome")
 
End Sub

Private Sub Form_Load()
  gbCopyPermissoes = False
  Set rsFuncionarios = db.OpenRecordset("SELECT * FROM Funcionários ORDER BY Código", dbOpenDynaset)
  Data1.DatabaseName = gsQuickDBFileName
  Cod_Func.Caption = gsCodigoFrom
  rsFuncionarios.FindFirst "Código = " & gsCodigoFrom
  Nome_func.Caption = rsFuncionarios("Nome")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsFuncionarios.Close
  Set rsFuncionarios = Nothing
  Call StatusMsg("")
End Sub
