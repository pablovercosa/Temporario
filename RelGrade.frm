VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório da Grade"
   ClientHeight    =   2070
   ClientLeft      =   1425
   ClientTop       =   2655
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1470
   Icon            =   "RelGrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2070
   ScaleWidth      =   8100
   Begin VB.Data datGrade 
      Caption         =   "Grade"
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
      Left            =   180
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Produtos"
      Top             =   2430
      Visible         =   0   'False
      Width           =   2055
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
      Height          =   400
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   7905
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2445
      Begin VB.OptionButton optPrinter 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1245
      End
      Begin VB.OptionButton optVideo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport rptReport 
      Left            =   7530
      Top             =   1020
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
   Begin SSDataWidgets_B.SSDBCombo cboProduto 
      Bindings        =   "RelGrade.frx":4E95A
      DataSource      =   "datGrade"
      Height          =   315
      Left            =   795
      TabIndex        =   1
      Top             =   180
      Width           =   1785
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
      Columns(0).Width=   8758
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Codigo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      _ExtentX        =   3149
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label lblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2625
      TabIndex        =   2
      Top             =   180
      Width           =   5400
   End
   Begin VB.Label Label1 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmRelGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
  Dim rsGrade As Recordset
  Dim rsTempo As Recordset
  Dim sSql As String
  Dim sCodigo As String
  Dim sTamCor As String
  
  'Apaga pesquisa anterior do arquivo temporario
  Call StatusMsg("Aguarde, preparando arquivo temporário ...")
  db.Execute "Delete * From [Grade - Tempo]"
  Call StatusMsg("")
  
  sSql = "SELECT DISTINCTROW([Códigos da Grade].Código) AS Grade FROM " & _
    "Produtos INNER JOIN [Códigos da Grade] ON Produtos.Código = " & _
    "[Códigos da Grade].[Código Original]"
  If lblNome.Caption <> "" Then
    sSql = sSql & " WHERE Produtos.Código = '" & cboProduto.Text & "'"
  End If
  
  Set rsGrade = db.OpenRecordset(sSql, dbOpenSnapshot)
  Set rsTempo = db.OpenRecordset("Grade - Tempo")
 
  With rsGrade
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do Until .EOF
        sCodigo = IIf(IsNull(!Grade), "", !Grade)
        If sCodigo <> "" Then
          sTamCor = Right(sCodigo, 6)
          sCodigo = Mid(sCodigo, 1, Len(sCodigo) - 6)
          With rsTempo
            .AddNew
            !Produto = sCodigo
            !Tamanho = Val(Left(sTamCor, 3))
            !Cor = Val(Right(sTamCor, 3))
            .Update
          End With
        End If
        .MoveNext
      Loop
    End If
    .Close
    rsTempo.Close
  End With
  Set rsGrade = Nothing
  Set rsTempo = Nothing
  
  With rptReport
    'Nome do BD
    .DataFiles(0) = gsQuickDBFileName
    If optVideo Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    .ReportFileName = gsReportPath & "RELGRADE.RPT"
    'Campo fórmula para o nome da empresa
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    MousePointer = vbHourglass
    Call StatusMsg("Aguarde, imprimindo...")
  
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 rptReport
  
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", rptReport)
  
    
    .Action = 1
  End With
  
  Call StatusMsg("")
  MousePointer = vbDefault

End Sub

Private Sub cboProduto_CloseUp()
  cboProduto.Text = cboProduto.Columns("Codigo").Text
  cboProduto_LostFocus
End Sub

Private Sub cboProduto_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub cboProduto_LostFocus()
  Call StatusMsg("")
  If cboProduto.Text <> "" And cboProduto.Text <> "0" Then
    lblNome.Caption = gsGetNameProduto(cboProduto.Text)
  Else
    lblNome.Caption = ""
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  datGrade.RecordSource = "SELECT Nome, Código FROM Produtos WHERE Código <> '0' AND Desativado = False AND Tipo = 'G'"
  datGrade.DatabaseName = gsQuickDBFileName
End Sub
