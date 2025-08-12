VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmEmissaoCertificado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificados"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmissaoCertificado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmEmissaoCertificado.frx":05CA
   ScaleHeight     =   3765
   ScaleWidth      =   7125
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
      TabIndex        =   17
      Top             =   960
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
      Begin VB.Label Label1 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Fim"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   1215
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
      TabIndex        =   14
      Top             =   2040
      Width           =   3375
      Begin VB.TextBox txtNotaInicio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNotaFim 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fim"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   1215
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
      TabIndex        =   12
      Top             =   2040
      Width           =   3375
      Begin VB.TextBox txtSequenciaInicio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtSequenciaFim 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Fim"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   -120
      TabIndex        =   9
      Top             =   -240
      Width           =   8175
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão de certificados"
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
         TabIndex        =   11
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Use os campos abaixo como filtro para os certificados a serem emitidas pelo sistema."
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Data datFiliais 
      Caption         =   "datFiliais"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
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
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin SSDataWidgets_B.SSDBCombo cboCodigoFilial 
      Bindings        =   "frmEmissaoCertificado.frx":0B54
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
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
      Columns(0).Width=   3200
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Filial"
   End
   Begin Crystal.CrystalReport crtView 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label8 
      Caption         =   "Filial"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   960
      Width           =   1215
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
      TabIndex        =   20
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmEmissaoCertificado"
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
  Dim strSelection      As String
  '03/03/2004 - Daniel
  'Manutenção para imprimir conforme a Qtde.
  'proveniente da tabela [Saídas - Produtos]
  Dim strSQL                As String
  Dim rstCertificados       As Recordset
  Dim rsttblRelCertificados As Recordset
  Dim sngContQtde           As Single
  Dim intAuxi               As Integer
  Dim intContador           As Integer
  
  'Limpando a tabela temporária tblRelCertificados
  dbTemp.Execute "DELETE * FROM tblRelCertificados"
  
  strSQL = " SELECT [Saídas - Produtos].Filial, [Saídas - Produtos].Sequência, [Saídas - Produtos].Qtde, Produtos.Código AS CodProd, Produtos.Nome AS NomeProd, Produtos.Obs, Saídas.Data, Saídas.[Nota Impressa], Cli_For.Nome AS NomeCliente "
  strSQL = strSQL & " FROM Saídas, [Saídas - Produtos], Produtos, Cli_For "
  strSQL = strSQL & " WHERE Saídas.Cliente <> 0 "
  
  If Not IsDate(mskDataInicio.Text) Then
    MsgBox "Data Início Inválida", vbExclamation, "Quick Store"
    mskDataInicio.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskDataFim.Text) Then
    MsgBox "Data Final Inválida", vbExclamation, "Quick Store"
    mskDataFim.SetFocus
    Exit Sub
  End If
  
  If CDate(mskDataFim.Text) < CDate(mskDataInicio.Text) Then
    MsgBox "Data Final Menor que a Inicial", vbExclamation, "Quick Store"
    mskDataFim.SetFocus
    Exit Sub
  End If
  
  strSQL = strSQL & " AND Saídas.Data >=#" & (Format((mskDataInicio.Text), "yyyy/mm/dd")) & "#"
  strSQL = strSQL & " AND Saídas.Data <=#" & (Format((mskDataFim.Text), "yyyy/mm/dd")) & "#"
  
  If IsNumeric(txtNotaInicio.Text) Then
    strSQL = strSQL & " AND Saídas.[Nota Impressa] >=" & CLng(txtNotaInicio.Text)
  End If
  
  If IsNumeric(txtNotaFim.Text) Then
    strSQL = strSQL & " AND Saídas.[Nota Impressa] <=" & CLng(txtNotaFim.Text)
  End If
  
  If IsNumeric(txtSequenciaInicio.Text) Then
    strSQL = strSQL & " AND Saídas.Sequência >=" & CLng(txtSequenciaInicio.Text)
  End If
  
  If IsNumeric(txtSequenciaFim.Text) Then
    strSQL = strSQL & " AND Saídas.Sequência <=" & CLng(txtSequenciaFim.Text)
  End If
  
  strSQL = strSQL & " AND [Saídas - Produtos].Filial = Saídas.Filial "
  strSQL = strSQL & " AND [Saídas - Produtos].Sequência = Saídas.Sequência "
  strSQL = strSQL & " AND Produtos.Código = [Saídas - Produtos].[Código sem Grade] "
  
  strSQL = strSQL & " AND Cli_For.Código = Saídas.Cliente "
  
  
  '-----[Trabalhando com as tables]-----
  Set rsttblRelCertificados = dbTemp.OpenRecordset(" SELECT * FROM tblRelCertificados ", dbOpenDynaset)
  
  Set rstCertificados = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstCertificados
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      
      Do Until .EOF
        sngContQtde = .Fields("Qtde").Value
        
          For intAuxi = 1 To sngContQtde
          
              intContador = intContador + 1
              
              With rsttblRelCertificados
                .AddNew
                
                .Fields("Nome do Produto").Value = rstCertificados.Fields("NomeProd").Value
                .Fields("Obs do Produto").Value = rstCertificados.Fields("Obs").Value
                .Fields("Nome do Cliente").Value = rstCertificados.Fields("NomeCliente").Value
                .Fields("Data da Saída").Value = rstCertificados.Fields("Data").Value
                .Fields("Código do Produto").Value = rstCertificados.Fields("CodProd").Value
                .Fields("Nota Fiscal").Value = rstCertificados.Fields("Nota Impressa").Value
                .Fields("Contador").Value = intContador
                
                .Update
              End With
          Next intAuxi
          
      .MoveNext
      Loop
    End If
  End With

  rstCertificados.Close
  Set rstCertificados = Nothing
  rsttblRelCertificados.Close
  Set rsttblRelCertificados = Nothing
  '-----------------------------------
  
  With crtView
    .Reset
    .ReportFileName = gsReportPath & "rptCertificado.rpt"
    .DataFiles(0) = gsTempDBFileName 'Fará a busca de apenas uma tabela temporária
    
    strSelection = " {tblRelCertificados.Contador} <> 0 "
    
    .SelectionFormula = strSelection
    .WindowState = crptMaximized
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
