VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de Clientes / Produtos"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   6525
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Importar"
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
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5265
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1080
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   6495
      Begin VB.CommandButton cmdCaminhoOrigem 
         Caption         =   "..."
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   555
         Width           =   435
      End
      Begin VB.TextBox txtBaseDadosOrigem 
         Enabled         =   0   'False
         Height          =   315
         Left            =   645
         TabIndex        =   3
         Top             =   540
         Width           =   5115
      End
      Begin MSComDlg.CommonDialog cdgOrigem 
         Left            =   5880
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Base de Dados de Origem"
         Height          =   225
         Left            =   645
         TabIndex        =   8
         Top             =   285
         Width           =   2670
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Importação"
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
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmImportacao.frx":058A
         ForeColor       =   &H00808080&
         Height          =   735
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strPathBaseOrigem  As String
Private m_strSenha           As String

Private Sub cmdCaminhoOrigem_Click()

On Error GoTo ErrHandler

  With cdgOrigem
    .CancelError = True
    .DialogTitle = "Localize a base de origem."
    .Filter = "Arquivos de dados (*.mdb)|*.mdb|Todos os arquivos (*.*)|*.*"
    .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    .ShowOpen
    m_strPathBaseOrigem = .FileName
  
  End With
  
  txtBaseDadosOrigem.Text = m_strPathBaseOrigem
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  If Err.Number <> cdlCancel Then
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End If

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImportar_Click()

  If Len(txtBaseDadosOrigem.Text) <= 0 Then
    MsgBox "Informe o caminho da base de origem:", vbExclamation, "Quick Store"
    Exit Sub
  End If

  m_strSenha = InputBox("Entre com a senha para a importação.", "Senha")
  If m_strSenha <> "789275" Then
    DisplayMsg "Senha incorreta, importação não realizada."
    Exit Sub
  End If

  Call Importacao

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

Private Sub Importacao()
  Dim dbOrigem  As Database

  Dim rstCliQS      As Recordset
  Dim rstProdQS     As Recordset
  Dim rstCliOrigem  As Recordset
  Dim rstProdOrigem As Recordset
  
  Dim blnInTransaction As Boolean
  
  On Error GoTo ErrHandler
  
  'Inicia transação
  ws.BeginTrans
  blnInTransaction = True
  
  Set dbOrigem = ws.OpenDatabase(m_strPathBaseOrigem, False, False)
  
  Set rstCliOrigem = dbOrigem.OpenRecordset("SELECT * FROM CatFicha ORDER BY Código", dbOpenSnapshot)
  
  Set rstCliQS = db.OpenRecordset("Cli_For", dbOpenDynaset)
  
  Call StatusMsg("Importando os clientes...")
  
  With rstCliOrigem
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        rstCliQS.AddNew
          rstCliQS.Fields("Código").Value = .Fields("Código").Value
          rstCliQS.Fields("Nome").Value = .Fields("Nome").Value & ""
          rstCliQS.Fields("CGC").Value = .Fields("Cpf_cgc").Value & ""
          rstCliQS.Fields("Inscrição").Value = .Fields("Insc_Rg").Value & ""
          rstCliQS.Fields("Endereço").Value = .Fields("Endereço").Value & ""
          rstCliQS.Fields("Bairro").Value = .Fields("Bairro").Value & ""
          rstCliQS.Fields("Cidade").Value = .Fields("Cidade").Value & ""
          rstCliQS.Fields("CEP").Value = Replace(.Fields("Cep").Value & "", "_", "")
          rstCliQS.Fields("Fone 1").Value = Replace(.Fields("Telefone1").Value & "", "_", "")
          rstCliQS.Fields("Fone 2").Value = Replace(.Fields("Telefone2").Value & "", "_", "")
          If Len(Replace(.Fields("Fax").Value & "", "_", "")) > 15 Then
            rstCliQS.Fields("Fax").Value = Mid((Replace(.Fields("Fax").Value & "", "_", "")), 1, 15)
          Else
            rstCliQS.Fields("Fax").Value = Replace(.Fields("Fax").Value & "", "_", "")
          End If
          rstCliQS.Fields("email").Value = .Fields("Email").Value & ""
          rstCliQS.Fields("Física_Jurídica").Value = .Fields("Pessoa").Value & ""
          If (.Fields("ClienFornec").Value) = "A" Or (.Fields("ClienFornec").Value) = "C" Then
            rstCliQS.Fields("Tipo").Value = "C"
          Else
            rstCliQS.Fields("Tipo").Value = "F"
          End If
          rstCliQS.Fields("Estado").Value = .Fields("Estado").Value & ""
        rstCliQS.Update
        
        .MoveNext
      Loop
      
    End If
    .Close
  End With
    
  Set rstCliOrigem = Nothing
  
  rstCliQS.Close
  Set rstCliQS = Nothing
  
  'Tratamento para os Produtos
  Set rstProdOrigem = dbOrigem.OpenRecordset("SELECT * FROM CatItem ORDER BY Coditem", dbOpenSnapshot)
  
  Set rstProdQS = db.OpenRecordset("Produtos", dbOpenDynaset)
  
  Call StatusMsg("Importando os produtos...")
  
  With rstProdOrigem
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        rstProdQS.AddNew
          rstProdQS.Fields("Código").Value = .Fields("Coditem").Value & ""
          rstProdQS.Fields("Código Ordenação").Value = Right(("++++++++++++++++++++" & .Fields("Coditem").Value & ""), 20)
          rstProdQS.Fields("Tipo").Value = "N"
          rstProdQS.Fields("Nome").Value = .Fields("Descrição").Value & ""
          rstProdQS.Fields("Unidade Venda").Value = .Fields("Unidade").Value & ""
          rstProdQS.Fields("Moeda").Value = 1
          rstProdQS.Fields("Classe").Value = 9999
          rstProdQS.Fields("Sub Classe").Value = 9999
        rstProdQS.Update
        
        .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstProdOrigem = Nothing
  
  rstProdQS.Close
  Set rstProdQS = Nothing
  
  'Fechamos a database de origem
  dbOrigem.Close
  Set dbOrigem = Nothing
  
  'Finaliza transação
  ws.CommitTrans
  blnInTransaction = False
  
  Call StatusMsg("")
  
  MsgBox "Importação de Clientes e Produtos realizada com sucesso.", vbInformation, "Quick Store"
  
  Exit Sub
  
ErrHandler:
  'Desfaz transação
  If blnInTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
    

End Sub
