VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendasPorCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas por cliente"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmRelVendasPorCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   6015
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cli_For"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Data datFilial 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
      Begin VB.OptionButton B_Vídeo 
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
         Left            =   120
         TabIndex        =   9
         Top             =   320
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton B_Impressora 
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
         Left            =   960
         TabIndex        =   10
         Top             =   320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Top             =   3120
      Width           =   3375
      Begin VB.OptionButton O_Normal 
         Caption         =   "Normal"
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
         Left            =   120
         TabIndex        =   12
         Top             =   320
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Grade 
         Caption         =   "Grade"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   320
         Width           =   855
      End
      Begin VB.OptionButton O_Edição 
         Caption         =   "Edição"
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
         Left            =   2400
         TabIndex        =   14
         Top             =   320
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   5775
      Begin VB.OptionButton O_Data 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton O_Código 
         Caption         =   "Código do Produto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1680
      End
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome Produto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   5775
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data_Ini 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
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
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.TextBox txtFilial 
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   360
      Width           =   4335
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "frmRelVendasPorCliente.frx":058A
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
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
      Columns(0).Width=   4022
      Columns(0).Caption=   "Filial"
      Columns(0).Name =   "Filial"
      Columns(0).DataField=   "Filial"
      Columns(0).FieldLen=   256
      Columns(1).Width=   5450
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
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
      DataFieldToDisplay=   "Filial"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Fornecedor 
      Bindings        =   "frmRelVendasPorCliente.frx":05A2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
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
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1984
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2355
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
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   3240
      Top             =   4680
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
   Begin VB.Label Nome_Fornecedor 
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
      Left            =   1560
      TabIndex        =   22
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Cliente / Fornecedor :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Filial:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmRelVendasPorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Imprime_Click()
  Dim rstVendasClientes As Recordset
  Dim rstVendas         As Recordset
  Dim rstVendasProdutos As Recordset
  Dim rstRelVendas      As Recordset
  
  Dim intTamanho        As Integer
  Dim intCor            As Integer
  Dim strSQL            As String
  
  Dim dblDescontoSubTotal As Double
  
  If Not IsDate(Data_Ini.Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    MsgBox "Data final inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  dbTemp.Execute " DELETE * FROM tblRelVendasCliente "
  '26/10/2004 - Daniel
  'Criado uma segunda table temporária chamada tblRelVendasCliente2
  'nela colocaremos as informações iniciais e em seguida a partir do
  'group by dela criaremos os registros já agrupados em tblRelVendasCliente
  'devido a solicitações de clientes
  dbTemp.Execute " DELETE * FROM tblRelVendasCliente2 "
  
  '--------------------------[Instruções]--------------------------
  strSQL = " SELECT DISTINCT Cliente FROM Saídas, [Operações Saída] WHERE Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
  strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
  strSQL = strSQL & " AND Saídas.Operação = [Operações Saída].Código AND [Operações Saída].Tipo = 'V' AND Saídas.Efetivada = TRUE "
  '20/05/2004 - Daniel
  'Adicionado mais um AND Saídas.[Nota Cancelada] = FALSE
  'para evitar de trazer valores de saídas canceladas após impressão de notas
  strSQL = strSQL & " AND Saídas.[Nota Cancelada] = FALSE "
  
  If Len(Trim(Nome_Fornecedor.Caption)) > 0 Then
    strSQL = strSQL & " AND Cliente = " & Combo_Fornecedor.Text
  End If
  
  If IsNumeric(cboFilial.Text) Then
    strSQL = strSQL & " AND Filial = " & cboFilial.Text
  End If
  
  Set rstVendasClientes = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstVendasClientes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      Do While Not .EOF
        strSQL = " SELECT * FROM Saídas, [Operações Saída] WHERE Cliente = " & .Fields("Cliente") & " AND Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
        strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
        strSQL = strSQL & " AND ( Saídas.Operação = [Operações Saída].Código ) AND [Operações Saída].Tipo = 'V' AND Saídas.Efetivada = TRUE"
        '20/05/2004 - Daniel
        'Adicionado mais um AND Saídas.[Nota Cancelada] = FALSE
        'para evitar de trazer valores de saídas canceladas após impressão de notas
        strSQL = strSQL & " AND Saídas.[Nota Cancelada] = FALSE "
        
        Set rstVendas = db.OpenRecordset(strSQL, dbOpenDynaset)
        
        With rstVendas
          If Not (.BOF And .EOF) Then
            .MoveFirst
            
            
            Do While Not .EOF
              strSQL = " SELECT * FROM [Saídas - Produtos], Produtos "
              strSQL = strSQL & " WHERE Filial = " & .Fields("Filial")
              strSQL = strSQL & " AND Sequência = " & .Fields("Sequência") & " AND [Saídas - Produtos].[Código Sem Grade] = Produtos.Código "
              
              Set rstVendasProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
              
              With rstVendasProdutos
                If Not (.BOF And .EOF) Then
                  .MoveFirst
                  
                  Do While Not .EOF
                    '---[ Adiciona os produtos no relatório de vendas por cliente ]---'
                      strSQL = " SELECT * FROM tblRelVendasCliente2 WHERE Filial = " & rstVendas.Fields("Filial")
                      strSQL = strSQL & " AND Data = #" & rstVendas.Fields("Data") & "# "
                      
                      intTamanho = 0
                      intCor = 0
                      
                      If O_Normal.Value Then
                        strSQL = strSQL & " AND Produto = '" & rstVendasProdutos.Fields("[Código Sem Grade]").Value & "' AND Tamanho = 0 AND Cor = 0 AND Edicao = 0"
                      End If
                      
                      If O_Grade.Value Then
                        If rstVendasProdutos.Fields("Tipo") = "G" Then
                          intTamanho = Left(Right(rstVendasProdutos.Fields("Saídas - Produtos.Código"), 6), 3)
                          intCor = Right(rstVendasProdutos.Fields("Saídas - Produtos.Código"), 3)
                          
                          strSQL = strSQL & " AND Produto = '" & rstVendasProdutos.Fields("Código Sem Grade").Value & "' "
                          strSQL = strSQL & " AND Tamanho = " & intTamanho & " AND Cor = " & intCor & " AND Edicao = 0"
                        End If
                      End If
                      
                      Set rstRelVendas = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
                      
                      rstRelVendas.AddNew
                      rstRelVendas.Fields("Filial") = .Fields("Filial")
                      rstRelVendas.Fields("Data") = rstVendas.Fields("Data")
                      rstRelVendas.Fields("Cliente") = rstVendas.Fields("Cliente")
                      rstRelVendas.Fields("Produto") = rstVendasProdutos.Fields("Produtos.Código")
                      rstRelVendas.Fields("Tamanho") = intTamanho
                      rstRelVendas.Fields("Cor") = intCor
                      rstRelVendas.Fields("Edicao") = 0
                      rstRelVendas.Fields("QtdeVendida") = rstRelVendas.Fields("QtdeVendida") + rstVendasProdutos.Fields("Qtde")
                      '05/07/2004 - Daniel
                      'Old: + rstVendasProdutos.Fields("Preço Final")
                      'Alterado para Apresentar o total da venda, caso houvesse desconto não estava mostrando
                      'o valor correto
                      rstRelVendas.Fields("ValorVendido") = rstRelVendas.Fields("ValorVendido") + rstVendasProdutos.Fields("Preço Final").Value '(Testar) menos rstVendas.Fields("DescontoSubTotal").Value
                      rstRelVendas.Update
                    '---[ Adiciona os produtos no relatório de vendas por cliente ]---'
                    .MoveNext
                  Loop
                End If
                
                .Close
                Set rstVendasProdutos = Nothing
              End With
              .MoveNext
            Loop
          End If
          
          .Close
          Set rstVendas = Nothing
        End With
        
        .MoveNext
      Loop
    End If
    
    .Close
    Set rstVendasClientes = Nothing
  End With
  
  '26/10/2004 - Daniel
  'Criado uma segunda table temporária chamada tblRelVendasCliente2
  'nela colocaremos as informações iniciais e em seguida a partir do
  'group by dela criaremos os registros já agrupados em tblRelVendasCliente
  'devido a solicitações de clientes
  Call AgruparValores
  
  With Rel
    .Reset
    
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsQuickDBFileName
    
    If O_Código.Value Then
      .SortFields(0) = "+{Produtos.Código Ordenação}"
    End If
    
    If O_Data.Value Then
      .SortFields(0) = "+{tblRelVendasCliente.Data}"
    End If
    
    If O_Nome.Value Then
      .SortFields(0) = "+{Produtos.Nome}"
    End If
    
    If O_Normal.Value Then
      .ReportFileName = gsReportPath & "rptVendasClienteNormal.rpt"
    End If
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
    
    If O_Grade.Value Then
      .DataFiles(5) = gsQuickDBFileName
      .DataFiles(6) = gsQuickDBFileName
      .ReportFileName = gsReportPath & "rptVendasClienteGrade.rpt"
    End If
    
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
  End With
End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
End Sub

Private Sub cboFilial_LostFocus()
  Dim rsFilial As Recordset
  
  txtFilial.Text = ""
  
  With cboFilial
    If Not IsNumeric(.Text) Then Exit Sub
    
    Set rsFilial = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
    
    If Not (rsFilial.BOF And rsFilial.EOF) Then
      txtFilial.Text = rsFilial.Fields("Nome") & ""
    End If
    
    rsFilial.Close
    Set rsFilial = Nothing
  End With
End Sub

Private Sub Combo_Fornecedor_CloseUp()
  Combo_Fornecedor.Text = Combo_Fornecedor.Columns(1).Text
  Combo_Fornecedor_LostFocus
End Sub

Private Sub Combo_Fornecedor_LostFocus()
  Dim rstFornecedores As Recordset
  
  Nome_Fornecedor.Caption = ""

  If Not IsNumeric(Combo_Fornecedor.Text) Then Exit Sub
  
  Set rstFornecedores = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código = " & Combo_Fornecedor.Text, dbOpenDynaset)
  
  With rstFornecedores
    If Not (.BOF And .EOF) Then
      Nome_Fornecedor.Caption = .Fields("Nome") & ""
    End If
    
    .Close
    Set rstFornecedores = Nothing
  End With
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End If
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End If
End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Data1.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
End Sub

Private Sub AgruparValores()
  Dim rstTemp           As Recordset
  Dim rstVendasClientes As Recordset
  Dim strQuery          As String
  
  strQuery = "SELECT Cliente, Filial, Data, Produto, Tamanho, Cor, Edicao, Sum(QtdeVendida) AS TotQtdeVendida, Sum(ValorVendido) AS Total "
  strQuery = strQuery & " FROM tblRelVendasCliente2 "
  strQuery = strQuery & " GROUP BY Cliente, Filial, Data, Produto, Tamanho, Cor, Edicao "
  
  Set rstTemp = dbTemp.OpenRecordset(strQuery, dbOpenDynaset)
  Set rstVendasClientes = dbTemp.OpenRecordset("tblRelVendasCliente", dbOpenDynaset)
  
  With rstTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
          rstVendasClientes.AddNew
           rstVendasClientes.Fields("Cliente").Value = .Fields("Cliente").Value
           rstVendasClientes.Fields("Filial").Value = .Fields("Filial").Value
           rstVendasClientes.Fields("Data").Value = .Fields("Data").Value
           rstVendasClientes.Fields("Produto").Value = .Fields("Produto").Value
           rstVendasClientes.Fields("Tamanho").Value = .Fields("Tamanho").Value
           rstVendasClientes.Fields("Cor").Value = .Fields("Cor").Value
           rstVendasClientes.Fields("Edicao").Value = .Fields("Edicao").Value
           rstVendasClientes.Fields("QtdeVendida").Value = .Fields("TotQtdeVendida").Value
           rstVendasClientes.Fields("ValorVendido").Value = .Fields("Total").Value
          rstVendasClientes.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstTemp = Nothing
  
  rstVendasClientes.Close
  Set rstVendasClientes = Nothing
  
End Sub
