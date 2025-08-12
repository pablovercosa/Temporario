VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPesquisaClientesProduto 
   Caption         =   "Clientes Por Produto"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "PesquisaClientesProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   5625
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   16457
            Text            =   "0 registros encontrados"
            TextSave        =   "0 registros encontrados"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnExportarPlanilha 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exportar Planilha"
      Height          =   465
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1830
      Width           =   4755
   End
   Begin VB.CommandButton btnPesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
      Height          =   465
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1830
      Width           =   4755
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultados 
      Height          =   3285
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Selecione a linha e dê duplo-clique para posicionamento."
      Top             =   2340
      Width           =   9615
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      ForeColorEven   =   4210752
      BackColorOdd    =   16777152
      RowHeight       =   370
      ExtraHeight     =   79
      Columns.Count   =   3
      Columns(0).Width=   7064
      Columns(0).Caption=   "Produto"
      Columns(0).Name =   "Descricao"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5292
      Columns(1).Caption=   "Cliente"
      Columns(1).Name =   "Descricao_2"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3519
      Columns(2).Caption=   "Telefone"
      Columns(2).Name =   "Descricao_3"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   16960
      _ExtentY        =   5794
      _StockProps     =   79
      BackColor       =   15066597
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Frame fraFiltro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filtros (opcional)"
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   30
      Width           =   9615
      Begin VB.TextBox txtCliente3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   6090
         TabIndex        =   11
         Top             =   1290
         Width           =   3345
      End
      Begin VB.TextBox txtCliente2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   6090
         TabIndex        =   10
         Top             =   870
         Width           =   3345
      End
      Begin VB.TextBox txtCliente1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   6090
         TabIndex        =   9
         Text            =   "CONSUMIDOR"
         Top             =   450
         Width           =   3345
      End
      Begin VB.TextBox txtProduto2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   150
         TabIndex        =   7
         Top             =   1170
         Width           =   5115
      End
      Begin VB.TextBox txtProduto1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   570
         Width           =   5115
      End
      Begin VB.Label lblClientes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes Que Não Serão Mostados na Pesquisa"
         Height          =   225
         Left            =   6120
         TabIndex        =   8
         Top             =   180
         Width           =   3405
      End
      Begin VB.Label lblProduto2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parte da Decrição do Produto (2)"
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   960
         Width           =   2475
      End
      Begin VB.Label lblProduto1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parte da Decrição do Produto (1)"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmPesquisaClientesProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As Recordset

Private Const QUERY As String = "SELECT DISTINCT p.[Nome] AS Produto, c.[Nome] AS Cliente, c.[Fone 1] AS Telefone" & _
                                " FROM (([Saídas - Produtos] AS sp" & _
                                " INNER JOIN [Produtos] AS p ON p.[Código] = sp.[Código])" & _
                                " INNER JOIN [Saídas] AS s ON s.[Filial] = sp.[Filial] AND s.[Sequência] = sp.[Sequência])" & _
                                " INNER JOIN [Cli_For] AS c ON c.[Código] = s.[Cliente]" & _
                                " WHERE c.[Tipo] = 'C' AND c.[Inativo] = FALSE AND LTRIM(RTRIM(p.[Nome])) <> ''" & _
                                "<CLIENTES_IGNORADOS>" & _
                                "<PRODUTOS>" & _
                                " ORDER BY p.[Nome] ASC, c.[Nome] ASC;"

Private Const QUERY_CLIENTES_IGNORADOS As String = " AND UCASE(c.[Nome]) NOT LIKE UCASE('*<CLIENTE>*')"
Private Const QUERY_PRODUTOS As String = " AND UCASE(p.[Nome]) LIKE UCASE('*<PRODUTO>*')"

'Private Const CAMINHO As String = "C:\Documents and Settings\Administrador\Desktop\"
Private Const ARQUIVO As String = "ClientesPorProduto_<SEQUENCIA>.xls"
Private Const CAMINHO As String = "\\tsclient\c\QuickStore\Planilhas\"

Private Sub PersistirDiretorio()
  Dim pasta As Variant
  Dim pastas() As String
  Dim raiz As String
  raiz = "\\tsclient\c"
  'raiz = "C:\Documents and Settings"
  pastas = Split(Replace(CAMINHO, raiz, ""), "\")
  
  For Each pasta In pastas
    If Trim(CStr(pasta)) <> "" Then
        raiz = raiz & "\" & CStr(pasta)
        If Len(Dir(raiz, vbDirectory) & "") = 0 Then MkDir raiz
    End If
  Next pasta
End Sub

Private Sub btnExportarPlanilha_Click()
On Error GoTo Erro
    If Not (rs Is Nothing) Then
        If (rs.RecordCount > 0) Then
            Dim planilha As String
            planilha = CAMINHO & Replace(ARQUIVO, "<SEQUENCIA>", Format(DateTime.Now, "yyyyMMddHHmmsss"))
            
            Me.btnExportarPlanilha.Enabled = False
            Me.btnPesquisar.Enabled = False
            
            Call PreencherExcel(planilha)

            MsgBox "O arquivo foi salvo no caminho " & planilha, vbOKOnly, "Informação"
        Else
            MsgBox "Sem dados para exportar", vbOKOnly, "Atenção"
        End If
    Else
        MsgBox "Sem dados para exportar", vbOKOnly, "Atenção"
    End If
    Me.btnExportarPlanilha.Enabled = True
    Me.btnPesquisar.Enabled = True
    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
    Me.btnExportarPlanilha.Enabled = True
    Me.btnPesquisar.Enabled = True
End Sub

Private Sub btnPesquisar_Click()
    Call Limpar
    Call BuscaDados
    Call PreencheGrid
End Sub

Private Sub Limpar()
    If Not (rs Is Nothing) Then rs.Close
    Set rs = Nothing
    
    grdResultados.RemoveAll
    Status.Panels.Item(1).Text = "0 registros encontrados"
End Sub

Private Sub BuscaDados()
On Error GoTo Erro
    Dim strClientesIgnorados As String
    Dim strProduto As String
    Dim srtQuery As String
    
    strClientesIgnorados = ""
    If Trim(Me.txtCliente1.Text) <> "" Then strClientesIgnorados = strClientesIgnorados & Replace(QUERY_CLIENTES_IGNORADOS, "<CLIENTE>", Replace(Trim(Me.txtCliente1.Text), " ", "*"))
    If Trim(Me.txtCliente2.Text) <> "" Then strClientesIgnorados = strClientesIgnorados & Replace(QUERY_CLIENTES_IGNORADOS, "<CLIENTE>", Replace(Trim(Me.txtCliente2.Text), " ", "*"))
    If Trim(Me.txtCliente3.Text) <> "" Then strClientesIgnorados = strClientesIgnorados & Replace(QUERY_CLIENTES_IGNORADOS, "<CLIENTE>", Replace(Trim(Me.txtCliente3.Text), " ", "*"))

    strProduto = ""
    If Trim(Me.txtProduto1) <> "" Then strProduto = strProduto & Replace(QUERY_PRODUTOS, "<PRODUTO>", Replace(Trim(Me.txtProduto1.Text), " ", "*"))
    If Trim(Me.txtProduto2) <> "" Then strProduto = strProduto & Replace(QUERY_PRODUTOS, "<PRODUTO>", Replace(Trim(Me.txtProduto2.Text), " ", "*"))
    
    srtQuery = Replace(QUERY, "<CLIENTES_IGNORADOS>", strClientesIgnorados)
    srtQuery = Replace(srtQuery, "<PRODUTOS>", strProduto)

    Set rs = db.OpenRecordset(srtQuery, dbOpenDynaset, dbReadOnly)
    
    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub PreencheGrid()
On Error GoTo Erro
    If rs.RecordCount = 0 Then Exit Sub
    If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
  
    grdResultados.Redraw = False
    While Not rs.EOF
        grdResultados.AddItem rs.Fields("Produto").Value & vbTab & _
            rs.Fields("Cliente").Value & vbTab & _
            rs.Fields("Telefone").Value
        rs.MoveNext
    Wend
    grdResultados.Redraw = True
    Status.Panels.Item(1).Text = CStr(rs.RecordCount) & " registros encontrados"
    
    Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub PreencherExcel(ByVal planilha As String)
On Error GoTo Erro
    If rs.RecordCount = 0 Then Exit Sub
    If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
    
    Dim Excel As New Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    
    Dim F As Field
    Dim K As Integer
    Dim i As Integer

    Set wb = Excel.Workbooks.Add
    Set ws = wb.Worksheets.Add
    ws.Cells.Clear

    K = 1
    For Each F In rs.Fields
        ws.Cells(1, K).Font.Bold = True
        ws.Cells(1, K) = F.Name
        K = K + 1
    Next
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        K = 1
        For Each F In rs.Fields
            If IsDate(rs.Fields(F.Name)) = True Then
                ws.Cells(i + 1, K) = Format(rs.Fields(F.Name))
            Else
                ws.Cells(i + 1, K) = rs.Fields(F.Name)
            End If

            ws.Columns(K).AutoFit
            K = K + 1
        Next
        rs.MoveNext
    Next

    wb.SaveAs planilha

    wb.Close True
    Set Excel = Nothing
    Set wb = Nothing
    Set ws = Nothing
Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_Load()
    Call PersistirDiretorio
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
End Sub
