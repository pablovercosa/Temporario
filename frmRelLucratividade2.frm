VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelLucratividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Lucratividade"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelLucratividade2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   10710
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5220
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Data datFilial 
      Caption         =   "datFilial"
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
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4650
      Width           =   10515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   855
      Left            =   4860
      TabIndex        =   7
      Top             =   1920
      Width           =   5745
      Begin VB.OptionButton B_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton B_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   4860
      TabIndex        =   10
      Top             =   2820
      Width           =   5745
      Begin VB.OptionButton O_Vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Separado por vendedor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3450
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton O_Classe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Separado por classe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1845
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordem"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   10485
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2370
         TabIndex        =   6
         Top             =   360
         Width           =   1080
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   750
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   2820
      Width           =   4695
      Begin VB.CommandButton cmd_calendarioDtFim 
         Height          =   420
         Left            =   4080
         Picture         =   "frmRelLucratividade2.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   270
         Width           =   465
      End
      Begin VB.CommandButton cmd_calendarioDtIni 
         Height          =   420
         Left            =   1770
         Picture         =   "frmRelLucratividade2.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   270
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   323
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Left            =   600
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   323
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   353
         Width           =   405
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fim"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2490
         TabIndex        =   20
         Top             =   353
         Width           =   285
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Escolha a tabela com o custo desejado"
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   4695
      Begin VB.ComboBox Combo_Preço 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   510
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   3720
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   1695
      Left            =   -120
      TabIndex        =   14
      Top             =   -240
      Width           =   10845
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "frmRelLucratividade2.frx":4FB1E
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Lucratividade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione abaixo a tabela de preços de acordo com o custo que você deseja calcular."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   2160
         TabIndex        =   16
         Top             =   690
         Width           =   6795
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "• Caso você utilize o sistema de desconto no sub-total, o Quick Store contabiliza os descontos dados como desconto financeiro. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Top             =   1110
         Width           =   8025
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "frmRelLucratividade2.frx":51986
      DataSource      =   "datFilial"
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   1500
      Width           =   1215
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9922
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1429
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   10260
      Top             =   5190
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
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   180
      TabIndex        =   23
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1890
      TabIndex        =   22
      Top             =   1500
      Width           =   8730
   End
End
Attribute VB_Name = "frmRelLucratividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dblTotalQuantidadeVendida As Double
Dim dblTotalQuantidadeDevolvida As Double
Dim dblTotalValorVenda As Double

''12/04/2007 - Anderson
'Private Enum ReturnDevolucaoTipo
'  Quantidade = 0
'  Valor = 1
'End Enum

Private Sub B_Imprime_Click()
  
  Dim dblValorTotalDev As Double
  '14/06/2007 - Anderson
  'Descontinuado para atender as novas exigências da Zue
  'Dim dblQuantidadeTotalDev As Double '12/04/2007 - Anderson - Utilizado para obter a quantidade total de produtos devolvidos.
  Dim dblTotalDescSub As Double
  
  If Len(Trim(Nome_Empresa.Caption)) <= 0 Then
    MsgBox "Filial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(Data_Ini.Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    MsgBox "A data inicial não pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  dbTemp.Execute "DELETE * FROM Lucratividade"
  
  '15/06/2007 - Anderson
  'Variáveis para serem contabilizadas no totalizador do relatório
  dblTotalQuantidadeVendida = 0
  dblTotalQuantidadeDevolvida = 0
  dblTotalValorVenda = 0
  
  '14/06/2007 - Anderson
  'Função utilizada para cálcular o total de vendas no relatório de Lucratividade
  Call StatusMsg("Calculando vendas, aguarde . . . ")
  Call CalcularVendas

  Call StatusMsg("Calculando devoluções, aguarde . . . ")
  Call CalcularDevolucao
  
  Call StatusMsg("Analisando os descontos no sub-total e devoluções, aguarde . . . ")
  '---[ Gera o total de Descontos do sub-total ]---'
    dblValorTotalDev = 0
    dblTotalDescSub = 0
    
    ReturnDescontoSubTotal dblTotalDescSub
    '14/06/2007 - Anderson
    'Descontinuado para atender a solicitação da Zue
    '12/04/2007 - Anderson
    ReturnDevolucaoNormal dblValorTotalDev
    ReturnDevolucaoGrade dblValorTotalDev
    'ReturnDevolucaoNormal dblValorTotalDev, Valor
    'ReturnDevolucaoGrade dblValorTotalDev, Valor
  '---[ Gera o total de Descontos do sub-total ]---'

  Call StatusMsg("")

  With Rel
    .WindowShowGroupTree = O_Classe.Value
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
    
    '.Destination = Not B_Vídeo.Value
    '29/11/2004 - Daniel
    'Bug: Não estava localizando a impressora default
    'ao imprimir direto na impressora
    If B_Vídeo.Value Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    
    If O_Normal.Value Then .ReportFileName = gsReportPath & "LUCRA.RPT"
    If O_Classe.Value Then .ReportFileName = gsReportPath & "LUCRA2.RPT"
    '27/06/2007 - Anderson
    'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
    'Solicitante: Zue
    If O_Vendedor.Value Then .ReportFileName = gsReportPath & "LUCRA3.RPT"
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
  
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    .Formulas(1) = "nome_filial = '" & Nome_Empresa.Caption & "'"
    .Formulas(2) = "data_ini = '" & Data_Ini.Text & "'"
    .Formulas(3) = "data_fim = '" & Data_Fim.Text & "'"
    .Formulas(4) = "tipo_rel = '" & Combo_Preço.Text & "'"
    .Formulas(5) = "DescSubTotal = " & Replace(dblTotalDescSub, ",", ".")
    .Formulas(6) = "DevolucoesValor = " & Replace(dblValorTotalDev, ",", ".")
    .Formulas(7) = "TotalQuantidadeVendida = " & Replace(dblTotalQuantidadeVendida, ",", ".")
    .Formulas(8) = "TotalQuantidadeDevolvida = " & Replace(dblTotalQuantidadeDevolvida, ",", ".")
    .Formulas(9) = "TotalValorVenda = " & Replace(dblTotalValorVenda, ",", ".")
    
    If O_Código.Value Then
      If O_Normal.Value Then
        Rel.SortFields(0) = "+{Lucratividade.Código Ordenação}"
        Rel.SortFields(1) = ""
        Rel.SortFields(2) = ""
      End If
      If O_Classe.Value Then
        Rel.SortFields(0) = "+{Lucratividade.Classe}"
        Rel.SortFields(1) = "+{Lucratividade.Sub Classe}"
        Rel.SortFields(2) = "+{Lucratividade.Código Ordenação}"
      End If
    End If
    
    If O_Nome.Value Then
      If O_Normal.Value Then
        Rel.SortFields(0) = "+{Lucratividade.Nome}"
        Rel.SortFields(1) = ""
        Rel.SortFields(2) = ""
      End If
      If O_Classe.Value Then
        Rel.SortFields(0) = "+{Lucratividade.Classe}"
        Rel.SortFields(1) = "+{Lucratividade.Sub Classe}"
        Rel.SortFields(2) = "+{Lucratividade.Nome}"
      End If
    End If
    
    Call StatusMsg("Aguarde, imprimindo...")
    
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("REL", Rel)
    
    .Action = 1
    
    Call StatusMsg("")
    '29/11/2004 - Daniel
    pgbProgress.Value = 0
    
    MousePointer = vbDefault
  End With
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Nome_Empresa.Caption = ""
  If Not IsNumeric(Combo.Text) Then Exit Sub
  
  With datFilial.Recordset
    .FindFirst "Filial = " & Combo.Text
    
    If Not .NoMatch Then
      Nome_Empresa.Caption = .Fields("Nome") & ""
    End If
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
  Dim rstTabelasPreco As Recordset
  
  Call CenterForm(Me)
  
  datFilial.DatabaseName = gsQuickDBFileName
  
  '---[ Preenchimento da combo de preços ]---'
  Set rstTabelasPreco = db.OpenRecordset("SELECT DISTINCTROW Tabela FROM [Tabela de Preços] ORDER BY Tabela", dbOpenSnapshot)
  
  With rstTabelasPreco
    Combo_Preço.Clear
    If Not (.BOF And .EOF) Then .MoveFirst
    
    Do Until .EOF
      Combo_Preço.AddItem .Fields("Tabela") & ""
      .MoveNext
    Loop
    
    .Close
    Set rstTabelasPreco = Nothing
  End With
  '---[ Preenchimento da combo de preços ]---'
End Sub

Private Function ReturnDevolucaoGrade(ByRef dblValorDevolucao As Double) As Boolean
  Dim strSQL As String
  Dim rstDev As Recordset
  Dim blnProdutoOK As Boolean

  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequência = [Entradas - Produtos].Sequência)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN [Códigos da Grade] ON [Entradas - Produtos].Código = [Códigos da Grade].Código) INNER JOIN Produtos ON [Códigos da Grade].[Código Original] = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Operações Entrada].Tipo)='D')) "


  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "

  If Len(Trim(Nome_Empresa.Caption)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & Combo.Text & ") "
  End If

  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)

  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do While Not .EOF
      
        dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("PrecoTotal"))
        
        .MoveNext

      Loop
    End If
  End With
End Function

Private Function ReturnDevolucaoNormal(ByRef dblValorDevolucao As Double) As Boolean
  Dim strSQL As String
  Dim rstDev As Recordset
  Dim blnProdutoOK As Boolean

  Dim rstProdutos As Recordset
  Dim rstGrade As Recordset

  Dim strCodigoProduto As String

  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequência = [Entradas - Produtos].Sequência) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN Produtos ON [Entradas - Produtos].Código = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Operações Entrada].Tipo)='D')) "

  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "

  If Len(Trim(Nome_Empresa.Caption)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & Combo.Text & ") "
  End If

  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)

  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do While Not .EOF

          dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("PrecoTotal"))

        .MoveNext
      Loop
    End If
  End With
End Function

Private Function ReturnDescontoSubTotal(ByRef dblValorDesconto As Double) As Double
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstVendas         As Recordset
  Dim rstProdutos       As Recordset
  Dim rstDescontoSubTotal As Recordset
  
  Dim dblDescontoSubTotal As Double
  Dim dblDescontoSomar  As Double
  Dim blnProdutoOK      As Boolean
  
  strSQL = " SELECT SUM(Saídas.DescontoSubTotal) AS DescontoSubTotal, [Saídas - Produtos].[Código sem Grade], Saídas.Filial, Saídas.Sequência "
  strSQL = strSQL & " FROM ((Saídas INNER JOIN [Saídas - Produtos] ON (Saídas.Sequência = [Saídas - Produtos].Sequência) AND (Saídas.Filial = [Saídas - Produtos].Filial)) INNER JOIN Produtos ON [Saídas - Produtos].[Código sem Grade] = Produtos.Código) INNER JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
  strSQL = strSQL & " GROUP BY Saídas.Filial, Saídas.Data, Saídas.Cliente, [Saídas - Produtos].[Código sem Grade], Saídas.Digitador, Produtos.Classe, Produtos.[Sub Classe], Saídas.Efetivada, Saídas.[Nota Cancelada], [Operações Saída].Tipo = 'V', Saídas.Sequência, Saídas.DescontoSubTotal "
  strSQL = strSQL & " HAVING ( Saídas.Efetivada ) AND ( NOT Saídas.[Nota Cancelada]) AND ( [Operações Saída].Tipo = 'V' ) AND Saídas.DescontoSubTotal > 0"
  
  strSQL = strSQL & " AND (Saídas.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Saídas.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(Nome_Empresa.Caption)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Filial = " & Combo.Text & ") "
  End If
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If (.BOF And .EOF) Then
      Exit Function
    End If
    
    .MoveLast
    .MoveFirst
    
    pgbProgress.min = 0
    pgbProgress.Max = .RecordCount + 1
  End With

  With rstVendas
    .MoveFirst
    
    dbTemp.Execute "DELETE * FROM tblRelVendasDescontoSubTotal"
    
    Do While Not .EOF
      strSQL = " SELECT * FROM tblRelVendasDescontoSubTotal WHERE filID = " & .Fields("Filial")
      strSQL = strSQL & " AND movSequencia = " & .Fields("Sequência")
      
      If CDbl(.Fields("DescontoSubTotal")) > 0 Then
        Set rstDescontoSubTotal = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        
        If (rstDescontoSubTotal.BOF And rstDescontoSubTotal.EOF) Then
          dblDescontoSomar = .Fields("DescontoSubTotal")
          
          rstDescontoSubTotal.AddNew
          rstDescontoSubTotal.Fields("filID") = .Fields("Filial")
          rstDescontoSubTotal.Fields("movSequencia") = .Fields("Sequência")
          rstDescontoSubTotal.Fields("movValorDesconto") = dblDescontoSomar
          rstDescontoSubTotal.Update
        Else
          dblDescontoSomar = 0
        End If
      Else
        dblDescontoSomar = 0
      End If
      rstDescontoSubTotal.Close
      Set rstDescontoSubTotal = Nothing
      
      dblDescontoSubTotal = dblDescontoSubTotal + dblDescontoSomar
      
      pgbProgress.Value = .AbsolutePosition
      .MoveNext
    Loop
  End With
  
  dblValorDesconto = dblDescontoSubTotal
  
  If Not rstVendas Is Nothing Then rstVendas.Close
  Set rstVendas = Nothing
End Function

Sub CalcularVendas()

  Dim strSQL As String
  Dim rstVendas As Recordset
  
  Dim rstSaidasProdutos As Recordset
  Dim rstLucra  As Recordset
  Dim rstProdutos As Recordset
  Dim rstPreco As Recordset
  Dim rstClasse As Recordset
  Dim rstSubClasse As Recordset
  
  Dim strProduto As String
  
  Dim intClasse As Integer
  Dim strClasse As String
  
  Dim intSubClasse As Integer
  Dim strSubClasse As String
  
  Dim strNomeProduto As String
  Dim strCodigoOrdenacao As String
  
  Dim dblPreco As Single
  
  '27/06/2007 - Anderson
  'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
  'Solicitante: Zue
  'strSQL = " SELECT Saídas.* FROM Saídas, [Operações Saída] "
  strSQL = "SELECT Saídas.*, Funcionários.Nome "
  strSQL = strSQL & "FROM [Operações Saída], Saídas INNER JOIN Funcionários ON Saídas.Digitador = Funcionários.Código "
  strSQL = strSQL & " WHERE Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
  strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
  '28/06/2004 - Daniel
  'Adicionado linha para buscar em Saídas por Filial, AND Filial
  strSQL = strSQL & " AND Saídas.Filial = " & CByte(Combo.Text)
''  strSQL = strSQL & " AND Saídas.Operação = [Operações Saída].Código AND [Operações Saída].Tipo = 'V' AND Saídas.Efetivada = TRUE AND ( NOT Saídas.[Nota Cancelada])"
  strSQL = strSQL & " AND Saídas.Operação = [Operações Saída].Código AND [Operações Saída].Tipo = 'V' AND Saídas.Efetivada = TRUE AND ( NOT Saídas.[Movimentação Desfeita])"
    
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      
      Call StatusMsg("Gerando arquivo de vendas, aguarde . . . ")
      pgbProgress.Max = .RecordCount
      pgbProgress.Value = 0
      
      Do Until .EOF
        strSQL = " SELECT * FROM [Saídas - Produtos] WHERE Filial = " & .Fields("Filial")
        strSQL = strSQL & " AND Sequência = " & .Fields("Sequência")
        
        Set rstSaidasProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
        With rstSaidasProdutos
          If Not (.BOF And .EOF) Then
            .MoveFirst
            
            Do Until .EOF
              Set rstProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE Código = '" & .Fields("Código Sem Grade") & "'", dbOpenSnapshot)
              strProduto = .Fields("Código Sem Grade") & ""
              
              If (rstProdutos.BOF And rstProdutos.EOF) Then
                intClasse = 0
                strClasse = ""
                intSubClasse = 0
                strSubClasse = ""
                
                strNomeProduto = "<Produto_não_cadastrado>"
                strCodigoOrdenacao = ""
              Else
                intClasse = rstProdutos.Fields("Classe")
                
                Set rstClasse = db.OpenRecordset("SELECT * FROM Classes WHERE Código = " & intClasse, dbOpenSnapshot)
                If (rstClasse.BOF And rstClasse.EOF) Then
                  strClasse = "<Classe_não_cadastrada>"
                Else
                  strClasse = rstClasse.Fields("Nome") & ""
                End If
                rstClasse.Close
                Set rstClasse = Nothing
                
                intSubClasse = rstProdutos.Fields("Sub Classe")
                
                '28/12/2004 - Daniel
                'BUG........: A query que estava sendo passada era SELECT * FROM Classes ...
                'Correção...: SELECT * FROM [Sub Classes] ...
                Set rstSubClasse = db.OpenRecordset("SELECT * FROM [Sub Classes] WHERE Código = " & intSubClasse, dbOpenSnapshot)
                If (rstSubClasse.BOF And rstSubClasse.EOF) Then
                  strSubClasse = "<Sub-Classe_não_cadastrada>"
                Else
                  strSubClasse = rstSubClasse.Fields("Nome") & ""
                End If
                rstSubClasse.Close
                Set rstSubClasse = Nothing
                
                strNomeProduto = rstProdutos.Fields("Nome") & ""
                strCodigoOrdenacao = rstProdutos.Fields("Código Ordenação") & ""
              End If
              
              rstProdutos.Close
              Set rstProdutos = Nothing
              
              Set rstPreco = db.OpenRecordset("SELECT * FROM Preços WHERE Tabela = '" & Combo_Preço.Text & "' AND Produto = '" & strProduto & "'")
              
              If (rstPreco.BOF And rstPreco.EOF) Then
                dblPreco = 0
              Else
                dblPreco = rstPreco.Fields("Preço")
              End If
              
              rstPreco.Close
              Set rstPreco = Nothing
              
              '---[ Preenche a tabela de lucratividade ]---'
                '27/06/2007 - Anderson
                'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
                'Solicitante: Zue
                'Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "'", dbOpenDynaset)
                Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "' And Vendedor=" & rstVendas("Digitador") & " AND Grupo='1 - Vendas'", dbOpenDynaset)
                
                If (rstLucra.BOF And rstLucra.EOF) Then
                  rstLucra.AddNew
                  
                  rstLucra("Produto") = strProduto
                  rstLucra("Código Ordenação") = strCodigoOrdenacao
                  rstLucra("Nome") = strNomeProduto
                  
                  rstLucra("Classe") = intClasse
                  rstLucra("Nome Classe") = strClasse
                  rstLucra("Sub Classe") = intSubClasse
                  rstLucra("Nome Sub") = strSubClasse
                  
                  '27/06/2007 - Anderson
                  'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
                  'Solicitante: Zue
                  rstLucra("Vendedor") = rstVendas("Digitador")
                  rstLucra("Nome Vendedor") = rstVendas("Nome")
  
                  rstLucra("Qtde") = 0
                  rstLucra("Valor") = 0
                  rstLucra("Custo") = 0
                  rstLucra("Lucro") = 0
                  '14/06/2007
                  'Alteração realizada para atender solicitação da Zue
                  'rstLucra("QtdeDevolvida") = 0
                Else
                  rstLucra.Edit
                End If
                
                '15/06/2007 - Anderson
                'Utilizado para contabilizar a quantidde de itens vendidos
                dblTotalQuantidadeVendida = dblTotalQuantidadeVendida + .Fields("Qtde")
                dblTotalValorVenda = dblTotalValorVenda + .Fields("Preço Final")
                '14/06/2007 - Anderson
                'Descontinuado o cálculo de devoluções
                  '12/04/2007 - Anderson
                  'Totaliza a quantidade total de devoluções
                  'dblValorTotalDev = 0
                  'dblTotalDescSub = 0
                
                '14/06/2007
                'Alterado para atender as novas exigências das Zue para o relatório de Lucratividade
                '27/04/2007 - Anderson
                'Implementada a inicialização da variável para evitar problemas com a quantidade do produto.
                'dblQuantidadeTotalDev = 0
                'ReturnDevolucaoNormal dblQuantidadeTotalDev, Quantidade, strProduto
                'ReturnDevolucaoGrade dblQuantidadeTotalDev, Quantidade, strProduto
                
                '14/06/2007
                'Alterado para atender as novas exigências das Zue para o relatório de Lucratividade
                '12/04/2007 - Anderson
                'rstLucra("QtdeDevolvida") = dblQuantidadeTotalDev
                rstLucra("Qtde") = rstLucra("Qtde") + .Fields("Qtde")
                '14/06/2007 - Anderson
                'Alterado para atender as novas exigências das Zue para o relatório de Lucratividade
                '12/04/2007 - Anderson
                'rstLucra("Valor") = rstLucra("Valor") + .Fields("Preço Final")
                'rstLucra("Valor") = ((rstLucra("Qtde") - rstLucra("QtdeDevolvida")) * (.Fields("Preço Final") / .Fields("Qtde")))
                rstLucra("Valor") = rstLucra("Valor") + .Fields("Preço Final")
                '14/06/2007 - Anderson
                'Alterado para atender as novas exigências das Zue para o relatório de Lucratividade
                '12/04/2007 - Anderson
                'rstLucra("Custo") = rstLucra("Custo") + (.Fields("Qtde") * dblPreco)
                'rstLucra("Custo") = rstLucra("Custo") + ((.Fields("Qtde") - dblQuantidadeTotalDev) * dblPreco)
                rstLucra("Custo") = rstLucra("Custo") + (.Fields("Qtde") * dblPreco)
                rstLucra("Lucro") = rstLucra("Valor") - rstLucra("Custo")
                '14/06/2007 - Anderson
                'Implementação do campo Grupo
                rstLucra("Grupo") = "1 - Vendas"
                
                rstLucra.Update
                rstLucra.Close
                Set rstLucra = Nothing
              '---[ Preenche a tabela de lucratividade ]---'
              .MoveNext
            Loop
          End If
          
          .Close
          Set rstSaidasProdutos = Nothing
        End With
        
        pgbProgress.Value = .AbsolutePosition
        
        .MoveNext
      Loop
    End If
    .Close
    Set rstVendas = Nothing
  End With

End Sub

'14/06/2007 - Anderson
'Função criada para atender as exigências da Zue.
Sub CalcularDevolucao()

  Dim rsDevolucao As Recordset
  Dim strSQL As String
  Dim rstEntradaProdutos As Recordset
  Dim rstLucra  As Recordset
  Dim rstProdutos As Recordset
  Dim rstPreco As Recordset
  Dim rstClasse As Recordset
  Dim rstSubClasse As Recordset
  
  Dim strProduto As String
  
  Dim intClasse As Integer
  Dim strClasse As String
  
  Dim intSubClasse As Integer
  Dim strSubClasse As String
  
  Dim strNomeProduto As String
  Dim strCodigoOrdenacao As String
  
  Dim dblPreco As Single

  '27/06/2007 - Anderson
  'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
  'Solicitante: Zue
  'strSQL = " SELECT Entradas.Filial, Entradas.Data, Entradas.Sequência, [Entradas - Produtos].Código, [Entradas - Produtos].[Código sem Grade], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequência = [Entradas - Produtos].Sequência) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN Produtos ON [Entradas - Produtos].[Código sem Grade] = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Sequência, [Entradas - Produtos].Código, [Entradas - Produtos].[Código sem Grade], Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING [Operações Entrada].Tipo='D' "
  strSQL = " SELECT Entradas.Filial, Entradas.Data, Entradas.Sequência, [Entradas - Produtos].Código, [Entradas - Produtos].[Código sem Grade], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal, Entradas.Digitador, Funcionários.Nome " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequência = [Entradas - Produtos].Sequência)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN Produtos ON [Entradas - Produtos].[Código sem Grade] = Produtos.Código) INNER JOIN Funcionários ON Entradas.Digitador = Funcionários.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Sequência, [Entradas - Produtos].Código, [Entradas - Produtos].[Código sem Grade], Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe], Entradas.Digitador, Funcionários.Nome " & _
           " HAVING [Operações Entrada].Tipo='D' "

  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "

  If Len(Trim(Nome_Empresa.Caption)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & Combo.Text & ") "
  End If

  Set rsDevolucao = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rsDevolucao
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      
      Call StatusMsg("Gerando arquivo de devoluções, aguarde . . . ")
      pgbProgress.Max = .RecordCount
      pgbProgress.Value = 0
      
      Do Until .EOF
        strSQL = " SELECT * FROM [Entradas - Produtos] WHERE Filial = " & .Fields("Filial")
        strSQL = strSQL & " AND Sequência = " & .Fields("Sequência") & " AND Código='" & .Fields("Código") & "' "
        
        Set rstEntradaProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
        With rstEntradaProdutos
          If Not (.BOF And .EOF) Then
            .MoveFirst
            
            Do Until .EOF
              Set rstProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE Código = '" & .Fields("Código Sem Grade") & "'", dbOpenSnapshot)
              strProduto = .Fields("Código Sem Grade") & ""
              
              If (rstProdutos.BOF And rstProdutos.EOF) Then
                intClasse = 0
                strClasse = ""
                intSubClasse = 0
                strSubClasse = ""
                
                strNomeProduto = "<Produto_não_cadastrado>"
                strCodigoOrdenacao = ""
              Else
                intClasse = rstProdutos.Fields("Classe")
                
                Set rstClasse = db.OpenRecordset("SELECT * FROM Classes WHERE Código = " & intClasse, dbOpenSnapshot)
                If (rstClasse.BOF And rstClasse.EOF) Then
                  strClasse = "<Classe_não_cadastrada>"
                Else
                  strClasse = rstClasse.Fields("Nome") & ""
                End If
                rstClasse.Close
                Set rstClasse = Nothing
                
                intSubClasse = rstProdutos.Fields("Sub Classe")
                
                Set rstSubClasse = db.OpenRecordset("SELECT * FROM [Sub Classes] WHERE Código = " & intSubClasse, dbOpenSnapshot)
                If (rstSubClasse.BOF And rstSubClasse.EOF) Then
                  strSubClasse = "<Sub-Classe_não_cadastrada>"
                Else
                  strSubClasse = rstSubClasse.Fields("Nome") & ""
                End If
                rstSubClasse.Close
                Set rstSubClasse = Nothing
                
                strNomeProduto = rstProdutos.Fields("Nome") & ""
                strCodigoOrdenacao = rstProdutos.Fields("Código Ordenação") & ""
              End If
              
              rstProdutos.Close
              Set rstProdutos = Nothing
              
              Set rstPreco = db.OpenRecordset("SELECT * FROM Preços WHERE Tabela = '" & Combo_Preço.Text & "' AND Produto = '" & strProduto & "'")
              
              If (rstPreco.BOF And rstPreco.EOF) Then
                dblPreco = 0
              Else
                dblPreco = rstPreco.Fields("Preço")
              End If
              
              rstPreco.Close
              Set rstPreco = Nothing
              
              '---[ Preenche a tabela de lucratividade ]---'
                '27/06/2007 - Anderson
                'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
                'Solicitante: Zue
                'Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "' AND Grupo='2 - Devoluções'", dbOpenDynaset)
                Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "' AND Vendedor=" & rsDevolucao("Digitador") & " AND Grupo='2 - Devoluções'", dbOpenDynaset)
                
                If (rstLucra.BOF And rstLucra.EOF) Then
                  rstLucra.AddNew
                  
                  rstLucra("Produto") = strProduto
                  rstLucra("Código Ordenação") = strCodigoOrdenacao
                  rstLucra("Nome") = strNomeProduto
                  
                  rstLucra("Classe") = intClasse
                  rstLucra("Nome Classe") = strClasse
                  rstLucra("Sub Classe") = intSubClasse
                  rstLucra("Nome Sub") = strSubClasse
                  
                  '27/06/2007 - Anderson
                  'Alteração realizada para a criação do relatório de lucrativdade separado por vendedor
                  'Solicitante: Zue
                  rstLucra("Vendedor") = rsDevolucao("Digitador")
                  rstLucra("Nome Vendedor") = rsDevolucao("Nome")
  
                  rstLucra("Qtde") = 0
                  rstLucra("Valor") = 0
                  rstLucra("Custo") = 0
                  rstLucra("Lucro") = 0
                Else
                  rstLucra.Edit
                End If
                
                '15/06/2007 - Anderson
                'Utilizado para contabilizar a quantidde de itens devolvidos
                dblTotalQuantidadeDevolvida = dblTotalQuantidadeDevolvida + .Fields("Qtde")
                
                rstLucra("Qtde") = -(Abs(rstLucra("Qtde")) + (.Fields("Qtde")))
                rstLucra("Valor") = -(Abs(rstLucra("Valor")) + (.Fields("Preço Final")))
                rstLucra("Custo") = -(Abs(rstLucra("Custo")) + (.Fields("Qtde") * dblPreco))
                rstLucra("Lucro") = -(Abs(rstLucra("Valor")) + Abs(rstLucra("Custo")))
                rstLucra("Grupo") = "2 - Devoluções"
                
                rstLucra.Update
                rstLucra.Close
                Set rstLucra = Nothing

              .MoveNext
            Loop
          End If
          
          .Close
          Set rstEntradaProdutos = Nothing
        End With
        
        pgbProgress.Value = .AbsolutePosition

        .MoveNext
      Loop
    End If
    .Close
    Set rsDevolucao = Nothing
  End With
  
End Sub
