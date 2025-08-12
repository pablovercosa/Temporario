VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelLucratividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relat�rio de Lucratividade"
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
      RecordSource    =   "SELECT Filial, Nome FROM [Par�metros Filial]"
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
      Caption         =   "Sa�da"
      Height          =   855
      Left            =   4860
      TabIndex        =   7
      Top             =   1920
      Width           =   5745
      Begin VB.OptionButton B_V�deo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "V�deo"
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
      Begin VB.OptionButton O_C�digo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "C�digo"
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
      Caption         =   "Per�odo"
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
         ToolTipText     =   "Pressione F2 para Calend�rio"
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
         ToolTipText     =   "Pressione F2 para Calend�rio"
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
      Begin VB.ComboBox Combo_Pre�o 
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
         Caption         =   "Relat�rio de Lucratividade"
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
         Caption         =   "Selecione abaixo a tabela de pre�os de acordo com o custo que voc� deseja calcular."
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
         Caption         =   "� Caso voc� utilize o sistema de desconto no sub-total, o Quick Store contabiliza os descontos dados como desconto financeiro. "
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
  'Descontinuado para atender as novas exig�ncias da Zue
  'Dim dblQuantidadeTotalDev As Double '12/04/2007 - Anderson - Utilizado para obter a quantidade total de produtos devolvidos.
  Dim dblTotalDescSub As Double
  
  If Len(Trim(Nome_Empresa.Caption)) <= 0 Then
    MsgBox "Filial inv�lida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(Data_Ini.Text) Then
    MsgBox "Data inicial inv�lida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(Data_Fim.Text) Then
    MsgBox "Data inicial inv�lida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    MsgBox "A data inicial n�o pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  dbTemp.Execute "DELETE * FROM Lucratividade"
  
  '15/06/2007 - Anderson
  'Vari�veis para serem contabilizadas no totalizador do relat�rio
  dblTotalQuantidadeVendida = 0
  dblTotalQuantidadeDevolvida = 0
  dblTotalValorVenda = 0
  
  '14/06/2007 - Anderson
  'Fun��o utilizada para c�lcular o total de vendas no relat�rio de Lucratividade
  Call StatusMsg("Calculando vendas, aguarde . . . ")
  Call CalcularVendas

  Call StatusMsg("Calculando devolu��es, aguarde . . . ")
  Call CalcularDevolucao
  
  Call StatusMsg("Analisando os descontos no sub-total e devolu��es, aguarde . . . ")
  '---[ Gera o total de Descontos do sub-total ]---'
    dblValorTotalDev = 0
    dblTotalDescSub = 0
    
    ReturnDescontoSubTotal dblTotalDescSub
    '14/06/2007 - Anderson
    'Descontinuado para atender a solicita��o da Zue
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
    
    '.Destination = Not B_V�deo.Value
    '29/11/2004 - Daniel
    'Bug: N�o estava localizando a impressora default
    'ao imprimir direto na impressora
    If B_V�deo.Value Then
      .Destination = crptToWindow
    Else
      .Destination = crptToPrinter
    End If
    
    If O_Normal.Value Then .ReportFileName = gsReportPath & "LUCRA.RPT"
    If O_Classe.Value Then .ReportFileName = gsReportPath & "LUCRA2.RPT"
    '27/06/2007 - Anderson
    'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
    'Solicitante: Zue
    If O_Vendedor.Value Then .ReportFileName = gsReportPath & "LUCRA3.RPT"
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 Rel
  
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'"
    .Formulas(1) = "nome_filial = '" & Nome_Empresa.Caption & "'"
    .Formulas(2) = "data_ini = '" & Data_Ini.Text & "'"
    .Formulas(3) = "data_fim = '" & Data_Fim.Text & "'"
    .Formulas(4) = "tipo_rel = '" & Combo_Pre�o.Text & "'"
    .Formulas(5) = "DescSubTotal = " & Replace(dblTotalDescSub, ",", ".")
    .Formulas(6) = "DevolucoesValor = " & Replace(dblValorTotalDev, ",", ".")
    .Formulas(7) = "TotalQuantidadeVendida = " & Replace(dblTotalQuantidadeVendida, ",", ".")
    .Formulas(8) = "TotalQuantidadeDevolvida = " & Replace(dblTotalQuantidadeDevolvida, ",", ".")
    .Formulas(9) = "TotalValorVenda = " & Replace(dblTotalValorVenda, ",", ".")
    
    If O_C�digo.Value Then
      If O_Normal.Value Then
        Rel.SortFields(0) = "+{Lucratividade.C�digo Ordena��o}"
        Rel.SortFields(1) = ""
        Rel.SortFields(2) = ""
      End If
      If O_Classe.Value Then
        Rel.SortFields(0) = "+{Lucratividade.Classe}"
        Rel.SortFields(1) = "+{Lucratividade.Sub Classe}"
        Rel.SortFields(2) = "+{Lucratividade.C�digo Ordena��o}"
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
    'Seta a impressora para relat�rio
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
  
  '---[ Preenchimento da combo de pre�os ]---'
  Set rstTabelasPreco = db.OpenRecordset("SELECT DISTINCTROW Tabela FROM [Tabela de Pre�os] ORDER BY Tabela", dbOpenSnapshot)
  
  With rstTabelasPreco
    Combo_Pre�o.Clear
    If Not (.BOF And .EOF) Then .MoveFirst
    
    Do Until .EOF
      Combo_Pre�o.AddItem .Fields("Tabela") & ""
      .MoveNext
    Loop
    
    .Close
    Set rstTabelasPreco = Nothing
  End With
  '---[ Preenchimento da combo de pre�os ]---'
End Sub

Private Function ReturnDevolucaoGrade(ByRef dblValorDevolucao As Double) As Boolean
  Dim strSQL As String
  Dim rstDev As Recordset
  Dim blnProdutoOK As Boolean

  strSQL = " SELECT Entradas.Filial, Entradas.Data, [C�digos da Grade].[C�digo Original], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Pre�o Final]) AS PrecoTotal " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) INNER JOIN [C�digos da Grade] ON [Entradas - Produtos].C�digo = [C�digos da Grade].C�digo) INNER JOIN Produtos ON [C�digos da Grade].[C�digo Original] = Produtos.C�digo " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [C�digos da Grade].[C�digo Original], Entradas.Fornecedor, [Opera��es Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Opera��es Entrada].Tipo)='D')) "


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

  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Entradas - Produtos].C�digo, Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Pre�o Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) INNER JOIN Produtos ON [Entradas - Produtos].C�digo = Produtos.C�digo " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [Entradas - Produtos].C�digo, Entradas.Fornecedor, [Opera��es Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Opera��es Entrada].Tipo)='D')) "

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
  
  strSQL = " SELECT SUM(Sa�das.DescontoSubTotal) AS DescontoSubTotal, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Filial, Sa�das.Sequ�ncia "
  strSQL = strSQL & " FROM ((Sa�das INNER JOIN [Sa�das - Produtos] ON (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia) AND (Sa�das.Filial = [Sa�das - Produtos].Filial)) INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
  strSQL = strSQL & " GROUP BY Sa�das.Filial, Sa�das.Data, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Digitador, Produtos.Classe, Produtos.[Sub Classe], Sa�das.Efetivada, Sa�das.[Nota Cancelada], [Opera��es Sa�da].Tipo = 'V', Sa�das.Sequ�ncia, Sa�das.DescontoSubTotal "
  strSQL = strSQL & " HAVING ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Sa�das.DescontoSubTotal > 0"
  
  strSQL = strSQL & " AND (Sa�das.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Sa�das.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(Nome_Empresa.Caption)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Filial = " & Combo.Text & ") "
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
      strSQL = strSQL & " AND movSequencia = " & .Fields("Sequ�ncia")
      
      If CDbl(.Fields("DescontoSubTotal")) > 0 Then
        Set rstDescontoSubTotal = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        
        If (rstDescontoSubTotal.BOF And rstDescontoSubTotal.EOF) Then
          dblDescontoSomar = .Fields("DescontoSubTotal")
          
          rstDescontoSubTotal.AddNew
          rstDescontoSubTotal.Fields("filID") = .Fields("Filial")
          rstDescontoSubTotal.Fields("movSequencia") = .Fields("Sequ�ncia")
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
  'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
  'Solicitante: Zue
  'strSQL = " SELECT Sa�das.* FROM Sa�das, [Opera��es Sa�da] "
  strSQL = "SELECT Sa�das.*, Funcion�rios.Nome "
  strSQL = strSQL & "FROM [Opera��es Sa�da], Sa�das INNER JOIN Funcion�rios ON Sa�das.Digitador = Funcion�rios.C�digo "
  strSQL = strSQL & " WHERE Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
  strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
  '28/06/2004 - Daniel
  'Adicionado linha para buscar em Sa�das por Filial, AND Filial
  strSQL = strSQL & " AND Sa�das.Filial = " & CByte(Combo.Text)
''  strSQL = strSQL & " AND Sa�das.Opera��o = [Opera��es Sa�da].C�digo AND [Opera��es Sa�da].Tipo = 'V' AND Sa�das.Efetivada = TRUE AND ( NOT Sa�das.[Nota Cancelada])"
  strSQL = strSQL & " AND Sa�das.Opera��o = [Opera��es Sa�da].C�digo AND [Opera��es Sa�da].Tipo = 'V' AND Sa�das.Efetivada = TRUE AND ( NOT Sa�das.[Movimenta��o Desfeita])"
    
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      
      Call StatusMsg("Gerando arquivo de vendas, aguarde . . . ")
      pgbProgress.Max = .RecordCount
      pgbProgress.Value = 0
      
      Do Until .EOF
        strSQL = " SELECT * FROM [Sa�das - Produtos] WHERE Filial = " & .Fields("Filial")
        strSQL = strSQL & " AND Sequ�ncia = " & .Fields("Sequ�ncia")
        
        Set rstSaidasProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
        With rstSaidasProdutos
          If Not (.BOF And .EOF) Then
            .MoveFirst
            
            Do Until .EOF
              Set rstProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE C�digo = '" & .Fields("C�digo Sem Grade") & "'", dbOpenSnapshot)
              strProduto = .Fields("C�digo Sem Grade") & ""
              
              If (rstProdutos.BOF And rstProdutos.EOF) Then
                intClasse = 0
                strClasse = ""
                intSubClasse = 0
                strSubClasse = ""
                
                strNomeProduto = "<Produto_n�o_cadastrado>"
                strCodigoOrdenacao = ""
              Else
                intClasse = rstProdutos.Fields("Classe")
                
                Set rstClasse = db.OpenRecordset("SELECT * FROM Classes WHERE C�digo = " & intClasse, dbOpenSnapshot)
                If (rstClasse.BOF And rstClasse.EOF) Then
                  strClasse = "<Classe_n�o_cadastrada>"
                Else
                  strClasse = rstClasse.Fields("Nome") & ""
                End If
                rstClasse.Close
                Set rstClasse = Nothing
                
                intSubClasse = rstProdutos.Fields("Sub Classe")
                
                '28/12/2004 - Daniel
                'BUG........: A query que estava sendo passada era SELECT * FROM Classes ...
                'Corre��o...: SELECT * FROM [Sub Classes] ...
                Set rstSubClasse = db.OpenRecordset("SELECT * FROM [Sub Classes] WHERE C�digo = " & intSubClasse, dbOpenSnapshot)
                If (rstSubClasse.BOF And rstSubClasse.EOF) Then
                  strSubClasse = "<Sub-Classe_n�o_cadastrada>"
                Else
                  strSubClasse = rstSubClasse.Fields("Nome") & ""
                End If
                rstSubClasse.Close
                Set rstSubClasse = Nothing
                
                strNomeProduto = rstProdutos.Fields("Nome") & ""
                strCodigoOrdenacao = rstProdutos.Fields("C�digo Ordena��o") & ""
              End If
              
              rstProdutos.Close
              Set rstProdutos = Nothing
              
              Set rstPreco = db.OpenRecordset("SELECT * FROM Pre�os WHERE Tabela = '" & Combo_Pre�o.Text & "' AND Produto = '" & strProduto & "'")
              
              If (rstPreco.BOF And rstPreco.EOF) Then
                dblPreco = 0
              Else
                dblPreco = rstPreco.Fields("Pre�o")
              End If
              
              rstPreco.Close
              Set rstPreco = Nothing
              
              '---[ Preenche a tabela de lucratividade ]---'
                '27/06/2007 - Anderson
                'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
                'Solicitante: Zue
                'Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "'", dbOpenDynaset)
                Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "' And Vendedor=" & rstVendas("Digitador") & " AND Grupo='1 - Vendas'", dbOpenDynaset)
                
                If (rstLucra.BOF And rstLucra.EOF) Then
                  rstLucra.AddNew
                  
                  rstLucra("Produto") = strProduto
                  rstLucra("C�digo Ordena��o") = strCodigoOrdenacao
                  rstLucra("Nome") = strNomeProduto
                  
                  rstLucra("Classe") = intClasse
                  rstLucra("Nome Classe") = strClasse
                  rstLucra("Sub Classe") = intSubClasse
                  rstLucra("Nome Sub") = strSubClasse
                  
                  '27/06/2007 - Anderson
                  'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
                  'Solicitante: Zue
                  rstLucra("Vendedor") = rstVendas("Digitador")
                  rstLucra("Nome Vendedor") = rstVendas("Nome")
  
                  rstLucra("Qtde") = 0
                  rstLucra("Valor") = 0
                  rstLucra("Custo") = 0
                  rstLucra("Lucro") = 0
                  '14/06/2007
                  'Altera��o realizada para atender solicita��o da Zue
                  'rstLucra("QtdeDevolvida") = 0
                Else
                  rstLucra.Edit
                End If
                
                '15/06/2007 - Anderson
                'Utilizado para contabilizar a quantidde de itens vendidos
                dblTotalQuantidadeVendida = dblTotalQuantidadeVendida + .Fields("Qtde")
                dblTotalValorVenda = dblTotalValorVenda + .Fields("Pre�o Final")
                '14/06/2007 - Anderson
                'Descontinuado o c�lculo de devolu��es
                  '12/04/2007 - Anderson
                  'Totaliza a quantidade total de devolu��es
                  'dblValorTotalDev = 0
                  'dblTotalDescSub = 0
                
                '14/06/2007
                'Alterado para atender as novas exig�ncias das Zue para o relat�rio de Lucratividade
                '27/04/2007 - Anderson
                'Implementada a inicializa��o da vari�vel para evitar problemas com a quantidade do produto.
                'dblQuantidadeTotalDev = 0
                'ReturnDevolucaoNormal dblQuantidadeTotalDev, Quantidade, strProduto
                'ReturnDevolucaoGrade dblQuantidadeTotalDev, Quantidade, strProduto
                
                '14/06/2007
                'Alterado para atender as novas exig�ncias das Zue para o relat�rio de Lucratividade
                '12/04/2007 - Anderson
                'rstLucra("QtdeDevolvida") = dblQuantidadeTotalDev
                rstLucra("Qtde") = rstLucra("Qtde") + .Fields("Qtde")
                '14/06/2007 - Anderson
                'Alterado para atender as novas exig�ncias das Zue para o relat�rio de Lucratividade
                '12/04/2007 - Anderson
                'rstLucra("Valor") = rstLucra("Valor") + .Fields("Pre�o Final")
                'rstLucra("Valor") = ((rstLucra("Qtde") - rstLucra("QtdeDevolvida")) * (.Fields("Pre�o Final") / .Fields("Qtde")))
                rstLucra("Valor") = rstLucra("Valor") + .Fields("Pre�o Final")
                '14/06/2007 - Anderson
                'Alterado para atender as novas exig�ncias das Zue para o relat�rio de Lucratividade
                '12/04/2007 - Anderson
                'rstLucra("Custo") = rstLucra("Custo") + (.Fields("Qtde") * dblPreco)
                'rstLucra("Custo") = rstLucra("Custo") + ((.Fields("Qtde") - dblQuantidadeTotalDev) * dblPreco)
                rstLucra("Custo") = rstLucra("Custo") + (.Fields("Qtde") * dblPreco)
                rstLucra("Lucro") = rstLucra("Valor") - rstLucra("Custo")
                '14/06/2007 - Anderson
                'Implementa��o do campo Grupo
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
'Fun��o criada para atender as exig�ncias da Zue.
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
  'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
  'Solicitante: Zue
  'strSQL = " SELECT Entradas.Filial, Entradas.Data, Entradas.Sequ�ncia, [Entradas - Produtos].C�digo, [Entradas - Produtos].[C�digo sem Grade], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Pre�o Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) INNER JOIN Produtos ON [Entradas - Produtos].[C�digo sem Grade] = Produtos.C�digo " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Sequ�ncia, [Entradas - Produtos].C�digo, [Entradas - Produtos].[C�digo sem Grade], Entradas.Fornecedor, [Opera��es Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING [Opera��es Entrada].Tipo='D' "
  strSQL = " SELECT Entradas.Filial, Entradas.Data, Entradas.Sequ�ncia, [Entradas - Produtos].C�digo, [Entradas - Produtos].[C�digo sem Grade], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Pre�o Final]) AS PrecoTotal, Entradas.Digitador, Funcion�rios.Nome " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) INNER JOIN Produtos ON [Entradas - Produtos].[C�digo sem Grade] = Produtos.C�digo) INNER JOIN Funcion�rios ON Entradas.Digitador = Funcion�rios.C�digo " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Sequ�ncia, [Entradas - Produtos].C�digo, [Entradas - Produtos].[C�digo sem Grade], Entradas.Fornecedor, [Opera��es Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe], Entradas.Digitador, Funcion�rios.Nome " & _
           " HAVING [Opera��es Entrada].Tipo='D' "

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
      
      Call StatusMsg("Gerando arquivo de devolu��es, aguarde . . . ")
      pgbProgress.Max = .RecordCount
      pgbProgress.Value = 0
      
      Do Until .EOF
        strSQL = " SELECT * FROM [Entradas - Produtos] WHERE Filial = " & .Fields("Filial")
        strSQL = strSQL & " AND Sequ�ncia = " & .Fields("Sequ�ncia") & " AND C�digo='" & .Fields("C�digo") & "' "
        
        Set rstEntradaProdutos = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
        With rstEntradaProdutos
          If Not (.BOF And .EOF) Then
            .MoveFirst
            
            Do Until .EOF
              Set rstProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE C�digo = '" & .Fields("C�digo Sem Grade") & "'", dbOpenSnapshot)
              strProduto = .Fields("C�digo Sem Grade") & ""
              
              If (rstProdutos.BOF And rstProdutos.EOF) Then
                intClasse = 0
                strClasse = ""
                intSubClasse = 0
                strSubClasse = ""
                
                strNomeProduto = "<Produto_n�o_cadastrado>"
                strCodigoOrdenacao = ""
              Else
                intClasse = rstProdutos.Fields("Classe")
                
                Set rstClasse = db.OpenRecordset("SELECT * FROM Classes WHERE C�digo = " & intClasse, dbOpenSnapshot)
                If (rstClasse.BOF And rstClasse.EOF) Then
                  strClasse = "<Classe_n�o_cadastrada>"
                Else
                  strClasse = rstClasse.Fields("Nome") & ""
                End If
                rstClasse.Close
                Set rstClasse = Nothing
                
                intSubClasse = rstProdutos.Fields("Sub Classe")
                
                Set rstSubClasse = db.OpenRecordset("SELECT * FROM [Sub Classes] WHERE C�digo = " & intSubClasse, dbOpenSnapshot)
                If (rstSubClasse.BOF And rstSubClasse.EOF) Then
                  strSubClasse = "<Sub-Classe_n�o_cadastrada>"
                Else
                  strSubClasse = rstSubClasse.Fields("Nome") & ""
                End If
                rstSubClasse.Close
                Set rstSubClasse = Nothing
                
                strNomeProduto = rstProdutos.Fields("Nome") & ""
                strCodigoOrdenacao = rstProdutos.Fields("C�digo Ordena��o") & ""
              End If
              
              rstProdutos.Close
              Set rstProdutos = Nothing
              
              Set rstPreco = db.OpenRecordset("SELECT * FROM Pre�os WHERE Tabela = '" & Combo_Pre�o.Text & "' AND Produto = '" & strProduto & "'")
              
              If (rstPreco.BOF And rstPreco.EOF) Then
                dblPreco = 0
              Else
                dblPreco = rstPreco.Fields("Pre�o")
              End If
              
              rstPreco.Close
              Set rstPreco = Nothing
              
              '---[ Preenche a tabela de lucratividade ]---'
                '27/06/2007 - Anderson
                'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
                'Solicitante: Zue
                'Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "' AND Grupo='2 - Devolu��es'", dbOpenDynaset)
                Set rstLucra = dbTemp.OpenRecordset("SELECT * FROM Lucratividade WHERE Produto = '" & strProduto & "' AND Vendedor=" & rsDevolucao("Digitador") & " AND Grupo='2 - Devolu��es'", dbOpenDynaset)
                
                If (rstLucra.BOF And rstLucra.EOF) Then
                  rstLucra.AddNew
                  
                  rstLucra("Produto") = strProduto
                  rstLucra("C�digo Ordena��o") = strCodigoOrdenacao
                  rstLucra("Nome") = strNomeProduto
                  
                  rstLucra("Classe") = intClasse
                  rstLucra("Nome Classe") = strClasse
                  rstLucra("Sub Classe") = intSubClasse
                  rstLucra("Nome Sub") = strSubClasse
                  
                  '27/06/2007 - Anderson
                  'Altera��o realizada para a cria��o do relat�rio de lucrativdade separado por vendedor
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
                rstLucra("Valor") = -(Abs(rstLucra("Valor")) + (.Fields("Pre�o Final")))
                rstLucra("Custo") = -(Abs(rstLucra("Custo")) + (.Fields("Qtde") * dblPreco))
                rstLucra("Lucro") = -(Abs(rstLucra("Valor")) + Abs(rstLucra("Custo")))
                rstLucra("Grupo") = "2 - Devolu��es"
                
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
