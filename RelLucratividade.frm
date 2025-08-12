VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelLucratividade_OLD 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Lucratividade"
   ClientHeight    =   7725
   ClientLeft      =   2640
   ClientTop       =   2760
   ClientWidth     =   6975
   ForeColor       =   &H80000008&
   HelpContextID   =   1620
   Icon            =   "RelLucratividade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7725
   ScaleWidth      =   6975
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   27
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   -120
      TabIndex        =   23
      Top             =   -240
      Width           =   8055
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "• Caso você utilize o sistema de desconto no sub-total, o Quick Store contabiliza os descontos dados como desconto financeiro. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   2160
         TabIndex        =   26
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione abaixo a tabela de preços a ser feita a comparação."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   2160
         TabIndex        =   25
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Lucratividade"
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
         Left            =   2040
         TabIndex        =   24
         Top             =   360
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "RelLucratividade.frx":058A
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Cálculo da Lucratividade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   4575
      Begin VB.ComboBox Combo_Preço 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Tabela de Preços"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   22
         Top             =   405
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Opções"
      Height          =   735
      Left            =   2040
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   5160
      Begin VB.CheckBox O_Zero 
         Caption         =   "Imprimir produtos com vendas - devolução = 0"
         Height          =   225
         Left            =   285
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   4110
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   4575
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
         Left            =   720
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fim"
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
         Left            =   2640
         TabIndex        =   19
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         TabIndex        =   18
         Top             =   390
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
      Begin VB.OptionButton O_Código 
         Caption         =   "Código"
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
         Left            =   105
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton O_Nome 
         Caption         =   "Nome"
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
         Left            =   105
         TabIndex        =   10
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   855
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   2055
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
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton O_Classe 
         Caption         =   "Separado por classe"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1845
      End
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
      Height          =   855
      Left            =   4800
      TabIndex        =   14
      Top             =   2040
      Width           =   2055
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1110
      End
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
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H0000C0C0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   1080
      Top             =   6240
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
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelLucratividade.frx":23F2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   1560
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
      _ExtentY        =   556
      _StockProps     =   93
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
   Begin VB.Label Nome_Empresa 
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
      Left            =   2280
      TabIndex        =   13
      Top             =   1560
      Width           =   4560
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial"
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
      TabIndex        =   12
      Top             =   1605
      Width           =   615
   End
End
Attribute VB_Name = "frmRelLucratividade_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLucra As Recordset
Dim rsPreços As Recordset
Dim rsProdutos As Recordset
Dim rsEstoque As Recordset
Dim rsParametros As Recordset
Dim rsClasse As Recordset
Dim rsSub_Classe As Recordset


Private Sub Ano_KeyPress(KeyAscii As Integer)
 KeyAscii = Verifica_Tecla_Integer(KeyAscii)
 
End Sub


Private Sub B_Imprime_Click()
 Dim Aux_Classe As Integer
 Dim Aux_Sub As Integer
 Dim Aux_Produto As String
 Dim Aux_Tamanho As Integer
 Dim Aux_Cor As Integer
 Dim Aux_Edição As Long
 Dim Aux_Ano As Integer
 Dim Aux_Data As Date
 Dim Str1 As String
 Dim Str_Rel As String
 Dim sSql As String
 Dim Nome_Classe As String
 Dim Nome_Sub As String
 Dim dblPreco As Double
 
 Call StatusMsg("")

 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo.SetFocus
   Exit Sub
 End If

 If Filial_Liberada <> 0 Then
   If Val(Combo.Text) <> Filial_Liberada Then
     DisplayMsg "Funcionário não tem acesso a esta filial."
     Exit Sub
   End If
 End If

 Rem Verifica data inicial
 If Not IsDate(Data_Ini.Text) Then
   DisplayMsg "Data inicial incorreta."
   Data_Ini.SetFocus
   Exit Sub
 End If

 Rem Verifica data final
 If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Data final incorreta."
    Data_Fim.SetFocus
    Exit Sub
 End If
 
 If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data inicial deve ser igual ou menor que a data final."
    Data_Ini.SetFocus
    Exit Sub
 End If

 If Combo_Preço.Text = "" Then
   DisplayMsg "Escolha uma tabela de preços."
   Combo_Preço.SetFocus
   Exit Sub
 End If
 
 MousePointer = vbHourglass

 Rem Zera o arquivo Lucra e começa
 sSql = "Delete * From Lucratividade"
 Call StatusMsg("Preparando o arquivo temporário... ")
 dbTemp.Execute sSql
 
 rsEstoque.Index = "Produto"
 rsLucra.Index = "Produto"
 rsPreços.Index = "Tabela"
 rsClasse.Index = "Código"
 rsSub_Classe.Index = "Código"
 rsProdutos.Index = "Código"
 Aux_Classe = 0
 Aux_Sub = 0
 Aux_Tamanho = 0
 Aux_Cor = 0
 Aux_Produto = 0
 Aux_Data = CDate(Data_Ini.Text)
 
LP_Vendas:
 rsEstoque.Seek ">", Val(Combo.Text), Aux_Data, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição
 If rsEstoque.NoMatch Then GoTo Fim_Vendas
 Aux_Produto = rsEstoque("Produto")
 Aux_Tamanho = rsEstoque("Tamanho")
 Aux_Cor = rsEstoque("Cor")
 Aux_Edição = rsEstoque("Edição")
 Aux_Data = rsEstoque("Data")
  
 Call StatusMsg("Lendo vendas ...." & str(Aux_Data))
 DoEvents
  
 If rsEstoque("Filial") <> Val(Combo.Text) Then GoTo Fim_Vendas
 If rsEstoque("Data") > CDate(Data_Fim.Text) Then GoTo Fim_Vendas
 
' If rsEstoque("Vendas") = 0 Or rsEstoque("Valor Vendas") = 0 Then GoTo LP_Vendas

 Rem Acha preço de custo do produto
 rsPreços.Seek "=", Combo_Preço.Text, Aux_Produto
 If rsPreços.NoMatch Then
   dblPreco = 0
 Else
   dblPreco = rsPreços.Fields("Preço").Value
 End If

 rsProdutos.Seek "=", Aux_Produto
 If rsProdutos.NoMatch Then GoTo LP_Vendas

 Nome_Classe = ""
 rsClasse.Seek "=", rsProdutos("Classe")
 If Not rsClasse.NoMatch Then
   Nome_Classe = rsClasse("Nome")
 End If
 
 Nome_Sub = ""
 rsSub_Classe.Seek "=", rsProdutos("Sub Classe")
 If Not rsSub_Classe.NoMatch Then
   Nome_Sub = rsSub_Classe("Nome")
 End If
 
 rsLucra.Seek "=", Aux_Produto
 If rsLucra.NoMatch Then
   rsLucra.AddNew
   rsLucra("Classe") = rsProdutos("Classe")
   rsLucra("Sub Classe") = rsProdutos("Sub Classe")
   rsLucra("Produto") = Aux_Produto
   rsLucra("Qtde") = 0
   rsLucra("Valor") = 0
   rsLucra("Custo") = 0
   rsLucra("Lucro") = 0
 Else
   rsLucra.Edit
 End If
   rsLucra("Nome") = rsProdutos("Nome")
   rsLucra("Código Ordenação") = rsProdutos("Código Ordenação")
   rsLucra("Nome Classe") = Nome_Classe
   rsLucra("Nome Sub") = Nome_Sub
   rsLucra("Qtde") = rsLucra("Qtde") + rsEstoque("Vendas")
   rsLucra("Valor") = rsLucra("Valor") + rsEstoque("Valor Vendas")
   rsLucra("Custo") = rsLucra("Custo") + (rsEstoque("Vendas") * dblPreco)   'rsPreços("Preço")
   rsLucra("Lucro") = rsLucra("Valor") - rsLucra("Custo")

  rsLucra.Update

  GoTo LP_Vendas

Fim_Vendas:
  Rem Apaga linhas com valor 0
  If O_Zero.Value = 0 Then
    sSql = "Delete * From Lucratividade Where Valor = 0"
    Call StatusMsg("Apagando vendas com valor nulo... ")
    dbTemp.Execute sSql
  End If

  '---[ Gera o total de Descontos do sub-total ]---'
    Dim dblValorTotalDev As Double
    Dim rstDescSubTotal As Recordset
    Dim curDescSubTotal As Currency
    Dim strSQL          As String
    
    If IsNumeric(Combo.Text) Then
      strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
               "Filial = " & CLng(Combo.Text) & " AND " & _
               "Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
               "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
    Else
      strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
               "Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
               "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
    End If
    
    Set rstDescSubTotal = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstDescSubTotal
      Call IsDataType(dtCurrency, .Fields("Total").Value, curDescSubTotal)
      If Not rstDescSubTotal Is Nothing Then .Close
      Set rstDescSubTotal = Nothing
    End With
    
    dblValorTotalDev = 0
    
    ReturnDevolucaoNormal dblValorTotalDev
    ReturnDevolucaoGrade dblValorTotalDev
  '---[ Gera o total de Descontos do sub-total ]---'

  Call StatusMsg("")

 Rel.WindowShowGroupTree = O_Classe.Value

 Rem  Nome do BD
  With Rel
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
  End With

 Rem Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1

 Rem Nome do arquivo .rpt
 If O_Normal.Value = True Then Str1 = gsReportPath & "LUCRA.RPT"
 If O_Classe.Value = True Then Str1 = gsReportPath & "LUCRA2.RPT"
 Rel.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel

 Rem Str_Rel = "STR_NOME = 'Empresa " + (DC_Empresas.Text)
 Rem Str_Rel = Str_Rel + " - " + C_Nome_Empresa + " de " + C_Data_Ini.Text + " a " + C_Data_Fim.Text + "'"
 Rem frmMenu.Relatório.Formulas(0) = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel.Formulas(1) = Str_Rel

 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini + "'"
 Rel.Formulas(2) = Str_Rel

 Str_Rel = "data_fim = '"
 Str_Rel = Str_Rel + Data_Fim + "'"
 Rel.Formulas(3) = Str_Rel
 
 Rel.Formulas(4) = "tipo_rel = '" & Combo_Preço.Text & "'"
 Rel.Formulas(5) = "DescSubTotal = " & Replace(curDescSubTotal, ",", ".")
 Rel.Formulas(6) = "DevolucoesValor = " & Replace(dblValorTotalDev, ",", ".")
 
 If O_Código.Value = True Then
  If O_Normal.Value = True Then
    Rel.SortFields(0) = "+{Lucratividade.Código Ordenação}"
    Rel.SortFields(1) = ""
    Rel.SortFields(2) = ""
  End If
  If O_Classe.Value = True Then
    Rel.SortFields(0) = "+{Lucratividade.Classe}"
    Rel.SortFields(1) = "+{Lucratividade.Sub Classe}"
    Rel.SortFields(2) = "+{Lucratividade.Código Ordenação}"
  End If
 End If
 
 If O_Nome.Value = True Then
  If O_Normal.Value = True Then
    Rel.SortFields(0) = "+{Lucratividade.Nome}"
    Rel.SortFields(1) = ""
    Rel.SortFields(2) = ""
  End If
  If O_Classe.Value = True Then
    Rel.SortFields(0) = "+{Lucratividade.Classe}"
    Rel.SortFields(1) = "+{Lucratividade.Sub Classe}"
    Rel.SortFields(2) = "+{Lucratividade.Nome}"
  End If
 End If

 Call StatusMsg("Aguarde, imprimindo...")
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
 
 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub Combo_CloseUp()
  Combo.Text = Combo.Columns(1).Text
  Combo_LostFocus
End Sub

Private Sub Combo_LostFocus()
  Call StatusMsg("")
 
  Nome_Empresa.Caption = ""
  If IsNull(Combo.Text) Then Exit Sub
  If Combo.Text = "" Then Exit Sub
  If Not IsNumeric(Combo.Text) Then Exit Sub
  If Val(Combo.Text) < 0 Then Exit Sub
  If Val(Combo.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

End Sub

Private Sub Data_Ini_LostFocus()
  Data_Ini.Text = Ajusta_Data(Data_Ini.Text)
End Sub

Private Sub Data_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
  End Select
End Sub

Private Sub Data_Fim_LostFocus()
  Data_Fim.Text = Ajusta_Data(Data_Fim.Text)
End Sub

Private Sub Data_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
  End Select
End Sub

Private Sub Form_Load()
 Dim Últ_Tabela As String
 Dim Lugar As Integer

  Call CenterForm(Me)

 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set rsLucra = dbTemp.OpenRecordset("Lucratividade")
 Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
 Set rsClasse = db.OpenRecordset("Classes", , dbReadOnly)
 Set rsSub_Classe = db.OpenRecordset("Sub Classes", , dbReadOnly)

 Data1.DatabaseName = gsQuickDBFileName

 Combo.Text = gnCodFilial

  Rem Pega as tabela usada e joga na lista
  rsPreços.Index = "Só Tabela"
  Lugar = 0
  Últ_Tabela = ""

  Do
    rsPreços.Seek ">", Últ_Tabela
    If Not rsPreços.NoMatch Then
       Últ_Tabela = rsPreços("Tabela")
       Combo_Preço.AddItem Últ_Tabela, Lugar
       Lugar = Lugar + 1
    End If
  Loop Until (rsPreços.NoMatch)

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsProdutos.Close
  rsEstoque.Close
  rsParametros.Close
  rsLucra.Close
  rsPreços.Close
  rsClasse.Close
  rsSub_Classe.Close

  Set rsProdutos = Nothing
  Set rsEstoque = Nothing
  Set rsParametros = Nothing
  Set rsLucra = Nothing
  Set rsPreços = Nothing
  Set rsClasse = Nothing
  Set rsSub_Classe = Nothing
End Sub

Private Function ReturnDevolucaoNormal(ByRef dblValorDevolucao As Double) As Boolean
  Dim strSQL As String
  Dim rstDev As Recordset
'  Dim blnProdutoOK As Boolean
  
  Dim rstProdutos As Recordset
 ' Dim rstGrade As Recordset
  
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
  
'  If Len(Trim(txtNomeCliente.Text)) > 0 Then
'    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboCliente.Text & ") "
'  End If
'
'  If Len(Trim(cboProduto.Text)) > 0 Then
'    strSQL = strSQL & " AND ([Entradas - Produtos].Código = '" & cboProduto.Text & "') "
'  End If
'
'  If Len(Trim(cboClasse.Text)) > 0 Then
'    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
'  End If
'
'  If Len(Trim(cboSubClasse.Text)) > 0 Then
'    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
'  End If
  
  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
'      blnProdutoOK = True
'      If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
'        blnProdutoOK = blnVerificaForncedor(.Fields("Código"))
'      End If
'
'      If blnProdutoOK Then
        dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("PrecoTotal"))
'      End If
    End If
  End With
End Function

Private Function ReturnDevolucaoGrade(ByRef dblValorDevolucao As Double) As Boolean
  Dim strSQL As String
  Dim rstDev As Recordset
'  Dim blnProdutoOK As Boolean
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequência = [Entradas - Produtos].Sequência)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN [Códigos da Grade] ON [Entradas - Produtos].Código = [Códigos da Grade].Código) INNER JOIN Produtos ON [Códigos da Grade].[Código Original] = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Operações Entrada].Tipo)='D')) "


  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(Nome_Empresa.Caption)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & Combo.Text & ") "
  End If
  
'  If Len(Trim(txtNomeCliente.Text)) > 0 Then
'    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboCliente.Text & ") "
'  End If
'
'  If Len(Trim(cboProduto.Text)) > 0 Then
'    strSQL = strSQL & " AND ([Códigos da Grade].[Código Original] = '" & cboProduto.Text & "') "
'  End If
'
'  If Len(Trim(cboClasse.Text)) > 0 Then
'    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
'  End If
'
'  If Len(Trim(cboSubClasse.Text)) > 0 Then
'    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
'  End If
  
  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
'      blnProdutoOK = True
'      If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
'        blnProdutoOK = blnVerificaForncedor(.Fields("Código Original"))
'      End If
'
'      If blnProdutoOK Then
        dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("PrecoTotal"))
'      End If
    End If
  End With
End Function

