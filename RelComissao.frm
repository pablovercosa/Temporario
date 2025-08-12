VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelComissoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Comissões"
   ClientHeight    =   7215
   ClientLeft      =   1830
   ClientTop       =   2385
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelComissao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7215
   ScaleWidth      =   11985
   Begin VB.Frame Frame6 
      Caption         =   "Processar Comissões"
      Height          =   975
      Left            =   10170
      TabIndex        =   41
      Top             =   1590
      Width           =   1725
      Begin VB.CommandButton btnProcessar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Processar Comissões"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.TextBox txt_comissaoLiquida 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10380
      TabIndex        =   38
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   6765
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox txt_comissaoBruta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10410
      TabIndex        =   37
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   6915
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox txt_total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7140
      TabIndex        =   34
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   7065
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox txt_totalLiquido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      TabIndex        =   33
      ToolTipText     =   "Digite 0 (zero) para selecionar todas as sequências"
      Top             =   6855
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmd_relatorio_grade2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exibir comissão por vendedor"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2610
      Width           =   3840
   End
   Begin VB.CommandButton cmd_relatorio_grade 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exibir comissão por vendedor x tabela"
      Height          =   435
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2610
      Width           =   3840
   End
   Begin VB.Frame Frame5 
      Caption         =   "Exibição do valor"
      Height          =   945
      Left            =   6810
      TabIndex        =   26
      Top             =   1620
      Width           =   3285
      Begin VB.ComboBox cboQtdeCasasDecimais 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de casas decimais"
         Height          =   195
         Left            =   450
         TabIndex        =   27
         Top             =   270
         Width           =   2160
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar relatório para impressão"
      Height          =   435
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2610
      Width           =   3840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1515
      Left            =   10170
      TabIndex        =   25
      Top             =   60
      Width           =   1695
      Begin VB.OptionButton O_Vídeo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   360
         TabIndex        =   13
         Top             =   390
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1170
      End
   End
   Begin VB.Data datFilial 
      Caption         =   "Filial"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   675
      Left            =   120
      TabIndex        =   20
      Top             =   900
      Width           =   7335
      Begin VB.CommandButton cmd_calendarioDtFim 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5490
         Picture         =   "RelComissao.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   150
         Width           =   465
      End
      Begin VB.CommandButton cmd_calendarioDtIni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         Picture         =   "RelComissao.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   150
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   210
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   210
         Width           =   1260
         _ExtentX        =   2223
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3360
         TabIndex        =   22
         Top             =   255
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   255
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de produto"
      Height          =   1515
      Left            =   7530
      TabIndex        =   19
      Top             =   60
      Width           =   2565
      Begin VB.OptionButton O_Rel_Edição 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Edição"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   390
         TabIndex        =   10
         Top             =   720
         Width           =   840
      End
      Begin VB.OptionButton O_Rel_Grade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Grade"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   390
         TabIndex        =   11
         Top             =   1110
         Width           =   840
      End
      Begin VB.OptionButton O_Rel_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   390
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de relatório"
      Height          =   945
      Left            =   120
      TabIndex        =   18
      Top             =   1620
      Width           =   6645
      Begin VB.CheckBox chkAgruparVendasPorCliente 
         Caption         =   "Resumir"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3540
         TabIndex        =   8
         Top             =   570
         Width           =   975
      End
      Begin VB.OptionButton optClientes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Mostrar Vendas por Clientes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4050
         TabIndex        =   7
         Top             =   330
         Width           =   2535
      End
      Begin VB.OptionButton O_Resumido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Resumido"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   360
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton O_Completo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Completo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2805
         TabIndex        =   6
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton O_Normal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1635
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Funcionário"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4410
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   450
      Visible         =   0   'False
      Width           =   2295
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "RelComissao.frx":4FB1E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   540
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
      Columns.Count   =   3
      Columns(0).Width=   6006
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2990
      Columns(1).Caption=   "Apelido"
      Columns(1).Name =   "Apelido"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Apelido"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1773
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   11610
      Top             =   2340
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
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "RelComissao.frx":4FB32
      DataSource      =   "datFilial"
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   180
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
      Columns(0).Width=   6350
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1005
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
      BackColor       =   12648447
   End
   Begin MSFlexGridLib.MSFlexGrid gridComissaoVendedor 
      Height          =   3885
      Left            =   120
      TabIndex        =   31
      Top             =   3090
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   6853
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   12648384
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   15066597
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Comissão Liquida"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8910
      TabIndex        =   40
      Top             =   6810
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Comissão Bruta"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9090
      TabIndex        =   39
      Top             =   6960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Venda Bruta"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   36
      Top             =   7110
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Venda Liquida"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   35
      Top             =   6900
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Filial"
      Height          =   255
      Left            =   150
      TabIndex        =   24
      Top             =   180
      Width           =   405
   End
   Begin VB.Label lblFilial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   23
      Top             =   180
      Width           =   5085
   End
   Begin VB.Label Label1 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   570
      Width           =   945
   End
   Begin VB.Label Nome_Vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   16
      Top             =   540
      Width           =   5085
   End
End
Attribute VB_Name = "frmRelComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFuncionarios As Recordset

Private Sub B_Imprime_Click()
  Dim Val1, Val2, Erro As Integer
  Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
  Dim Str_Rel As String
  Dim Aux_Data1 As Variant
  
  On Error GoTo TratarErro
  
  Call StatusMsg("")
  
  Erro = False
  If IsNull(Data_Ini.Text) Then Erro = True
  If Erro = False Then If Not IsDate(Data_Ini.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data inválida, verifique."
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  Erro = False
  If IsNull(Data_Fim.Text) Then Erro = True
  If Erro = False Then If Not IsDate(Data_Fim.Text) Then Erro = True
  If Erro = True Then
    DisplayMsg "Data inválida, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  Data_Ini.Text = Format(CDate(Data_Ini.Text), "dd/mm/yyyy")
  Data_Fim.Text = Format(CDate(Data_Fim.Text), "dd/mm/yyyy")
  
  If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
    DisplayMsg "Data final menor que data inicial, verifique."
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  '---[ Gera o total de Descontos do sub-total ]---'
    Dim rstDescSubTotal As Recordset
    Dim curDescSubTotal As Currency
    Dim strSQL          As String
    
    If IsNumeric(cboFilial.Text) Then
      '28/10/2004 - Daniel
      'Adicionado filtro por Vendedor na pesquisa
      If Len(Nome_Vendedor.Caption) <= 0 Then
        strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
                 "Filial = " & CLng(cboFilial.Text) & " AND " & _
                 "Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
                 "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
      Else 'Filtrou Vendedor
        strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
                 "Filial = " & CLng(cboFilial.Text) & " AND Digitador = " & CInt(Combo_Vendedor.Text) & " AND " & _
                 "Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
                 "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
      End If
    Else
      '28/10/2004 - Daniel
      'Adicionado filtro por Vendedor na pesquisa
      If Len(Nome_Vendedor.Caption) <= 0 Then
        strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
                 "Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
                 "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
      Else 'Filtrou Vendedor
        strSQL = "SELECT Sum(DescontoSubTotal) AS Total FROM Saídas WHERE " & _
                 "Digitador = " & CInt(Combo_Vendedor.Text) & _
                 "AND Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & _
                 "# AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#;"
      End If
    End If
    
    Set rstDescSubTotal = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rstDescSubTotal
      Call IsDataType(dtCurrency, .Fields("Total").Value, curDescSubTotal)
      If Not rstDescSubTotal Is Nothing Then .Close
      Set rstDescSubTotal = Nothing
    End With
  '---[ Gera o total de Descontos do sub-total ]---'
  
  
  '08/11/2004 - Daniel
  If chkAgruparVendasPorCliente.Value = vbChecked Then
    'Nome do BD
    Rel.Reset
    Rel.DataFiles(0) = gsTempDBFileName
    Rel.DataFiles(1) = gsQuickDBFileName
    Rel.DataFiles(2) = gsQuickDBFileName
    Rel.DataFiles(3) = gsQuickDBFileName
    'Criamos a tabela temporária
    Call CriarComissaoTemp
  Else
    'Nome do BD
    Rel.Reset
    Str1 = gsQuickDBFileName
    Rel.DataFiles(0) = Str1
  End If
  
  'Saída
  If O_Vídeo = True Then Rel.Destination = 0
  If O_Impressora = True Then Rel.Destination = 1
  
  'Estado da janela
  Rel.WindowState = crptMaximized
  
  '22/06/2005 - Daniel
  'Estava sendo exibido os tipos de relatórios de forma trocada
  'COMISS2.RPT x COMISS3.RPT
  
  'Nome do arquivo .rpt
  If O_Completo.Value = True Then Str1 = gsReportPath & "COMISS1.RPT"
  
  If O_Normal.Value = True Then
    If O_Rel_Grade.Value = True Then Str1 = gsReportPath & "COMISS2G.RPT"
    If O_Rel_Edição.Value = True Then Str1 = gsReportPath & "COMISS2E.RPT"
    If O_Rel_Normal.Value = True Then Str1 = gsReportPath & "COMISS3.RPT"   '"COMISS2.RPT" 22/06/2005 - Daniel
  End If
  
  If O_Resumido.Value = True Then
    Str1 = gsReportPath & "COMISS2.RPT"   '"COMISS3.RPT" 22/06/2005 - Daniel
  End If
  '29/10/2004 - Daniel
  'Adicionado Relatório de Comissão dos Vendedores
  'Ao invés de mostrarmos os produtos, serão mostrados
  'totalizadores por clientes
  If optClientes.Value Then
    If chkAgruparVendasPorCliente.Value = vbChecked Then
      Str1 = gsReportPath & "COMISSCLIRESUMIDO.RPT"
    Else
      Str1 = gsReportPath & "COMISSCLI.RPT"
    End If
  End If
  
  Rel.ReportFileName = Str1
  
  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel
  
  '08/11/2004 - Daniel
  'Quando o rel. de agrupamento das vendas por cliente não
  'for chamado teremos a seleção
  If chkAgruparVendasPorCliente.Value = vbUnchecked Then
    'Seleção
    Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
    Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
    
    If Nome_Vendedor.Caption <> "" Then
      Str_Rel = "{Comissão.Vendedor} = " + Combo_Vendedor.Text + " AND "
    Else
      Str_Rel = ""
    End If
    
    Str_Rel = Str_Rel + " {Comissão.Data} >="
    Str_Rel = Str_Rel + Str_Data1
    Str_Rel = Str_Rel + " And {Comissão.Data} <=" + Str_Data2
     
    If lblFilial.Caption <> "" Then
      Str_Rel = Str_Rel & " And {Comissão.Filial} = " & cboFilial
    End If
     
    Rel.SelectionFormula = Str_Rel
  Else
    Rel.SelectionFormula = ""
  End If
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  
  Rel.Formulas(0) = Str_Rel
  
  'data inicial
  Str_Rel = "dia_ini = '"
  Str_Rel = Str_Rel + Data_Ini.Text + "'"
  Rel.Formulas(1) = Str_Rel
  
  'data final
  Str_Rel = "dia_fim = '"
  Str_Rel = Str_Rel + Data_Fim.Text + "'"
  Rel.Formulas(2) = Str_Rel
  
  '18/07/2003 - mpdea
  'Fórmula para a quantidade de casas decimais
  'na exibição dos valores de comissão
  Rel.Formulas(3) = "QtdeCasasDecimaisComissao = " & cboQtdeCasasDecimais.Text
  Rel.Formulas(4) = "DescSubTotal = " & Replace(curDescSubTotal, ",", ".")
  
  '12/05/2005 - Daniel
  'Correção para exibição dos botões de Configuração
  'de Impressoras e Botão de Pesquisas
  Rel.WindowShowPrintSetupBtn = True
  Rel.WindowShowSearchBtn = True
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  DoEvents
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  Rel.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
  Exit Sub
  
TratarErro:
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub

Private Sub cboFilial_CloseUp()
  lblFilial.Caption = cboFilial.Columns("Nome").Text
  cboFilial.Text = cboFilial.Columns("Filial").Text
End Sub

Private Sub cboFilial_KeyPress(KeyAscii As Integer)
  If Not cboFilial.DroppedDown Then
    KeyAscii = gnLimitKeyPress(cboFilial, 2, KeyAscii, True)
  End If
End Sub

Private Sub cboFilial_LostFocus()
  lblFilial.Caption = gsGetNameFilial(Val(cboFilial.Text))
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub cmd_relatorio_grade_Click()
On Error GoTo Erro

'''  If lblFilial.Caption = "" Then
'''    DisplayMsg "Escolha a filial."
'''    cboFilial.SetFocus
'''    Exit Sub
'''  End If
'''
'''  If Not IsDate(Data_Ini.Text) Then
'''    DisplayMsg "Escolha um período de datas."
'''    Data_Ini.SetFocus
'''    Exit Sub
'''  End If
'''
'''  If Not IsDate(Data_Fim.Text) Then
'''    DisplayMsg "Escolha um período de datas."
'''    Data_Fim.SetFocus
'''    Exit Sub
'''  End If
'''
'''  gridComissaoVendedor.Rows = 1
'''  gridComissaoVendedor.Row = 0
'''
'''  TotalizarCartoes


 
  Dim rsComissao As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long

  If lblFilial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    cboFilial.SetFocus
    Exit Sub
  End If

  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Ini.SetFocus
    Exit Sub
  End If

  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Fim.SetFocus
    Exit Sub
  End If

  gridComissaoVendedor.Rows = 1
  gridComissaoVendedor.Row = 0

  strSQL = strSQL & " SELECT SUM(C.Valor), SUM(C.Comissão),C.tabela, C.Vendedor, F.Nome  "
  strSQL = strSQL & " From Comissão C, Funcionários F "
  strSQL = strSQL & " where C.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " C.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " C.Filial=" & cboFilial.Text

  If Nome_Vendedor.Caption <> "" Then
      strSQL = strSQL & " and C.Vendedor=" & Combo_Vendedor.Text
  End If

  strSQL = strSQL & " and C.Vendedor=F.Código"
  strSQL = strSQL & " GROUP BY C.tabela, C.Vendedor, F.Nome "
  strSQL = strSQL & " ORDER BY 5 "

''''  strSQL = strSQL & " SELECT C.Valor, C.Comissão, C.tabela, C.Vendedor, F.Nome  "
''''  strSQL = strSQL & " From Comissão C, Funcionários F "
''''  strSQL = strSQL & " where C.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
''''  strSQL = strSQL & " C.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
''''  strSQL = strSQL & " C.Filial=" & cboFilial.Text
''''
''''  If Nome_Vendedor.Caption <> "" Then
''''      strSQL = strSQL & " and C.Vendedor=" & Combo_Vendedor.Text
''''  End If
''''
''''  strSQL = strSQL & " and C.Vendedor=F.Código"
''''  strSQL = strSQL & " ORDER BY 4,3 "

  Screen.MousePointer = vbHourglass

  Set rsComissao = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

'  Dim iVendedor As Integer
'  Dim sNomeVendedor As String
'  Dim sTabela As String
'  Dim dTotalVendas As Double
'  Dim dTotalComissao As Double

  If Not (rsComissao.EOF And rsComissao.BOF) Then
    rsComissao.MoveFirst
'    iVendedor = rsComissao.Fields(3).Value
'    sNomeVendedor = rsComissao.Fields(4).Value
'    sTabela = rsComissao.Fields(2).Value
  End If

'  dTotalVendas = 0
'  dTotalComissao = 0
'  lngContadorRegGrid = 1
  While Not rsComissao.EOF

'      If iVendedor = rsComissao.Fields(3).Value And sTabela = rsComissao.Fields(2).Value Then
'        dTotalVendas = dTotalVendas + rsComissao.Fields(0).Value
'        dTotalComissao = dTotalComissao + rsComissao.Fields(1).Value
'      Else
'        gridComissaoVendedor.AddItem iVendedor & vbTab & _
'                      sNomeVendedor & vbTab & _
'                      sTabela & vbTab & _
'                      FormataValorTexto(dTotalVendas, 2) & vbTab & _
'                      FormataValorTexto(dTotalComissao, 2)
'
'        iVendedor = rsComissao.Fields(3).Value
'        sNomeVendedor = rsComissao.Fields(4).Value
'        sTabela = rsComissao.Fields(2).Value
'      End If

      gridComissaoVendedor.AddItem rsComissao.Fields(3).Value & vbTab & _
                      rsComissao.Fields(4).Value & vbTab & _
                      rsComissao.Fields(2).Value & vbTab & _
                      FormataValorTexto(rsComissao.Fields(0).Value, 2) & vbTab & _
                      FormataValorTexto(rsComissao.Fields(1).Value, 2)

      rsComissao.MoveNext
'      lngContadorRegGrid = lngContadorRegGrid + 1
  Wend
  rsComissao.Close
  Set rsComissao = Nothing

  Screen.MousePointer = vbDefault


  Exit Sub
Erro:
  If Not (rsComissao Is Nothing) Then
      rsComissao.Close
      Set rsComissao = Nothing
  End If

  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

'Formata o valor de acordo com o número de casas decimais e substitui separador decimal por ponto
Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
  
End Function

Private Sub cmd_relatorio_grade2_Click()
On Error GoTo Erro

'''''  If lblFilial.Caption = "" Then
'''''    DisplayMsg "Escolha a filial."
'''''    cboFilial.SetFocus
'''''    Exit Sub
'''''  End If
'''''
'''''  If Not IsDate(Data_Ini.Text) Then
'''''    DisplayMsg "Escolha um período de datas."
'''''    Data_Ini.SetFocus
'''''    Exit Sub
'''''  End If
'''''
'''''  If Not IsDate(Data_Fim.Text) Then
'''''    DisplayMsg "Escolha um período de datas."
'''''    Data_Fim.SetFocus
'''''    Exit Sub
'''''  End If
'''''
'''''  gridComissaoVendedor.Rows = 1
'''''  gridComissaoVendedor.Row = 0
'''''
'''''  TotalizarCartoes
 
  Dim rsComissao As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long

  If lblFilial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    cboFilial.SetFocus
    Exit Sub
  End If

  If Not IsDate(Data_Ini.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Ini.SetFocus
    Exit Sub
  End If

  If Not IsDate(Data_Fim.Text) Then
    DisplayMsg "Escolha um período de datas."
    Data_Fim.SetFocus
    Exit Sub
  End If

  gridComissaoVendedor.Rows = 1
  gridComissaoVendedor.Row = 0

  strSQL = strSQL & " SELECT SUM(C.Valor), SUM(C.Comissão),C.Vendedor, F.Nome  "
  strSQL = strSQL & " From Comissão C, Funcionários F "
  strSQL = strSQL & " where C.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " C.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " C.Filial=" & cboFilial.Text

  If Nome_Vendedor.Caption <> "" Then
      strSQL = strSQL & " and C.Vendedor=" & Combo_Vendedor.Text
  End If

  strSQL = strSQL & " and C.Vendedor=F.Código"
  strSQL = strSQL & " GROUP BY C.Vendedor, F.Nome "
  strSQL = strSQL & " ORDER BY 4 "

  Screen.MousePointer = vbHourglass

  Set rsComissao = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

  lngContadorRegGrid = 1

  If Not (rsComissao.EOF And rsComissao.BOF) Then
    rsComissao.MoveFirst
  End If
  While Not rsComissao.EOF
      gridComissaoVendedor.AddItem rsComissao.Fields(2).Value & vbTab & _
                      rsComissao.Fields(3).Value & vbTab & _
                      "" & vbTab & _
                      FormataValorTexto(rsComissao.Fields(0).Value, 2) & vbTab & _
                      FormataValorTexto(rsComissao.Fields(1).Value, 2)

      rsComissao.MoveNext
      lngContadorRegGrid = lngContadorRegGrid + 1
  Wend
  rsComissao.Close
  Set rsComissao = Nothing

  Screen.MousePointer = vbDefault

  Exit Sub
Erro:
  If Not (rsComissao Is Nothing) Then
      rsComissao.Close
      Set rsComissao = Nothing
  End If

  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
  End If

  MsgBox "Erro ao realizar pesquisa...Detalhes do Erro: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub TotalizarCartoes()
  '24/03/2005 - Daniel
  'Rotina que monta o total de cartões
  'por administradora
  Dim rstSaidas            As Recordset
  Dim rstCR                As Recordset
  Dim rstTotalCartoes      As Recordset
  Dim rstTotalCartoesGroup As Recordset
  Dim strSQL               As String
 
  Dim dTotalComissaoBruta As Double
  Dim dTotalComissaoLiquida As Double
  Dim dTotalGrade As Double
  Dim sValorTotalGrade As String
  
  Dim dTotalGradeLiquido As Double
  Dim sValorTotalGradeLiquido As String
 
  dTotalComissaoBruta = 0
  dTotalComissaoLiquida = 0
 
On Error GoTo ErrHandler
 
 dbTemp.Execute "DELETE * FROM TotalCartoes"
 Set rstTotalCartoes = dbTemp.OpenRecordset("TotalCartoes", dbOpenDynaset)
 
 '---[Primeiro buscamos todas às saídas onde houve recebimento com cartão e usou o "caixa escolhido"]---
 strSQL = "SELECT Sequência AS Seq FROM Saídas "
 strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
 strSQL = strSQL & " AND [Recebe - Cartão] <> " & 0
 strSQL = strSQL & " AND Data >= #" & Format(CDate(Data_Ini.Text), "MM/DD/YYYY") & "#"
 strSQL = strSQL & " AND Data <= #" & Format(CDate(Data_Fim.Text), "MM/DD/YYYY") & "#"
 
 If Nome_Vendedor.Caption <> "" Then
    strSQL = strSQL & " AND Digitador = " & Combo_Vendedor.Text
 End If
 
 Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
 
 With rstSaidas
  If Not (.BOF And .EOF) Then
    .MoveFirst
    
    Do Until .EOF
      strSQL = ""
      strSQL = "SELECT * "
      strSQL = strSQL & " FROM [Contas a Receber] "
      strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
      strSQL = strSQL & " AND Sequência = " & rstSaidas.Fields("Seq").Value
      strSQL = strSQL & " AND Tipo = '" & "O" & "'"
      strSQL = strSQL & " AND [Data Emissão] >= #" & Format(CDate(Data_Ini.Text), "MM/DD/YYYY") & "#"
      strSQL = strSQL & " AND [Data Emissão] <= #" & Format(CDate(Data_Fim.Text), "MM/DD/YYYY") & "#"
    
      Set rstCR = db.OpenRecordset(strSQL, dbOpenDynaset)
      
      If Not (rstCR.BOF And rstCR.EOF) Then
        rstCR.MoveFirst
         Do Until rstCR.EOF
            'Criamos o registro temporário
            rstTotalCartoes.AddNew
             rstTotalCartoes.Fields("Sequencia").Value = rstCR.Fields("Sequência").Value
             rstTotalCartoes.Fields("Administradora").Value = rstCR.Fields("Administradora").Value
             rstTotalCartoes.Fields("Vl_Bruto").Value = rstCR.Fields("Valor Cartão").Value
             rstTotalCartoes.Fields("Vl_Liquido").Value = rstCR.Fields("Valor").Value
            rstTotalCartoes.Update
          rstCR.MoveNext
         Loop
      End If
      rstCR.Close
      Set rstCR = Nothing
      
     .MoveNext
    Loop
    
  End If
  .Close
 End With
 
 rstTotalCartoes.Close
 Set rstTotalCartoes = Nothing
 Set rstSaidas = Nothing
 '---[Fim da busca em saídas]---
 
  Dim sPercentualComissao As String
  Dim sPercentualComissaoBruta As String
  Dim sPercentualComissaoLiquida As String
  
  sPercentualComissao = "0"
  If Nome_Vendedor.Caption <> "" Then
      Dim rsFunc As Recordset
      strSQL = "Select * from Funcionários where código = " & Combo_Vendedor.Text
      Set rsFunc = db.OpenRecordset(strSQL, dbOpenDynaset)
      If Not (rsFunc.EOF And rsFunc.BOF) Then
          sPercentualComissao = rsFunc.Fields("Comissão").Value
      End If
      rsFunc.Close
      Set rsFunc = Nothing
  End If
 
 
 'A partir daqui já temos às informações necessárias na tabela temporária TotalCartoes
 'onde poderemos agrupar os registros para a TotalCartoesGroup
 dbTemp.Execute "DELETE * FROM TotalCartoesGroup"
 
 Set rstTotalCartoesGroup = dbTemp.OpenRecordset("TotalCartoesGroup", dbOpenDynaset)
 
 strSQL = ""
 strSQL = "SELECT Administradora, SUM(Vl_Bruto) AS Bruto, SUM(Vl_Liquido) AS Liquido FROM TotalCartoes "
 strSQL = strSQL & " GROUP BY Administradora "
 
 Set rstTotalCartoes = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
 
 With rstTotalCartoes
  If Not (.BOF And .EOF) Then
    .MoveFirst
    
    Do Until .EOF
      rstTotalCartoesGroup.AddNew
        rstTotalCartoesGroup.Fields("Administradora").Value = .Fields("Administradora").Value
        rstTotalCartoesGroup.Fields("Nome").Value = getNomeAdministradora(.Fields("Administradora").Value) & ""
        rstTotalCartoesGroup.Fields("Vl_Bruto").Value = .Fields("Bruto").Value
        rstTotalCartoesGroup.Fields("Vl_Liquido").Value = .Fields("Liquido").Value
      rstTotalCartoesGroup.Update
      
      If sPercentualComissao <> "0" Then
          sPercentualComissaoBruta = .Fields("Bruto").Value
          sPercentualComissaoLiquida = .Fields("Liquido").Value

          sPercentualComissaoBruta = CDbl(sPercentualComissaoBruta) * (CDbl(sPercentualComissao) / 100)
          sPercentualComissaoLiquida = CDbl(sPercentualComissaoLiquida) * (CDbl(sPercentualComissao) / 100)
      Else
          sPercentualComissaoBruta = ""
          sPercentualComissaoLiquida = ""
      End If
      
      If sPercentualComissaoBruta = "" Then
          gridComissaoVendedor.AddItem vbTab & .Fields("Administradora").Value & vbTab & _
                getNomeAdministradora(.Fields("Administradora").Value) & "" & vbTab & _
                FormataValorTexto(.Fields("Bruto").Value, 2) & vbTab & _
                FormataValorTexto(.Fields("Liquido").Value, 2) & vbTab & _
                sPercentualComissaoBruta & vbTab & _
                sPercentualComissaoLiquida
      Else
          gridComissaoVendedor.AddItem vbTab & .Fields("Administradora").Value & vbTab & _
                getNomeAdministradora(.Fields("Administradora").Value) & "" & vbTab & _
                FormataValorTexto(.Fields("Bruto").Value, 2) & vbTab & _
                FormataValorTexto(.Fields("Liquido").Value, 2) & vbTab & _
                FormataValorTexto(sPercentualComissaoBruta, 2) & vbTab & _
                FormataValorTexto(sPercentualComissaoLiquida, 2)
      End If

      sValorTotalGrade = Format(.Fields("Bruto").Value, FORMAT_VALUE)
      dTotalGrade = dTotalGrade + CDbl(sValorTotalGrade)
    
      sValorTotalGradeLiquido = Format(.Fields("Liquido").Value, FORMAT_VALUE)
      dTotalGradeLiquido = dTotalGradeLiquido + CDbl(sValorTotalGradeLiquido)
      
      If sPercentualComissaoBruta <> "" Then
          sPercentualComissaoBruta = Format(sPercentualComissaoBruta, FORMAT_VALUE)
          dTotalComissaoBruta = dTotalComissaoBruta + CDbl(sPercentualComissaoBruta)
          
          sPercentualComissaoLiquida = Format(sPercentualComissaoLiquida, FORMAT_VALUE)
          dTotalComissaoLiquida = dTotalComissaoLiquida + CDbl(sPercentualComissaoLiquida)
      End If
  
     .MoveNext
    Loop
    
  End If
  .Close
 End With
 
  txt_total.Text = Format(dTotalGrade, FORMAT_VALUE)
  txt_totalLiquido.Text = Format(dTotalGradeLiquido, FORMAT_VALUE)
 
  txt_comissaoBruta.Text = Format(dTotalComissaoBruta, FORMAT_VALUE)
  txt_comissaoLiquida.Text = Format(dTotalComissaoLiquida, FORMAT_VALUE)
 
 
 Set rstTotalCartoes = Nothing
 
 rstTotalCartoesGroup.Close
 Set rstTotalCartoesGroup = Nothing
 
 Exit Sub
 
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  Exit Sub
  
End Sub

Private Function getNomeAdministradora(ByVal CodAdmi As String) As String
  '23/03/2005 - Daniel
  Dim rstCartoes As Recordset
  
  Set rstCartoes = db.OpenRecordset("SELECT Nome FROM Cartões WHERE Código = " & CodAdmi, dbOpenDynaset)
  
  With rstCartoes
    If Not (.BOF And .EOF) Then
      .MoveFirst
      getNomeAdministradora = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstCartoes = Nothing
  
End Function

Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Vendedor_LostFocus()
  Call StatusMsg("")
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub
  If Val(Combo_Vendedor.Text) > 9999 Then Exit Sub
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Vendedor.Text)
  If rsFuncionarios.NoMatch Then Exit Sub
  Nome_Vendedor.Caption = rsFuncionarios("Apelido")
End Sub

Private Sub btnProcessar_Click()
  If lblFilial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    cboFilial.SetFocus
    Exit Sub
  End If

  If Trim(Replace(Data_Ini.Text, "/", "")) = "" Then
    DisplayMsg "Insira uma data inicial"
    Data_Ini.SetFocus
    Exit Sub
  End If
  
  If Trim(Replace(Data_Fim.Text, "/", "")) = "" Then
    DisplayMsg "Insira uma data final"
    Data_Fim.SetFocus
    Exit Sub
  End If
  
  If DateDiff("d", Data_Ini.Text, Data_Fim.Text) > 62 Then
    DisplayMsg "Intervalo (" & DateDiff("d", Data_Ini.Text, Data_Fim.Text) & ") maior que 62 dias"
    Exit Sub
  End If
  
  Dim intervalo As String
  intervalo = Data_Ini.Text & " a " & Data_Fim.Text
  If MsgBox("Todas as comissões no intervalo (" & intervalo & ") serão calculadas", vbYesNo, "ATENÇÃO") = vbYes Then
    Dim sql As String
    sql = "SELECT DISTINCT"
    sql = sql & "  s.Data"
    sql = sql & ", s.Digitador AS Vendedor1"
    sql = sql & ", s.PrestadorServico AS Vendedor2"
    sql = sql & ", sp.Código AS Produto"
    sql = sql & ", sp.Qtde"
    sql = sql & ", sp.[Preço Final]-s.[Recebe - Vale] AS Valor"
    sql = sql & ", (s.Desconto+CDbl(Trim(s.DescontoSubTotal))*100)/s.Produtos AS PorcentagemDesconto"
    sql = sql & ", (sp.[Preço Final]-s.[Recebe - Vale])-((sp.[Preço Final]-s.[Recebe - Vale])*(((s.Desconto+CDbl(Trim(s.DescontoSubTotal))*100)/s.Produtos)/100)) AS ValorTotal"
    sql = sql & ", f1.Comissão AS Comissao1"
    sql = sql & ", f2.Comissão AS Comissao2"
    sql = sql & ", p.[Comissão Sobrepõe] AS ComProdSobrepoe"
    sql = sql & ", p.Comissão AS ComissProduto"
    sql = sql & ", tp.[Multiplicador Comissão] AS Multiplicador"
    sql = sql & ", sp.Sequência"
    sql = sql & ", s.Cliente"
    sql = sql & ", s.Tabela"
    sql = sql & ", s.Filial"
    sql = sql & ", ((sp.[Preço Final]*(s.[Recebe - Cartão]*100/s.[Total]))/100) AS VlPagoEmCartao"
    sql = sql & ", ((sp.[Preço Final]*(s.[Recebe - Cartão]*100/s.[Total]))/100)-((((sp.[Preço Final]*(s.[Recebe - Cartão]*100/s.[Total]))/100)*c.Taxa)/100) AS VlPagoEmCartaoComRetencao"
    sql = sql & ", c.Taxa AS TaxaRetencao"
    sql = sql & ", s.Total"
    sql = sql & " FROM"
    sql = sql & " [Operações Saída] AS o"
    sql = sql & " INNER JOIN (Cartões AS c"
    sql = sql & " RIGHT JOIN (([Tabela de Preços] AS tp"
    sql = sql & " INNER JOIN ((Saídas AS s"
    sql = sql & " INNER JOIN ([Saídas - Produtos] AS sp"
    sql = sql & " INNER JOIN Produtos AS p ON sp.Código = p.Código) ON (s.Sequência = sp.Sequência) AND (s.Filial = sp.Filial))"
    sql = sql & " INNER JOIN Funcionários AS f1 ON s.Digitador = f1.Código) ON tp.Tabela = s.Tabela)"
    sql = sql & " LEFT JOIN Funcionários AS f2 ON s.PrestadorServico = f2.Código) ON c.Código = s.[Recebe - Emp Cartão]) ON o.Código = s.Operação"
    sql = sql & " WHERE"
    sql = sql & " (((s.Data) Between #" & Format(Data_Ini.Text, "m/d/yyyy") & "# And #" & Format(Data_Fim.Text, "m/d/yyyy") & "#)"
    sql = sql & " AND ((sp.Sequência) Not In (SELECT DISTINCT cm.Sequência FROM Comissão AS cm WHERE cm.Filial = s.Filial))"
    sql = sql & " AND ((s.Filial)=" & cboFilial.Text & ")"
    sql = sql & " AND ((s.Efetivada)=True)"
    sql = sql & " AND ((s.[Movimentação Desfeita])=False)"
    sql = sql & " AND ((tp.[Multiplicador Comissão])<>0)"
    sql = sql & " AND ((o.Comissão)=True))"
    sql = sql & " ORDER BY s.Data, sp.Sequência, sp.Código;"
    
    Dim rsComissoes As Recordset
    Set rsComissoes = db.OpenRecordset(sql, dbOpenDynaset)
    
    Dim objComissao As clsComissao
    Dim vendedores As Collection
    Dim porcentagem As Double
    Dim cont As Integer
    
    With rsComissoes
      Do While Not .EOF
        cont = cont + 1
        Set vendedores = New Collection
        vendedores.Add .Fields("Vendedor1").Value
        If Not IsNull(.Fields("Vendedor2").Value) Then
          If Val(.Fields("Vendedor2").Value) > 0 Then vendedores.Add .Fields("Vendedor2").Value
        End If
        Dim i As Integer
        For i = 1 To vendedores.Count
          porcentagem = IIf(.Fields("ComProdSobrepoe").Value, .Fields("ComissProduto").Value, .Fields("Comissao" & i).Value)
          porcentagem = porcentagem * .Fields("Multiplicador").Value
          
          Set objComissao = New clsComissao
          objComissao.Data = .Fields("Data").Value
          objComissao.Vendedor = .Fields("Vendedor" & i).Value
          objComissao.Produto = .Fields("Produto").Value
          objComissao.Tamanho = 0
          objComissao.Cor = 0
          objComissao.Edição = 0
          objComissao.Qtde = .Fields("Qtde").Value
          objComissao.Valor = Round(.Fields("ValorTotal").Value, 2)
          objComissao.Comissão = Round(porcentagem * objComissao.Valor / 100, 2)
          objComissao.Sequência = .Fields("Sequência").Value
          objComissao.Cliente = .Fields("Cliente").Value
          objComissao.Tabela = .Fields("Tabela").Value
          objComissao.Filial = .Fields("Filial").Value
          objComissao.VlPagoEmCartao = .Fields("VlPagoEmCartao").Value
          objComissao.VlPagoEmCartaoComRetencao = IIf(IsNull(.Fields("VlPagoEmCartaoComRetencao").Value), 0, .Fields("VlPagoEmCartaoComRetencao").Value)
          objComissao.TaxaRetencao = IIf(IsNull(.Fields("TaxaRetencao").Value), 0, .Fields("TaxaRetencao").Value)
          objComissao.Insert
          Set objComissao = Nothing
        Next
        Set vendedores = Nothing
        .MoveNext
      Loop
      .Close
    End With
    Set rsComissoes = Nothing
    DisplayMsg "Comissões calculadas e salvas com sucesso!"
  End If
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
  Call CenterForm(Me)
  
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  datFilial.DatabaseName = gsQuickDBFileName
  
  cboFilial.Text = gnCodFilial
  Data_Fim.Text = Format(Date, "dd/mm/yyyy")

  If gbGrade = False Then O_Rel_Grade.Enabled = False
  If gbEdicao = False Then O_Rel_Edição.Enabled = False
  
  
  '21/07/2003 - mpdea
  'Preenche combo com a quantidade de casas decimais para
  'exibição do valor de comissão
  With cboQtdeCasasDecimais
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .ListIndex = 0
  End With
  
'''  gridComissaoVendedor.ColWidth(0) = 0
'''  gridComissaoVendedor.ColWidth(1) = 1500
'''  gridComissaoVendedor.ColWidth(2) = 3500
'''  gridComissaoVendedor.ColWidth(3) = 1600
'''  gridComissaoVendedor.ColWidth(4) = 1600
'''  gridComissaoVendedor.ColWidth(5) = 1600
'''  gridComissaoVendedor.ColWidth(6) = 1600
'''
'''  gridComissaoVendedor.Row = 0
'''  gridComissaoVendedor.TextMatrix(0, 0) = ""
'''  gridComissaoVendedor.TextMatrix(0, 1) = "Código"
'''  gridComissaoVendedor.TextMatrix(0, 2) = "Nome"
'''  gridComissaoVendedor.TextMatrix(0, 3) = "Venda Bruta"
'''  gridComissaoVendedor.TextMatrix(0, 4) = "Venda Liquida"
'''  gridComissaoVendedor.TextMatrix(0, 5) = "Comissão Bruta"
'''  gridComissaoVendedor.TextMatrix(0, 6) = "Comissão Liquida"

  gridComissaoVendedor.ColWidth(0) = 1000
  gridComissaoVendedor.ColWidth(1) = 4500
  gridComissaoVendedor.ColWidth(2) = 2500
  gridComissaoVendedor.ColWidth(3) = 1700
  gridComissaoVendedor.ColWidth(4) = 1700
    
  gridComissaoVendedor.Row = 0
  gridComissaoVendedor.TextMatrix(0, 0) = "Código"
  gridComissaoVendedor.TextMatrix(0, 1) = "Nome"
  gridComissaoVendedor.TextMatrix(0, 2) = "Tabela"
  gridComissaoVendedor.TextMatrix(0, 3) = "Total Venda"
  gridComissaoVendedor.TextMatrix(0, 4) = "Total Comissão"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub O_Completo_Click()
  O_Rel_Normal.Value = True
  O_Rel_Grade.Enabled = False
  O_Rel_Edição.Enabled = False
  chkAgruparVendasPorCliente.Value = vbUnchecked
  chkAgruparVendasPorCliente.Enabled = False
End Sub

Private Sub O_Normal_Click()
  O_Rel_Grade.Enabled = True
  O_Rel_Edição.Enabled = True
  chkAgruparVendasPorCliente.Value = vbUnchecked
  chkAgruparVendasPorCliente.Enabled = False
End Sub

Private Sub O_Resumido_Click()
  O_Rel_Normal.Value = True
  O_Rel_Grade.Enabled = False
  O_Rel_Edição.Enabled = False
  chkAgruparVendasPorCliente.Value = vbUnchecked
  chkAgruparVendasPorCliente.Enabled = False
End Sub

'Private Sub O_Serviços_Click()
'  O_Rel_Normal.Value = True
'  O_Rel_Grade.Enabled = False
'  O_Rel_Edição.Enabled = False
'End Sub
Private Sub optClientes_Click()
  chkAgruparVendasPorCliente.Enabled = True
End Sub

Private Sub CriarComissaoTemp()
  Dim rstComissao As Recordset
  Dim rstTemp     As Recordset
  Dim strSQL      As String
  
  dbTemp.Execute "DELETE * FROM Comissao"
  
  strSQL = "SELECT Data, Vendedor, Cliente, Filial, SUM(Valor) AS ValorTot, SUM(Comissão) AS ComissaoTot "
  strSQL = strSQL & " FROM Comissão "
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Data >= #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Data <= #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#"
  
  If Len(Nome_Vendedor.Caption) > 0 Then strSQL = strSQL & " AND Vendedor = " & CInt(Combo_Vendedor.Text)
  
  strSQL = strSQL & " GROUP BY Data, Vendedor, Cliente, Filial "
  
  Set rstComissao = db.OpenRecordset(strSQL, dbOpenDynaset)
  Set rstTemp = dbTemp.OpenRecordset("Comissao", dbOpenDynaset)
  
  With rstComissao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        rstTemp.AddNew
         rstTemp.Fields("Data").Value = .Fields("Data").Value
         rstTemp.Fields("Vendedor").Value = .Fields("Vendedor").Value
         rstTemp.Fields("Valor").Value = Format(.Fields("ValorTot").Value, "#0.00")
         rstTemp.Fields("Comissao").Value = Format(.Fields("ComissaoTot").Value, "#0.00")
         rstTemp.Fields("Cliente").Value = .Fields("Cliente").Value
         rstTemp.Fields("Filial").Value = .Fields("Filial").Value
        rstTemp.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstComissao = Nothing
  
  rstTemp.Close
  Set rstTemp = Nothing
  
End Sub

