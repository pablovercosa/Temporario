VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelReceber2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contas a Receber por Cliente"
   ClientHeight    =   4035
   ClientLeft      =   3255
   ClientTop       =   2280
   ClientWidth     =   6165
   ForeColor       =   &H80000008&
   Icon            =   "RelReceber2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   6165
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   22
      Top             =   3060
      Width           =   3135
      Begin Threed.SSOption optResumidoOrdenarVendedor 
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Resumido - Vendedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton optResumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optCompleto 
         Caption         =   "Completo"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3885
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Conta"
      Top             =   4785
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3690
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
         Left            =   1080
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
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   21
         Top             =   375
         Width           =   885
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Recebimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   105
      TabIndex        =   17
      Top             =   1950
      Width           =   5970
      Begin VB.OptionButton O_Todos 
         Caption         =   "Todos"
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Carteira 
         Caption         =   "Carteira"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   660
         Width           =   1065
      End
      Begin VB.OptionButton O_Carnet 
         Caption         =   "Carnet"
         Height          =   225
         Left            =   1575
         TabIndex        =   8
         Top             =   660
         Width           =   1065
      End
      Begin VB.OptionButton O_Banco1 
         Caption         =   "Banco"
         Height          =   225
         Left            =   1575
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin SSDataWidgets_B.SSDBCombo Combo_Banco 
         Bindings        =   "RelReceber2.frx":058A
         DataSource      =   "Data3"
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   315
         Width           =   750
         DataFieldList   =   "Descrição"
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
         Columns(0).Width=   6376
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3704
         Columns(1).Caption=   "Conta"
         Columns(1).Name =   "Conta"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Conta"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1720
         Columns(2).Caption=   "Código"
         Columns(2).Name =   "Código"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Código"
         Columns(2).DataType=   2
         Columns(2).FieldLen=   256
         _ExtentX        =   1323
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin VB.Label Nome_Banco 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         TabIndex        =   18
         Top             =   315
         Width           =   2430
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   16
      Top             =   3060
      Width           =   1335
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   1890
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   4815
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4785
      Visible         =   0   'False
      Width           =   1680
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   5520
      Top             =   1155
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelReceber2.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   585
      Width           =   1125
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
      Columns(0).Width=   8943
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2196
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1984
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelReceber2.frx":05B2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   735
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
      Columns(0).Width=   9393
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1561
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Cliente :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   660
      Width           =   615
   End
   Begin VB.Label Nome_Cliente 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2235
      TabIndex        =   14
      Top             =   570
      Width           =   3840
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2235
      TabIndex        =   13
      Top             =   120
      Width           =   3825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   12
      Top             =   225
      Width           =   735
   End
End
Attribute VB_Name = "frmRelReceber2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsCliFor As Recordset
Dim rsContas As Recordset

Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 
 
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


 Rem Verifica fornecedor
 If Nome_Cliente.Caption = "" And Val(Combo_Cliente.Text) <> 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Cliente incorreto, verifique."
   Combo_Cliente.SetFocus
   Exit Sub
 End If


 Rem Verifica Data
 Erro = False
 If IsNull(Data_Ini.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Data_Ini.SetFocus
   Exit Sub
 End If
 
 Rem Verifica Data Final
 Erro = False
 If IsNull(Data_Fim.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Data_Fim.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Data_Fim.SetFocus
   Exit Sub
 End If


 If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
   DisplayMsg "Data inicial deve ser menor ou igual a data final."
   Data_Ini.SetFocus
   Exit Sub
 End If


 Rem  Nome do BD
 Str1 = gsQuickDBFileName
 Rel1.DataFiles(0) = Str1

 Rem Saída
 If B_Vídeo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1

 Rem Nome do arquivo .rpt
 '09/11/2006 - Anderson
 'Alteração do relatório para o cliente 2227727 - (BYBELOT) - VALDIRIA A PEREIRA ME
 '27/02/2007 - Anderson - Desenvolvimento do relatório de contar a receber por cliente, ordenado por vendedor, conforme solicitação do Geferson da NewTech]
 '
 If optCompleto = True Then
  Str1 = gsReportPath & "RECEBE2C.RPT"
 'Else
 ElseIf optResumido = True Then
  Str1 = gsReportPath & "RECEBE2.RPT"
 Else
  Str1 = gsReportPath & "RECEBE2B.RPT"
 End If
 Rel1.ReportFileName = Str1
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 Rem Seleção
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Contas a Receber.Filial} =" + Combo.Text
 Str_Rel = Str_Rel + " And {Contas a Receber.Vencimento} >="
 Str_Rel = Str_Rel + Str_Data1
 Str_Rel = Str_Rel + " And {Contas a Receber.Vencimento} <=" + Str_Data2
 Str_Rel = Str_Rel + " And {Contas a Receber.Valor Recebido} = 0"
 If Nome_Cliente.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Contas a Receber.Cliente} = " + Combo_Cliente.Text
 End If
 Str_Rel = Str_Rel + " And {Contas a Receber.Tipo} = 'R'"


 If O_Carteira.Value = True Then
   Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'C'"
 End If
 If O_Carnet.Value = True Then
   Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'T'"
 End If
 If O_Banco1.Value = True Then
   Str_Rel = Str_Rel + " And {Contas a Receber.Tipo Parcelamento} = 'B'"
   If Nome_Banco.Caption <> "" Then
     Str_Rel = Str_Rel + " And {Contas a Receber.Conta Boleto} = " + str(Combo_Banco.Text)
   End If
 End If


 Rel1.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel1.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel1.Formulas(1) = Str_Rel

 Rem data inicial
 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini.Text + "'"
 Rel1.Formulas(2) = Str_Rel

 Rem data final
 Str_Rel = "data_fim = '"
 Str_Rel = Str_Rel + Data_Fim.Text + "'"
 Rel1.Formulas(3) = Str_Rel

 Rem Taxa de juros
 Str_Rel = "juros = "
 Str_Rel = Str_Rel + str(rsParametros("Juros"))
 Rel1.Formulas(4) = Str_Rel

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub Combo_Banco_CloseUp()
  Combo_Banco.Text = Combo_Banco.Columns(2).Text
  Combo_Banco_LostFocus
End Sub

Private Sub Combo_Banco_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Banco_LostFocus()
  Call StatusMsg("")
  Nome_Banco.Caption = ""
  
  If IsNull(Combo_Banco.Text) Then Exit Sub
  If Combo_Banco.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Banco.Text) Then Exit Sub
  If Val(Combo_Banco.Text) > 9999 Then Exit Sub
  If Val(Combo_Banco.Text) < 1 Then Exit Sub
  
  rsContas.Index = "Código"
  
  rsContas.Seek "=", Val(Combo_Banco.Text)
  If rsContas.NoMatch Then Exit Sub
  
  Nome_Banco.Caption = rsContas("Descrição") & ""

End Sub

Private Sub Combo_Cliente_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Call StatusMsg("")
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub

  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Combo_Cliente.Text
  If Not rsCliFor.NoMatch Then
    Nome_Cliente.Caption = rsCliFor("Nome")
  Else
    Combo_Cliente.Text = 0
  End If

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
  Call CenterForm(Me)
Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
Set rsContas = db.OpenRecordset("Contas Bancárias", , dbReadOnly)

 Data1.DatabaseName = gsQuickDBFileName
 Data2.DatabaseName = gsQuickDBFileName
 Data3.DatabaseName = gsQuickDBFileName

 Combo.Text = gnCodFilial
 
End Sub

Private Sub M_Escolhe_Click()
End Sub


Private Sub O_Banco1_Click()
 Combo_Banco.Enabled = True
 Nome_Banco.Enabled = True
 
End Sub

Private Sub O_Carnet_Click()
 Combo_Banco.Enabled = False
 Nome_Banco.Enabled = False

End Sub

Private Sub O_Carteira_Click()
 Combo_Banco.Enabled = False
 Nome_Banco.Enabled = False

End Sub

Private Sub O_Todos_Click()
 Combo_Banco.Enabled = False
 Nome_Banco.Enabled = False

End Sub
