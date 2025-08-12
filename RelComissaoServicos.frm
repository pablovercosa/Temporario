VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelComServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comissões de Serviços"
   ClientHeight    =   4620
   ClientLeft      =   3405
   ClientTop       =   2625
   ClientWidth     =   5310
   Icon            =   "RelComissaoServicos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   5310
   Begin VB.Frame Frame5 
      Caption         =   "Exibição do valor"
      Height          =   855
      Left            =   2040
      TabIndex        =   19
      Top             =   2640
      Width           =   3135
      Begin VB.ComboBox cboQtdeCasasDecimais 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Quantidade de casas decimais"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período"
      Height          =   855
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   4935
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3480
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   360
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
         Top             =   360
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Funcionário"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, Código FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Filial"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelComissaoServicos.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
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
      Columns(0).Width=   8361
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   4680
      Top             =   0
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
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   540
         Width           =   1170
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
      Begin VB.OptionButton O_Completo 
         Caption         =   "Completo"
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   555
         Width           =   1065
      End
      Begin VB.OptionButton O_Resumido 
         Caption         =   "Resumido"
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Funcionário 
      Bindings        =   "RelComissaoServicos.frx":059E
      DataSource      =   "Data2"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   960
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
      Columns(0).Width=   7594
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2672
      Columns(1).Caption=   "Apelido"
      Columns(1).Name =   "Apelido"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Apelido"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1588
      Columns(2).Caption=   "Código"
      Columns(2).Name =   "Código"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Código"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   1693
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   300
   End
   Begin VB.Label Nome_Funcionário 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Técnico"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   585
   End
End
Attribute VB_Name = "frmRelComServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsParametros As Recordset
Private rsFuncionarios As Recordset

Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Str_Rel2 As String
 Dim Data1 As Variant
 
 
 Call StatusMsg("")


 If Nome_Filial.Caption = "" Then
   DisplayMsg "Filial incorreta, verifique."
   Combo_Filial.SetFocus
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
 Rel.DataFiles(0) = Str1

 Rem Saída
 If O_Vídeo = True Then Rel.Destination = 0
 If O_Impressora = True Then Rel.Destination = 1
 
 If O_Resumido.Value = True Then Str1 = gsReportPath & "COMSERV1.RPT"
 If O_Completo.Value = True Then Str1 = gsReportPath & "COMSERV2.RPT"
 Rel.ReportFileName = Str1

 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel
 
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
 
 Str_Rel = "{Comissão Serviços.Filial} =" + Combo_Filial.Text
 Str_Rel = Str_Rel + " And {Comissão Serviços.Data} >=" + Str_Data1
 Str_Rel = Str_Rel + " And {Comissão Serviços.Data} <=" + Str_Data2

 

 If Nome_Funcionário.Caption <> "" Then
   Str_Rel = Str_Rel + " And {Comissão Serviços.Vendedor} = " + Combo_Funcionário.Text
' Else
'   Str_Rel = Str_Rel + " And {Comissão Serviços.Vendedor} <> -1"
 End If

 Rel.SelectionFormula = Str_Rel
 
 
 
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Filial.Caption + "'"

 Rel.Formulas(1) = Str_Rel

 Rem data inicial
 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini.Text + "'"
 Rel.Formulas(2) = Str_Rel

 Rem data final
 Str_Rel = "data_fim = '"
 Str_Rel = Str_Rel + Data_Fim.Text + "'"
 Rel.Formulas(3) = Str_Rel

  
  '25/07/2003 - mpdea
  'Fórmula para a quantidade de casas decimais
  'na exibição dos valores de comissão
  Rel.Formulas(4) = "QtdeCasasDecimaisComissao = " & cboQtdeCasasDecimais.Text


 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  
 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub

Private Sub Combo_Filial_CloseUp()
  Combo_Filial.Text = Combo_Filial.Columns(1).Text
  Combo_Filial_LostFocus
End Sub

Private Sub Combo_Filial_LostFocus()
  Nome_Filial.Caption = ""
  If IsNull(Combo_Filial.Text) Then Exit Sub
  If Combo_Filial.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Filial.Text) Then Exit Sub
  If Val(Combo_Filial.Text) < 1 Then Exit Sub
  If Val(Combo_Filial.Text) > 99 Then Exit Sub
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Filial.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Filial.Caption = rsParametros("Nome") & ""
End Sub

Private Sub Combo_Funcionário_CloseUp()
  Combo_Funcionário.Text = Combo_Funcionário.Columns(2).Text
  Combo_Funcionário_LostFocus
End Sub

Private Sub Combo_Funcionário_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Funcionário_LostFocus()

  Call StatusMsg("")
  Nome_Funcionário.Caption = ""
  If IsNull(Combo_Funcionário.Text) Then Exit Sub
  If Combo_Funcionário.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Funcionário.Text) Then Exit Sub
  If Val(Combo_Funcionário.Text) < 1 Then Exit Sub
  If Val(Combo_Funcionário.Text) > 9999 Then Exit Sub
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Val(Combo_Funcionário.Text)
  If rsFuncionarios.NoMatch Then Exit Sub
  
  Nome_Funcionário.Caption = rsFuncionarios("Nome") & ""
 
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
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  
  Combo_Filial.Text = gnCodFilial
  Data_Fim.Text = Format(Date, "dd/mm/yyyy")
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  
  '25/07/2003 - mpdea
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
  
End Sub
