VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelServicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Serviços"
   ClientHeight    =   3360
   ClientLeft      =   3645
   ClientTop       =   2730
   ClientWidth     =   6045
   Icon            =   "RelServicos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   6045
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      Height          =   795
      Left            =   90
      TabIndex        =   18
      Top             =   1545
      Width           =   5145
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3690
         TabIndex        =   4
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
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2820
         TabIndex        =   20
         Top             =   375
         Width           =   885
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3915
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   4305
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2055
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Serviço"
      Top             =   4290
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   4275
      Visible         =   0   'False
      Width           =   1830
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "RelServicos.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   210
      Width           =   750
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
      _ExtentX        =   1323
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelServicos.frx":059E
      DataSource      =   "Data3"
      Height          =   315
      Left            =   1035
      TabIndex        =   2
      ToolTipText     =   "Use 0 para todos os clientes."
      Top             =   1065
      Width           =   1275
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
      Columns(0).Width=   8678
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
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   5415
      Top             =   1635
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
      Height          =   400
      Left            =   4575
      TabIndex        =   9
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      Height          =   855
      Left            =   1680
      TabIndex        =   13
      Top             =   2415
      Width           =   1485
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   570
         Width           =   1170
      End
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   255
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   105
      TabIndex        =   12
      Top             =   2415
      Width           =   1380
      Begin VB.OptionButton O_Completo 
         Caption         =   "Completo"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   555
         Width           =   1065
      End
      Begin VB.OptionButton O_Resumido 
         Caption         =   "Resumido"
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Serviço 
      Bindings        =   "RelServicos.frx":05B2
      DataSource      =   "Data2"
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todos os serviços."
      Top             =   630
      Width           =   960
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
      Columns.Count   =   2
      Columns(0).Width=   8467
      Columns(0).Caption=   "Descrição"
      Columns(0).Name =   "Descrição"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Descrição"
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
      _ExtentX        =   1693
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Nome_Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2505
      TabIndex        =   17
      Top             =   210
      Width           =   3465
   End
   Begin VB.Label Label7 
      Caption         =   "Filial :"
      Height          =   225
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   645
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2505
      TabIndex        =   15
      Top             =   1065
      Width           =   3465
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente :"
      Height          =   225
      Left            =   90
      TabIndex        =   14
      Top             =   1095
      Width           =   750
   End
   Begin VB.Label Nome_Serviço 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2505
      TabIndex        =   11
      Top             =   630
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "Serviço :"
      Height          =   225
      Left            =   105
      TabIndex        =   10
      Top             =   660
      Width           =   750
   End
End
Attribute VB_Name = "frmRelServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsParametros As Recordset
Private rsServicos As Recordset
Private rsClientes As Recordset

Private Sub B_Imprime_Click()
  Dim Erro As Integer
  Dim Str1 As String, Str_Data1 As String, Str_Data2 As String
  Dim Str_Rel As String
 
 
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
  
  If O_Resumido.Value = True Then
    Str1 = gsReportPath & "SERVICO1.RPT"
  End If
  
  If O_Completo.Value = True Then
    Str1 = gsReportPath & "SERVICO2.RPT"
  End If
  
  
  Rel.ReportFileName = Str1

  ' Modelo 1 ou 2
  'SetPrinterModeloPwd2 Rel
 
  Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
  Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
  
  Str_Rel = "{Comissão Serviços.Filial} =" + Combo_Filial.Text
  Str_Rel = Str_Rel + " And {Comissão Serviços.Data} >=" + Str_Data1
  Str_Rel = Str_Rel + " And {Comissão Serviços.Data} <=" + Str_Data2
  
  
  If Nome_Serviço.Caption <> "" Then
    Str_Rel = Str_Rel + " And {Comissão Serviços.Serviço} = " + Combo_Serviço.Text
  End If
  
  If Nome_Cliente.Caption <> "" Then
    Str_Rel = Str_Rel + " And {Comissão Serviços.Cliente} = " + Combo_Cliente.Text
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
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  
  Rel.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault

End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Cliente_LostFocus()
  Call StatusMsg("")
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Combo_Cliente.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 1 Then Exit Sub
  If Val(Combo_Cliente.Text) > 99999999 Then Exit Sub
  
  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  
  Nome_Cliente.Caption = rsClientes("Nome") & ""
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

Private Sub Combo_Serviço_CloseUp()
  Combo_Serviço.Text = Combo_Serviço.Columns(1).Text
  Combo_Serviço_LostFocus
End Sub

Private Sub Combo_Serviço_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Serviço_LostFocus()
  Call StatusMsg("")
  Nome_Serviço.Caption = ""
  If IsNull(Combo_Serviço.Text) Then Exit Sub
  If Combo_Serviço.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Serviço.Text) Then Exit Sub
  If Val(Combo_Serviço.Text) < 1 Then Exit Sub
  If Val(Combo_Serviço.Text) > 9999 Then Exit Sub
  
  rsServicos.Index = "Código"
  rsServicos.Seek "=", Val(Combo_Serviço.Text)
  If rsServicos.NoMatch Then Exit Sub
  
  Nome_Serviço.Caption = rsServicos("Descrição") & ""
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
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsServicos = db.OpenRecordset("Serviços", , dbReadOnly)
  
  
  Combo_Filial.Text = gnCodFilial
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
 
End Sub
