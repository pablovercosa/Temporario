VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmImprimeCarneCodigoBarras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Impressão de Carnês com Código de Barras"
   ClientHeight    =   4755
   ClientLeft      =   3885
   ClientTop       =   2460
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1650
   Icon            =   "ImprimeCarneCodigoBarras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4755
   ScaleWidth      =   9450
   Begin VB.CommandButton cmd_OutraTelaCarne 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "Sem Código de Barras"
      Height          =   465
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4170
      Width           =   9285
   End
   Begin VB.TextBox txtSequencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1020
      TabIndex        =   2
      Top             =   930
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período"
      Height          =   795
      Left            =   90
      TabIndex        =   21
      Top             =   1350
      Width           =   5475
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   3855
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
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
         Left            =   1215
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   300
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
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2895
         TabIndex        =   23
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir"
      Height          =   795
      Left            =   90
      TabIndex        =   20
      Top             =   2220
      Width           =   5475
      Begin VB.OptionButton O_N_Impresso 
         Appearance      =   0  'Flat
         Caption         =   "Somente os &não impressos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   510
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.OptionButton O_Impresso 
         Appearance      =   0  'Flat
         Caption         =   "Somente os &já impressos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   510
         Width           =   2115
      End
      Begin VB.OptionButton O_Todos 
         Appearance      =   0  'Flat
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções"
      Height          =   795
      Left            =   5640
      TabIndex        =   19
      Top             =   1350
      Width           =   3705
      Begin VB.OptionButton optEsquerda 
         Appearance      =   0  'Flat
         Caption         =   "Esquerda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optDireita 
         Appearance      =   0  'Flat
         Caption         =   "Direita"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   450
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Saída"
      Height          =   795
      Left            =   5640
      TabIndex        =   18
      Top             =   2220
      Width           =   3705
      Begin VB.OptionButton O_Vídeo 
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Impressora 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3630
      Width           =   9285
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
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
      Height          =   345
      Left            =   780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   3870
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
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
      Height          =   345
      Left            =   2340
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3900
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Imprimir"
      Height          =   465
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3090
      Width           =   9285
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Empresa 
      Bindings        =   "ImprimeCarneCodigoBarras.frx":4E95A
      DataSource      =   "Data1"
      Height          =   345
      Left            =   1020
      TabIndex        =   0
      Top             =   90
      Width           =   1245
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
      Columns(0).Width=   7805
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1244
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   2196
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "ImprimeCarneCodigoBarras.frx":4E96E
      DataSource      =   "Data2"
      Height          =   345
      Left            =   1020
      TabIndex        =   1
      ToolTipText     =   "Use 0 para todos os clientes"
      Top             =   495
      Width           =   1245
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
      Columns(0).Width=   8996
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1720
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2196
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   3930
      Top             =   3870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Caption         =   "Sequência"
      Height          =   255
      Left            =   90
      TabIndex        =   24
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2310
      TabIndex        =   17
      Top             =   90
      Width           =   7035
   End
   Begin VB.Label Label6 
      Caption         =   "Filial"
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   165
      Width           =   615
   End
   Begin VB.Label Nome_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2310
      TabIndex        =   15
      Top             =   495
      Width           =   7035
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   225
      Left            =   90
      TabIndex        =   14
      Top             =   555
      Width           =   735
   End
End
Attribute VB_Name = "frmImprimeCarneCodigoBarras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As Recordset
Dim rsClientes As Recordset
Dim rsReceber As Recordset

Private Sub B_Imprime_Click()
  Dim Erro As Integer
  Dim strSQL As String
  Dim strSequencia As String
  Dim lngImpressos As Long
  Dim Str_Rel As String
  Dim Str_Data1 As String
  Dim Str_Data2 As String
   
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  Rem Verifica empresa
  If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
    DisplayMsg "Escolha a empresa."
    Combo_Empresa.SetFocus
    Exit Sub
  End If
  
  If Filial_Liberada <> 0 Then
    If Val(Combo_Empresa.Text) <> Filial_Liberada Then
      DisplayMsg "Funcionário não tem acesso a esta filial."
      Exit Sub
    End If
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
  
  If IsNull(Nome_Cliente.Caption) Then Combo_Cliente.Text = 0
  If Nome_Cliente.Caption = "" Then Combo_Cliente.Text = 0
  
  If IsNull(txtSequencia.Text) Then txtSequencia.Text = 0
  If txtSequencia.Text = "" Then txtSequencia.Text = 0
 
  Call StatusMsg("")
  
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM [Contas a Receber] "
  strSQL = strSQL & "WHERE CarneCodigoBarras <>'' "
  strSQL = strSQL & "  AND Filial = " & Combo_Empresa.Text & " "
  
  If Combo_Cliente.Text <> "0" Then
    strSQL = strSQL & "AND Cliente = " & Combo_Cliente.Text & " "
  End If
  
  If txtSequencia.Text <> "0" Then
    strSQL = strSQL & "AND Sequência = " & txtSequencia.Text & " "
  End If
  
  strSQL = strSQL & "  AND [Data Emissão]>=#" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
  strSQL = strSQL & "  AND [Data Emissão]<=#" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
  
  If O_N_Impresso Then
    strSQL = strSQL & "  AND [Carnet Impresso]= 0 "
  End If
  
  If O_Impresso Then
    strSQL = strSQL & "  AND [Carnet Impresso]<> 0 "
  End If
  
  strSQL = strSQL & " ORDER BY Sequência, Contador"
  
  Set rsReceber = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  strSequencia = ""
  
  With rsReceber
    Do Until .EOF
      If strSequencia <> .Fields("Sequência") Then
        strSequencia = .Fields("Sequência")
        lngImpressos = lngImpressos + 1
      End If
      .Edit
      .Fields("Carnet Impresso").Value = -1
      .Update
      .MoveNext
    Loop
  End With
  
  If lngImpressos > 0 Then
    Rem  Nome do BD
    Rel.DataFiles(0) = gsQuickDBFileName
    
    Rem Saída
    If O_Vídeo = True Then Rel.Destination = 0
    If O_Impressora = True Then Rel.Destination = 1
    
    If optDireita.Value = True Then Rel.ReportFileName = gsReportPath & "CARNEDIREITO.RPT"
    If optEsquerda.Value = True Then Rel.ReportFileName = gsReportPath & "CARNEESQUERDO.RPT"
    
    Rem Seleção
    If O_N_Impresso Then Str_Rel = "{Contas a Receber.Carnet Impresso} = False "
    If O_Impresso Then Str_Rel = "{Contas a Receber.Carnet Impresso} <> False "
    
    Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
    Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")
    
    If Len(Trim(Str_Rel)) <> 0 Then Str_Rel = Str_Rel + " And"
    
    Str_Rel = Str_Rel + " {Contas a Receber.Data Emissão} >=" + Str_Data1
    Str_Rel = Str_Rel + " And {Contas a Receber.Data Emissão} <=" + Str_Data2
    
    If (IsNumeric(Combo_Empresa.Text)) Then
      If CInt(Combo_Empresa.Text) <> 0 Then
        Str_Rel = Str_Rel & " And {Contas a Receber.Filial} = " & Combo_Empresa.Text
      End If
    End If
    
    If Combo_Cliente.Text <> "0" Then
      Str_Rel = Str_Rel & " And {Contas a Receber.Cliente} = " & Combo_Cliente.Text
    End If
    
    If txtSequencia.Text <> "0" Then
      Str_Rel = Str_Rel & " And {Contas a Receber.Sequência} = " & txtSequencia.Text
    End If
    Str_Rel = Str_Rel & " And {Contas a Receber.CarneCodigoBarras} <>'' "
    
    Rel.SelectionFormula = Str_Rel
    
    Call StatusMsg("Aguarde, imprimindo...")
    MousePointer = vbHourglass
     
    '25/07/2003 - mpdea
    'Seta a impressora para relatório
    Call SetPrinterName("CARNÊ", Rel)

    Rel.Action = 1
    
    Call StatusMsg("")
    MousePointer = vbDefault
    
    DisplayMsg "Final de impressão, foram impressos " + str(lngImpressos) + " carnês."
  Else
    DisplayMsg "Não há dados para serem impressos, favor verificar os parâmetros selecionados."
  End If
  
  rsReceber.Close
  Set rsReceber = Nothing
  
  Exit Sub

ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao Imprimir documento."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Private Sub cmd_OutraTelaCarne_Click()
    frmImprimeCarnes.Show
End Sub

Private Sub cmdCancel_Click()
  gbToCancel = True
  Unload Me
End Sub

Private Sub Combo_Cliente_CloseUp()
 Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
 Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub
  If Val(Combo_Cliente.Text) < 0 Or Val(Combo_Cliente.Text) > 99999999 Then Exit Sub

  rsClientes.Index = "Código"
  rsClientes.Seek "=", Val(Combo_Cliente.Text)
  If rsClientes.NoMatch Then Exit Sub
  Nome_Cliente.Caption = rsClientes("Nome")

End Sub

Private Sub Combo_Empresa_CloseUp()
 Combo_Empresa.Text = Combo_Empresa.Columns(1).Text
 Combo_Empresa_LostFocus

End Sub

Private Sub Combo_Empresa_LostFocus()
  Nome_Empresa.Caption = ""
  If IsNull(Combo_Empresa.Text) Then Exit Sub
  If Not IsNumeric(Combo_Empresa.Text) Then Exit Sub
  If Val(Combo_Empresa.Text) < 0 Or Val(Combo_Empresa.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Empresa.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Empresa.Caption = rsParametros("Nome")

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
  Set rsClientes = db.OpenRecordset("CLi_For", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName

  Combo_Empresa = gnCodFilial

End Sub


Private Sub Form_Unload(Cancel As Integer)
  rsParametros.Close
  rsClientes.Close
  
  Set rsParametros = Nothing
  Set rsClientes = Nothing
End Sub
