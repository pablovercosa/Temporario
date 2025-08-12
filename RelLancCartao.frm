VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelLancCartao 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Lançamentos de Cartão de Crédito"
   ClientHeight    =   3630
   ClientLeft      =   1845
   ClientTop       =   1455
   ClientWidth     =   8445
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
   Icon            =   "RelLancCartao.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   8445
   Begin VB.Data datCartoes 
      Caption         =   "datCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cartões ORDER BY Código"
      Top             =   3540
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3030
      Width           =   8205
   End
   Begin VB.Frame fraR 
      Caption         =   "Tipo do Relatório"
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
      TabIndex        =   19
      Top             =   1680
      Width           =   3975
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         Height          =   285
         Left            =   1710
         TabIndex        =   8
         Top             =   315
         Width           =   1095
      End
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opção"
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
      Left            =   4170
      TabIndex        =   18
      Top             =   900
      Width           =   4155
      Begin VB.OptionButton optReceber 
         Caption         =   "A Receber"
         Height          =   255
         Left            =   2670
         TabIndex        =   6
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optRecebidos 
         Caption         =   "Recebidos"
         Height          =   255
         Left            =   1410
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   330
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   16
      Top             =   900
      Width           =   3975
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   2370
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   270
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
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
         Left            =   540
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   270
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   22
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   330
         Width           =   255
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
      Height          =   735
      Left            =   4170
      TabIndex        =   15
      Top             =   1680
      Width           =   4155
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton B_Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   390
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   8205
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   3540
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   3840
      Top             =   60
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
      Bindings        =   "RelLancCartao.frx":4E95A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   180
      Width           =   1005
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
      BevelColorFrame =   -2147483633
      BevelColorHighlight=   -2147483633
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   7752
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1879
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1773
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
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
   Begin SSDataWidgets_B.SSDBCombo cboCartoes 
      Bindings        =   "RelLancCartao.frx":4E96E
      DataSource      =   "datCartoes"
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   540
      Width           =   1005
      DataFieldList   =   "Código"
      _Version        =   196617
      BevelColorFrame =   -2147483633
      BevelColorHighlight=   -2147483633
      BackColorOdd    =   12648447
      Columns(0).Width=   3200
      _ExtentX        =   1773
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
      DataFieldToDisplay=   "Código"
   End
   Begin Crystal.CrystalReport RelSintetico 
      Left            =   4200
      Top             =   3540
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
   Begin VB.Label lblNomeCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   21
      Top             =   540
      Width           =   6405
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Cartão"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Nome_Empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   180
      Width           =   6405
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   300
   End
End
Attribute VB_Name = "frmRelLancCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsParametros As Recordset

Private Sub B_Imprime_Click()
 Dim Val1, Val2, Erro As Integer
 Dim Str1, Str2, Str3, Str_Data1, Str_Data2 As String
 Dim Str_Rel As String
 Dim Data1 As Variant
 
 On Error GoTo TrataErro
 
 Call StatusMsg("")

 'Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a empresa."
   Combo.SetFocus
   Exit Sub
 End If

 'Verifica funcionário
 If Filial_Liberada <> 0 Then
   If Val(Combo.Text) <> Filial_Liberada Then
     DisplayMsg "Funcionário não tem acesso a esta filial."
     Exit Sub
   End If
 End If

 'Verifica Data Inicial
 Erro = False
 If IsNull(Data_Ini.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Data_Ini.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Data_Ini.SetFocus
   Exit Sub
 End If
 
 'Verifica Data Final
 Erro = False
 If IsNull(Data_Fim.Text) Then Erro = True
 If Not Erro Then If Not IsDate(Data_Fim.Text) Then Erro = True
 If Erro = True Then
   DisplayMsg "Data incorreta, verifique."
   Data_Fim.SetFocus
   Exit Sub
 End If

 'Validações das datas
 If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
   DisplayMsg "Data inicial deve ser menor ou igual a data final."
   Data_Ini.SetFocus
   Exit Sub
 End If

 '29/03/2005 - Daniel
 '
 'Desenvolvido relatório sintético para atender
 'inicialmente às necessidades do cliente Bem Me Quer
 If optSintetico.Value Then
  Call StatusMsg("Montando relatório sintético, aguarde...")
  Screen.MousePointer = vbHourglass
  Call MontarRelSintetico
  Call ExibirRelSintetico
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  
  Exit Sub
 End If
 '----------------------------------------------------------
 
 'Seta Valores e Manda Relatório

 'Nome do BD
 Str1 = gsQuickDBFileName
 Rel.DataFiles(0) = Str1

 'Saída
 If B_Vídeo = True Then Rel.Destination = 0
 If B_Impressora = True Then Rel.Destination = 1
 Rem If B_Arquivo = True Then
 Rem    frmMenu.Relatório.Destination = 2
 Rem    frmMenu.Relatório.PrintFileName = T_Arquivo.Text
 Rem End If

 'Nome do arquivo .rpt
 Str1 = gsReportPath & "LANCA1.RPT"
 Rel.ReportFileName = Str1

 'Seleção
 Str_Data1 = "Date" + Format$(Data_Ini.Text, "(yyyy,mm,dd)")
 Str_Data2 = "Date" + Format$(Data_Fim.Text, "(yyyy,mm,dd)")

 Str_Rel = "{Contas a Receber.Filial} =" + Combo.Text
 Str_Rel = Str_Rel + " And {Contas a Receber.Vencimento} >="
 Str_Rel = Str_Rel + Str_Data1
 Str_Rel = Str_Rel + " And {Contas a Receber.Vencimento} <=" + Str_Data2
 Str_Rel = Str_Rel + " And {Contas a Receber.Tipo} = 'O' "
 '29/03/2005 - Daniel
 'Adicionado filtro administradora
 If Len(lblNomeCartao.Caption) > 0 Then Str_Rel = Str_Rel + " And {Contas a Receber.Administradora} = " & CByte(Trim(cboCartoes.Text))
 
 '22/10/2003 - Maikel
 '---[ Adicionado filtro para cartões recebidos e a receber ]---'
  If optReceber.Value Then
    Str_Rel = Str_Rel & " AND {Contas a Receber.Valor Recebido} < {Contas a Receber.Valor}"
  End If
  
  If optRecebidos.Value Then
    Str_Rel = Str_Rel & " AND {Contas a Receber.Valor Recebido} >= {Contas a Receber.Valor}"
  End If
 '---[ Adicionado filtro para cartões recebidos e a receber ]---'
 
 Rel.SelectionFormula = Str_Rel
 
 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"

 Rel.Formulas(0) = Str_Rel

 Str_Rel = "nome_filial = '"
 Str_Rel = Str_Rel + Nome_Empresa.Caption + "'"
 Rel.Formulas(1) = Str_Rel


 'data inicial
 Str_Rel = "data_ini = '"
 Str_Rel = Str_Rel + Data_Ini.Text + "'"
 Rel.Formulas(2) = Str_Rel

 'data final
 Str_Rel = "data_fim = '"
 Str_Rel = Str_Rel + Data_Fim.Text + "'"
 Rel.Formulas(3) = Str_Rel

 'Opção escolhida
 If optTodos.Value Then Rel.Formulas(4) = "Opcao = '" & "Opção Escolhida: Todos (Recebidos e A Receber)" & "'"
 If optRecebidos.Value Then Rel.Formulas(4) = "Opcao = '" & "Opção Escolhida: Recebidos" & "'"
 If optReceber.Value Then Rel.Formulas(4) = "Opcao = '" & "Opção Escolhida: A Receber" & "'"

 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)

 Rel.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

 Exit Sub
 
TrataErro:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Sub cmdFechar_Click()
  Unload Me
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
  Data1.DatabaseName = gsQuickDBFileName
  datCartoes.DatabaseName = gsQuickDBFileName
  
  Combo.Text = gnCodFilial
End Sub

Private Sub cboCartoes_CloseUp()
  cboCartoes.Text = cboCartoes.Columns(0).Text
  cboCartoes_LostFocus
End Sub

Private Sub cboCartoes_LostFocus()
  Dim rstCartoes As Recordset
  
  lblNomeCartao.Caption = ""
  If Not IsNumeric(cboCartoes.Text) Then Exit Sub
  
  Set rstCartoes = db.OpenRecordset("SELECT Código, Nome FROM Cartões WHERE Código = " & cboCartoes.Text, dbOpenSnapshot)
  
  With rstCartoes
    If Not (.BOF And .EOF) Then
      lblNomeCartao.Caption = .Fields("Nome") & ""
    End If
    
    If Not rstCartoes Is Nothing Then .Close
    Set rstCartoes = Nothing
  End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
 rsParametros.Close
 Set rsParametros = Nothing
End Sub

Private Sub MontarRelSintetico()
  '29/03/2005 - Daniel
  'Desenvolvido relatório sintético para atender
  'inicialmente às necessidades do cliente Bem Me Quer
  Dim rstTotalCartoesGroup As Recordset
  Dim rstContasReceber     As Recordset
  Dim strQuery             As String
  
  On Error GoTo TratarErro
    'Tratamento para a table temporária
    dbTemp.Execute "DELETE * FROM TotalCartoesGroup"
    Set rstTotalCartoesGroup = dbTemp.OpenRecordset("TotalCartoesGroup", dbOpenDynaset)
    'Fim Tratamento
    
    'Busca no Contas a Receber
    strQuery = "SELECT Administradora, SUM([Valor Cartão]) AS Bruto, SUM(Valor) AS Liquido "
    strQuery = strQuery & " FROM [Contas a Receber] "
    strQuery = strQuery & " WHERE Filial =" + Combo.Text
    strQuery = strQuery & " AND Tipo = '" & "O" & "'"
    strQuery = strQuery & " AND Vencimento >= #" & Format(Data_Ini.Text, "MM/DD/YYYY") & "#"
    strQuery = strQuery & " AND Vencimento <= #" & Format(Data_Fim.Text, "MM/DD/YYYY") & "#"
    'Filtro Administradora
    If Len(lblNomeCartao.Caption) > 0 Then strQuery = strQuery & " AND Administradora = " & CByte(Trim(cboCartoes.Text))
    'Recebidos
    If optRecebidos.Value Then strQuery = strQuery & " AND [Valor Recebido] >= [Valor] "
    'A Receber
    If optReceber.Value Then strQuery = strQuery & " AND [Valor Recebido] < [Valor] "
  
    strQuery = strQuery & " GROUP BY Administradora "
  
    Set rstContasReceber = db.OpenRecordset(strQuery, dbOpenDynaset)
  
    With rstContasReceber
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do Until .EOF
          rstTotalCartoesGroup.AddNew
            rstTotalCartoesGroup.Fields("Administradora").Value = .Fields("Administradora").Value
            rstTotalCartoesGroup.Fields("Nome").Value = getNomeAdministradora(.Fields("Administradora").Value) & ""
            rstTotalCartoesGroup.Fields("Vl_Bruto").Value = .Fields("Bruto").Value
            rstTotalCartoesGroup.Fields("Vl_Liquido").Value = .Fields("Liquido").Value
          rstTotalCartoesGroup.Update
        
         .MoveNext
        Loop
        
      End If
      .Close
    End With
    
    Set rstContasReceber = Nothing
  
    rstTotalCartoesGroup.Close
    Set rstTotalCartoesGroup = Nothing
  
  Exit Sub

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub

Private Function getNomeAdministradora(ByVal CodAdmi As String) As String
  '30/03/2005 - Daniel
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

Private Sub ExibirRelSintetico()
  '30/03/2005 - Daniel
  Dim strReport As String
  
  On Error GoTo TratarErro
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptCartoesCreditoSintetico.rpt"
  
  With RelSintetico
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 RelSintetico
    
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "Periodo = '" & "Período: " & (Data_Ini.Text) & " à " & (Data_Fim.Text) & "'"
    .Formulas(2) = "Filial = '" & "Filial: " & (Combo.Text) & "'"
    If optTodos.Value Then .Formulas(3) = "Opcao = '" & "Opção Escolhida: Todos (Recebidos e A Receber)" & "'"
    If optRecebidos.Value Then .Formulas(3) = "Opcao = '" & "Opção Escolhida: Recebidos" & "'"
    If optReceber.Value Then .Formulas(3) = "Opcao = '" & "Opção Escolhida: A Receber" & "'"
    
    'Ordenação
    .SortFields(0) = "+{TotalCartoesGroup.Administradora}"
    
    .WindowState = crptMaximized
    .Destination = IIf(B_Vídeo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", RelSintetico)
  
    .Action = 1
  End With
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
TratarErro:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub
  
End Sub
