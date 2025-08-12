VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmGrafico2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Vendas por Classe"
   ClientHeight    =   3030
   ClientLeft      =   3885
   ClientTop       =   2760
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Grafico2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   6900
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   6675
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
         Left            =   2880
         Picture         =   "Grafico2.frx":4E95A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   322
         Width           =   465
      End
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
         Left            =   5850
         Picture         =   "Grafico2.frx":4F23C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   322
         Width           =   465
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   4485
         TabIndex        =   2
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   375
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         Left            =   1530
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   375
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         BackColor       =   &H80000000&
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   510
         TabIndex        =   8
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3495
         TabIndex        =   7
         Top             =   405
         Width           =   885
      End
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
      Left            =   675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Parâmetro"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton B_Imprime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2385
      Width           =   6645
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   -45
      Top             =   2790
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Filial 
      Bindings        =   "Grafico2.frx":4FB1E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   330
      Width           =   1305
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
      Columns(0).Width=   9419
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
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   12648447
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
      Caption         =   "Filial :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Nome_Filial 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      Top             =   330
      Width           =   4695
   End
End
Attribute VB_Name = "frmGrafico2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEstoque As Recordset
Dim rsParametros As Recordset
Dim rsClasses As Recordset
Dim rsTempo As Recordset
Dim Resumo(10000) As Double

Private Function Acha_Maior() As Long
  Dim i As Long
  Dim Maior As Double
  Dim Índice As Long
   
  Maior = 0
  For i = 1 To 9999
   If Resumo(i) > Maior Then
      Maior = Resumo(i)
      Índice = i
   End If
  Next i
  
  If Maior = 0 Then
     Acha_Maior = 0
     Exit Function
  End If
  
  Acha_Maior = Índice
  
End Function

Private Sub Grava_Tempo(ByVal Classe As Long, ByVal Valor As Double)

  rsClasses.Index = "Código"
  If Valor = 0 Then Exit Sub
  
  rsClasses.Seek "=", Classe
  If rsClasses.NoMatch Then Exit Sub
  
  rsTempo.AddNew
    rsTempo("Nome") = Left(rsClasses("Nome"), 20)
    rsTempo("Valor Vendas") = Format(Valor, "##############0.00")
  rsTempo.Update

End Sub

Private Sub B_Imprime_Click()
On Error GoTo Erro:

 Dim Data_Str As String
 Dim Mês As Integer
 Dim Mês1 As Integer
 Dim i As Long
 Dim Aux_Tamanho As Integer
 Dim Aux_Cor As Integer
 Dim Aux_Edição As Long
 Dim Aux_Produto As String
 Dim Aux_Contador As Double
 Dim Classe As Long
 Dim Erro As Integer
 Dim Restante As Double
 Dim sSql As String
 Dim Str_Rel As String
 Dim Aux_Data As Date
 Dim Aux_Data2 As Date
 
 Dim Maior1 As Long
 Dim Maior2 As Long
 Dim Maior3 As Long
 Dim Maior4 As Long
 Dim Maior5 As Long
 Dim Maior6 As Long
 Dim Maior7 As Long
 Dim Maior8 As Long
 Dim Maior9 As Long
 Dim Maior10 As Long
 Dim Maior11 As Long
 Dim Maior12 As Long
 Dim Maior13 As Long
 Dim Maior14 As Long
 Dim Valor1 As Double
 Dim Valor2 As Double
 Dim Valor3 As Double
 Dim Valor4 As Double
 Dim Valor5 As Double
 Dim Valor6 As Double
 Dim Valor7 As Double
 Dim Valor8 As Double
 Dim Valor9 As Double
 Dim Valor10 As Double
 Dim Valor11 As Double
 Dim Valor12 As Double
 Dim Valor13 As Double
 Dim Valor14 As Double
 
 Dim dbl_valor_total_desconto_sub_total As Double
 Dim dbl_valor_total_devolucao As Double
 

 If Nome_Filial.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo_Filial.SetFocus
   Exit Sub
 End If

 If Not IsDate(Data_Ini.Text) Then
   DisplayMsg "Digite uma data."
   Data_Ini.SetFocus
   Exit Sub
 End If
 If Not IsDate(Data_Fim.Text) Then
   DisplayMsg "Digite uma data."
   Data_Fim.SetFocus
   Exit Sub
 End If
 If CDate(Data_Fim.Text) < CDate(Data_Ini.Text) Then
   DisplayMsg "Data final deve ser superior à data inicial."
   Data_Fim.SetFocus
   Exit Sub
 End If
 
 'Limpa tabela interna
 Erase Resumo
 
 rsEstoque.Index = "Produto"
 Aux_Tamanho = 0
 Aux_Cor = 0
 Aux_Produto = 0
 Aux_Contador = 0
 Aux_Data = CDate(Data_Ini.Text)
 Aux_Edição = 0
 
Lp1:
 rsEstoque.Seek ">", Val(Combo_Filial.Text), Aux_Data, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição
 If rsEstoque.NoMatch Then GoTo Fim
 If CDate(rsEstoque("Data")) > CDate(Data_Fim.Text) Then GoTo Fim
 
 Aux_Tamanho = rsEstoque("Tamanho")
 Aux_Cor = rsEstoque("Cor")
 Aux_Produto = rsEstoque("Produto")
 Aux_Edição = rsEstoque("Edição")
 Aux_Data = rsEstoque("Data")
 
 If rsEstoque("Filial") <> Val(Combo_Filial.Text) Then GoTo Fim
 
  If Aux_Data <> Aux_Data2 Then
    Call StatusMsg("Aguarde, verificando vendas .." & str(Aux_Data))
    Aux_Data2 = Aux_Data
  End If
  
 If rsEstoque("Classe") = 0 Then GoTo Lp1
 
 Classe = rsEstoque("Classe")
 Resumo(Classe) = Resumo(Classe) + rsEstoque("Valor Vendas")
 GoTo Lp1
 
Fim:
 'classifica os mais vendidos
 Maior1 = 0
 Maior2 = 0
 Maior3 = 0
 Maior4 = 0
 Maior5 = 0
 Maior6 = 0
 Maior7 = 0
 Maior8 = 0
 Maior9 = 0
 Maior10 = 0
 Maior11 = 0
 Maior12 = 0
 Maior13 = 0
 Maior14 = 0
 
 Valor1 = 0
 Valor2 = 0
 Valor3 = 0
 Valor4 = 0
 Valor5 = 0
 Valor6 = 0
 Valor7 = 0
 Valor8 = 0
 Valor9 = 0
 Valor10 = 0
 Valor11 = 0
 Valor12 = 0
 Valor13 = 0
 Valor14 = 0
 
  Maior1 = Acha_Maior
  If Maior1 = 0 Then GoTo Imprime
  Valor1 = Resumo(Maior1)
  Resumo(Maior1) = 0
 
  Maior2 = Acha_Maior
  If Maior2 = 0 Then GoTo Imprime
  Valor2 = Resumo(Maior2)
  Resumo(Maior2) = 0
 
 
  Maior3 = Acha_Maior
  If Maior3 = 0 Then GoTo Imprime
  Valor3 = Resumo(Maior3)
  Resumo(Maior3) = 0

 
  Maior4 = Acha_Maior
  If Maior4 = 0 Then GoTo Imprime
  Valor4 = Resumo(Maior4)
  Resumo(Maior4) = 0


  Maior5 = Acha_Maior
  If Maior5 = 0 Then GoTo Imprime
  Valor5 = Resumo(Maior5)
  Resumo(Maior5) = 0


  Maior6 = Acha_Maior
  If Maior6 = 0 Then GoTo Imprime
  Valor6 = Resumo(Maior6)
  Resumo(Maior6) = 0


  Maior7 = Acha_Maior
  If Maior7 = 0 Then GoTo Imprime
  Valor7 = Resumo(Maior7)
  Resumo(Maior7) = 0


  Maior8 = Acha_Maior
  If Maior8 = 0 Then GoTo Imprime
  Valor8 = Resumo(Maior8)
  Resumo(Maior8) = 0


  Maior9 = Acha_Maior
  If Maior9 = 0 Then GoTo Imprime
  Valor9 = Resumo(Maior9)
  Resumo(Maior9) = 0


  Maior10 = Acha_Maior
  If Maior10 = 0 Then GoTo Imprime
  Valor10 = Resumo(Maior10)
  Resumo(Maior10) = 0


  Maior11 = Acha_Maior
  If Maior11 = 0 Then GoTo Imprime
  Valor11 = Resumo(Maior11)
  Resumo(Maior11) = 0


  Maior12 = Acha_Maior
  If Maior12 = 0 Then GoTo Imprime
  Valor12 = Resumo(Maior12)
  Resumo(Maior12) = 0


  Maior13 = Acha_Maior
  If Maior13 = 0 Then GoTo Imprime
  Valor13 = Resumo(Maior13)
  Resumo(Maior13) = 0


  Maior14 = Acha_Maior
  If Maior14 = 0 Then GoTo Imprime
  Valor14 = Resumo(Maior14)
  Resumo(Maior14) = 0


Imprime:

  For i = 1 To 9999
    Restante = Restante + Resumo(i)
  Next i
  
  sSql = "Delete * From ZZZGráfico2"
  db.Execute sSql
  
  Call Grava_Tempo(Maior1, Valor1)
  Call Grava_Tempo(Maior2, Valor2)
  Call Grava_Tempo(Maior3, Valor3)
  Call Grava_Tempo(Maior4, Valor4)
  Call Grava_Tempo(Maior5, Valor5)
  Call Grava_Tempo(Maior6, Valor6)
  Call Grava_Tempo(Maior7, Valor7)
  Call Grava_Tempo(Maior8, Valor8)
  Call Grava_Tempo(Maior9, Valor9)
  Call Grava_Tempo(Maior10, Valor10)
  Call Grava_Tempo(Maior11, Valor11)
  Call Grava_Tempo(Maior12, Valor12)
  Call Grava_Tempo(Maior13, Valor13)
  Call Grava_Tempo(Maior14, Valor14)

  If Restante > 0 Then
    rsTempo.AddNew
      rsTempo("Nome") = "OUTRAS ..."
      rsTempo("Valor Vendas") = Restante
    rsTempo.Update
  End If
  
  Rem Nome do arquivo .rpt
  Str_Rel = gsReportPath & "GRAFI2.RPT"
  Rel.ReportFileName = Str_Rel
 
  Rel.DataFiles(0) = gsQuickDBFileName
 
  Rem Saída
  Rel.Destination = 0
 
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"

  Rel.Formulas(0) = Str_Rel

  Str_Rel = "nome_filial = '"
  Str_Rel = Str_Rel + Nome_Filial.Caption + "'"

  Rel.Formulas(1) = Str_Rel
  
  '06/07/2006 - Andrea
  'Passagem dos parâmetros data inicial e final
  
  Str_Rel = "data_ini = '"
  Str_Rel = Str_Rel + Data_Ini.Text + "'"

  Rel.Formulas(2) = Str_Rel
  
  Str_Rel = "data_fim = '"
  Str_Rel = Str_Rel + Data_Fim.Text + "'"

  Rel.Formulas(3) = Str_Rel
  
  ' Fim da alteracao - Andrea
  
 
  '17/04/2009 - mpdea
  'Incluído o valor total de desconto no sub total e devolução no cálculo final
  dbl_valor_total_desconto_sub_total = GetValorTotalDescontoSubTotal
  dbl_valor_total_devolucao = GetValorTotalDevolucao
  Rel.Formulas(4) = "total_desconto_sub_total = " & Replace(Format(CStr(dbl_valor_total_desconto_sub_total), "###0.00"), gsCurrencyDecimal, ".")
  Rel.Formulas(5) = "total_devolucao = " & Replace(Format(CStr(dbl_valor_total_devolucao), "###0.00"), gsCurrencyDecimal, ".")
  
  
  Rel.WindowState = crptMaximized
  Call StatusMsg("Aguarde, imprimindo...")
  
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
  
  Rel.Action = 1

  Call StatusMsg("")
  
  Exit Sub
Erro:
    MsgBox "Erro na função de geração do relatório " & Err.Number & " " & Err.Description, vbInformation, "Atenção"

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
  If Val(Combo_Filial.Text) < 0 Then Exit Sub
  If Val(Combo_Filial.Text) > 99 Then Exit Sub

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Val(Combo_Filial.Text)
  If rsParametros.NoMatch Then Exit Sub
  Nome_Filial.Caption = rsParametros("Nome")

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsEstoque = db.OpenRecordset("Estoque", , dbReadOnly)
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsTempo = db.OpenRecordset("ZZZGráfico2")
  
  Data1.DatabaseName = gsQuickDBFileName

End Sub

'17/04/2009 - mpdea
'Obtém o valor total de desconto no sub total de vendas no período
Private Function GetValorTotalDescontoSubTotal() As Double
  Dim str_sql As String
  Dim rst_total As Recordset
  Dim dbl_total As Double
  
  str_sql = "SELECT SUM(s.DescontoSubTotal) AS ValorTotal "
  str_sql = str_sql & "FROM Saídas s INNER JOIN [Operações Saída] op ON s.Operação = op.Código "
  str_sql = str_sql & "WHERE s.Efetivada AND NOT s.[Nota Cancelada] AND op.Tipo = 'V' "
  str_sql = str_sql & "AND s.DescontoSubTotal > 0 "
  str_sql = str_sql & "AND s.Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
  str_sql = str_sql & "AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
  str_sql = str_sql & "AND s.Filial = " & Combo_Filial.Text
  
  Set rst_total = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  With rst_total
    Call IsDataType(dtDouble, .Fields("ValorTotal").Value, dbl_total)
    .Close
  End With
  Set rst_total = Nothing
  
  GetValorTotalDescontoSubTotal = dbl_total
End Function

'17/04/2009 - mpdea
'Obtém o valor total de devoluções no período
Private Function GetValorTotalDevolucao() As Double
  Dim str_sql As String
  Dim rst_total As Recordset
  Dim dbl_total As Double
  
  str_sql = "SELECT SUM(ep.[Preço Final]) AS ValorTotal "
  str_sql = str_sql & "FROM (Entradas e INNER JOIN [Entradas - Produtos] ep "
  str_sql = str_sql & "ON (e.Sequência = ep.Sequência) AND (e.Filial = ep.Filial)) "
  str_sql = str_sql & "INNER JOIN [Operações Entrada] op ON e.Operação = op.Código "
  str_sql = str_sql & "WHERE e.Efetivada AND NOT e.[Nota Cancelada] AND op.Tipo = 'D' "
  str_sql = str_sql & "AND e.Data BETWEEN #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "# "
  str_sql = str_sql & "AND #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "# "
  str_sql = str_sql & "AND e.Filial = " & Combo_Filial.Text
  
  Set rst_total = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  With rst_total
    Call IsDataType(dtDouble, .Fields("ValorTotal").Value, dbl_total)
    .Close
  End With
  Set rst_total = Nothing
  
  GetValorTotalDevolucao = dbl_total
End Function

