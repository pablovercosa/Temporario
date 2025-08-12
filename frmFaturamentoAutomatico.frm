VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmFaturamentoAutomatico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento Automático"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamentoAutomatico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11760
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCarregar 
      Caption         =   "Ca&rregar"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame fraCancelamento 
      Caption         =   "Cancelamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   7440
      TabIndex        =   23
      Top             =   100
      Width           =   4270
      Begin VB.CommandButton cmdDesistir 
         Caption         =   "&Desistir"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1150
         Width           =   915
      End
      Begin VB.CheckBox chk4 
         Caption         =   "4"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chk3 
         Caption         =   "3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chk2 
         Caption         =   "2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chk1 
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdConfirmarCancelamento 
         Caption         =   "Confirm&ar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   1150
         Width           =   915
      End
      Begin VB.CommandButton cmdIniciarCancelamento 
         Caption         =   "&Iniciar"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Cancelar uma Linha de cada vez..."
         Top             =   280
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Parcela(s):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   24
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Frame fraImpressao 
      Caption         =   "Nota Fiscal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   755
      Left            =   0
      TabIndex        =   22
      Top             =   860
      Width           =   3700
      Begin VB.OptionButton optNaoImpresso 
         Appearance      =   0  'Flat
         Caption         =   "Não Impresso"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optImpresso 
         Appearance      =   0  'Flat
         Caption         =   "Impresso"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraAutorizacoes 
      Caption         =   "Intervalo de Nº de Autorizações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   755
      Left            =   3720
      TabIndex        =   19
      Top             =   100
      Width           =   3700
      Begin VB.TextBox txtNumIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         MaxLength       =   8
         TabIndex        =   2
         Top             =   280
         Width           =   1215
      End
      Begin VB.TextBox txtNumFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   3
         Top             =   280
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   340
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   340
         Width           =   90
      End
   End
   Begin VB.Frame fraData 
      Caption         =   "Intervalo de Data Assinatura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   755
      Left            =   0
      TabIndex        =   16
      Top             =   100
      Width           =   3700
      Begin MSMask.MaskEdBox mskDataAssinaturaIni 
         Height          =   315
         Left            =   600
         TabIndex        =   0
         ToolTipText     =   "Pressione F2 para obter calendário."
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataAssinaturaFin 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Pressione F2 para obter calendário."
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   1920
         TabIndex        =   18
         Top             =   340
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   340
         Width           =   195
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultado 
      Height          =   4140
      Left            =   45
      TabIndex        =   25
      Top             =   1680
      Width           =   11655
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   14
      BevelColorFrame =   -2147483632
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      AllowRowSizing  =   0   'False
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   14
      Columns(0).Width=   1746
      Columns(0).Caption=   "Data"
      Columns(0).Name =   "DataAssinatura"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   7
      Columns(0).FieldLen=   256
      Columns(1).Width=   1799
      Columns(1).Caption=   "Autorização"
      Columns(1).Name =   "NumeroAutorizacao"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   741
      Columns(2).Caption=   "Mês"
      Columns(2).Name =   "MesX"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   2
      Columns(2).FieldLen=   256
      Columns(3).Width=   1984
      Columns(3).Caption=   "Programação"
      Columns(3).Name =   "Programacao"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1826
      Columns(4).Caption=   "Valor Program."
      Columns(4).Name =   "VlProgramacao"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(5).Width=   3228
      Columns(5).Caption=   "Cliente"
      Columns(5).Name =   "Cliente"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1693
      Columns(6).Caption=   "Parc 1"
      Columns(6).Name =   "Parc1"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1693
      Columns(7).Caption=   "Venc 1"
      Columns(7).Name =   "Venc1"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1693
      Columns(8).Caption=   "Parc 2"
      Columns(8).Name =   "Parc2"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1693
      Columns(9).Caption=   "Venc 2"
      Columns(9).Name =   "Venc2"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   1693
      Columns(10).Caption=   "Parc 3"
      Columns(10).Name=   "Parc3"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1693
      Columns(11).Caption=   "Venc 3"
      Columns(11).Name=   "Venc3"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   1693
      Columns(12).Caption=   "Parc 4"
      Columns(12).Name=   "Parc4"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1693
      Columns(13).Caption=   "Venc 4"
      Columns(13).Name=   "Venc4"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      _ExtentX        =   20558
      _ExtentY        =   7302
      _StockProps     =   79
      Caption         =   "Autorizações existentes no intervalo solicitado"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFaturamentoAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngNumAutorizacao    As Long
Private m_intMesX              As Integer

Private Sub cmdCarregar_Click()
  Dim rstContrato    As Recordset
  Dim strSQL         As String
  
  'Validação das Datas...
  If Not IsDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Entrar com Data Inicial válida.", vbExclamation, "Quick Store"
    mskDataAssinaturaIni.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(mskDataAssinaturaFin.Text) Then
    MsgBox "Entrar com Data Final válida.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If
  
  If CDate(mskDataAssinaturaFin.Text) < CDate(mskDataAssinaturaIni.Text) Then
    MsgBox "Data Final menor que a Data Inicial.", vbExclamation, "Quick Store"
    mskDataAssinaturaFin.SetFocus
    Exit Sub
  End If
  'Fim da Validação das Datas
  
  
  'Preparando a grid...
  With grdResultado
    'Não permite atualizar o layout da grid
    .Redraw = False
    'Limpa o grid
    .RemoveAll
    'Permite atualizar o layout da grid
    .Redraw = True
  End With
  'Fim da preparação da grid
  
  'Início do SELECT...
  strSQL = " SELECT Contrato.[Num Autorizacao] AS AutorizacaoPai, Programacao.[Num Autorizacao] AS AutorizacaoFilho, * "
  strSQL = strSQL & " FROM Contrato, Programacao "
  strSQL = strSQL & " WHERE Contrato.[Data Assinatura] >= #" & CDate(mskDataAssinaturaIni.Text) & "#"
  strSQL = strSQL & " AND Contrato.[Data Assinatura] <= #" & CDate(mskDataAssinaturaFin.Text) & "#"
  strSQL = strSQL & " AND Programacao.[Num Autorizacao] = Contrato.[Num Autorizacao] "
  
  If IsNumeric(txtNumIni.Text) And IsNumeric(txtNumFin.Text) Then
    strSQL = strSQL & " AND Contrato.[Num Autorizacao] >= " & CLng(txtNumIni.Text)
    strSQL = strSQL & " AND Contrato.[Num Autorizacao] <= " & CLng(txtNumFin.Text)
  End If
  
  'Já foi impresso a NF
  If optNaoImpresso.Value Then
    strSQL = strSQL & " AND Programacao.ImpressoNF = FALSE "
  Else
    strSQL = strSQL & " AND Programacao.ImpressoNF = TRUE "
  End If
  
  strSQL = strSQL & " ORDER BY Contrato.[Data Assinatura], Programacao.[Num Autorizacao], Programacao.MesX "
  'Fim do SELECT
  
  Set rstContrato = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstContrato
    If Not (.BOF And .EOF) Then
      .MoveFirst
    
      Do Until .EOF
        grdResultado.AddNew
        
        grdResultado.Columns("DataAssinatura").Text = .Fields("Data Assinatura").Value
        grdResultado.Columns("NumeroAutorizacao").Text = .Fields("AutorizacaoFilho").Value
        grdResultado.Columns("MesX").Text = .Fields("MesX").Value
        
        m_intMesX = .Fields("MesX").Value
        m_lngNumAutorizacao = .Fields("AutorizacaoFilho").Value
        
        grdResultado.Columns("Programacao").Text = .Fields("Programacao").Value
        grdResultado.Columns("VlProgramacao").Text = .Fields("Valor Total").Value
        grdResultado.Columns("Cliente").Text = .Fields("Cod Cliente").Value & " - " & GetNomeCliFor(.Fields("Cod Cliente").Value)
        
        If .Fields("Status1").Value Then
          grdResultado.Columns("Parc1").Text = .Fields("Valor1").Value
          grdResultado.Columns("Venc1").Text = .Fields("Vencimento1").Value
        End If
        
        If .Fields("Status2").Value Then
          grdResultado.Columns("Parc2").Text = .Fields("Valor2").Value
          grdResultado.Columns("Venc2").Text = .Fields("Vencimento2").Value
        End If
        
        If .Fields("Status3").Value Then
          grdResultado.Columns("Parc3").Text = .Fields("Valor3").Value
          grdResultado.Columns("Venc3").Text = .Fields("Vencimento3").Value
        End If
        
        If .Fields("Status4").Value Then
          grdResultado.Columns("Parc4").Text = .Fields("Valor4").Value
          grdResultado.Columns("Venc4").Text = .Fields("Vencimento4").Value
        End If
        
        grdResultado.Update
        .MoveNext
      Loop
    
    Else
      MsgBox "Nenhuma informação foi encontrada dentro dos critérios escolhidos.", vbInformation, "Quick Store"
      mskDataAssinaturaIni.SetFocus
    End If
    .Close
  End With
  
  Set rstContrato = Nothing
  
End Sub

Private Sub cmdConfirmarCancelamento_Click()
  Dim intAuxi        As Integer
  Dim varBook        As Variant
  Dim lngNumero      As Long
  Dim intMesX        As Integer
  Dim strQuery       As String
  Dim rstProgramacao As Recordset
  Dim dblValor       As Double
  
  If VerificaQuantasLinhas Then Exit Sub
  
  For intAuxi = 0 To (grdResultado.SelBookmarks.Count - 1)
    varBook = grdResultado.SelBookmarks(intAuxi)
    grdResultado.Bookmark = varBook

    lngNumero = grdResultado.Columns("NumeroAutorizacao").CellValue(varBook)
    intMesX = grdResultado.Columns("MesX").CellValue(varBook)

    strQuery = "SELECT [Num Autorizacao], MesX, Cancel1, Cancel2, Cancel3, Cancel4, SomaCancelamento "
    strQuery = strQuery & " FROM Programacao "
    strQuery = strQuery & " WHERE [Num Autorizacao] = " & lngNumero
    strQuery = strQuery & " AND MesX = " & intMesX
    
    Set rstProgramacao = db.OpenRecordset(strQuery, dbOpenDynaset)
  
    With rstProgramacao
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        .Edit
        
        If chk1.Value Then
          .Fields("Cancel1").Value = True
          If Len(grdResultado.Columns("Parc1").CellValue(varBook)) > 0 Then dblValor = dblValor + CDbl(grdResultado.Columns("Parc1").CellValue(varBook))
        End If
        
        If chk2.Value Then
          .Fields("Cancel2").Value = True
          If Len(grdResultado.Columns("Parc2").CellValue(varBook)) > 0 Then dblValor = dblValor + CDbl(grdResultado.Columns("Parc2").CellValue(varBook))
        End If
        
        If chk3.Value Then
          .Fields("Cancel3").Value = True
          If Len(grdResultado.Columns("Parc3").CellValue(varBook)) > 0 Then dblValor = dblValor + CDbl(grdResultado.Columns("Parc3").CellValue(varBook))
        End If
        
        If chk4.Value Then
          .Fields("Cancel4").Value = True
          If Len(grdResultado.Columns("Parc4").CellValue(varBook)) > 0 Then dblValor = dblValor + CDbl(grdResultado.Columns("Parc4").CellValue(varBook))
        End If
        
        .Fields("SomaCancelamento").Value = Format(dblValor, "##,###,###,##0.00")
        
        .Update
      End If
      .Close
    End With
    
    Set rstProgramacao = Nothing
    
  Next intAuxi

  MsgBox "Parcela(s) cancelada(s) com sucesso.", vbExclamation, "Quick Store"

  chk1.Value = vbUnchecked
  chk2.Value = vbUnchecked
  chk3.Value = vbUnchecked
  chk4.Value = vbUnchecked

End Sub

Private Sub cmdDesistir_Click()
  'Desabilitar os objetos para o cancelamento
  chk1.Enabled = False
  chk2.Enabled = False
  chk3.Enabled = False
  chk4.Enabled = False
  cmdConfirmarCancelamento.Enabled = False
  cmdDesistir.Enabled = False
  
  chk1.Value = vbUnchecked
  chk2.Value = vbUnchecked
  chk3.Value = vbUnchecked
  chk4.Value = vbUnchecked
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim intAuxi   As Integer
  Dim varBook   As Variant
  Dim lngNumero As Long
  Dim intMesX   As Integer

  If ExaminaSelecao Then Exit Sub

  Screen.MousePointer = vbHourglass

  For intAuxi = 0 To (grdResultado.SelBookmarks.Count - 1)
    varBook = grdResultado.SelBookmarks(intAuxi)
    grdResultado.Bookmark = varBook

    lngNumero = grdResultado.Columns("NumeroAutorizacao").CellValue(varBook)
    intMesX = grdResultado.Columns("MesX").CellValue(varBook)

    Call PrintNota(lngNumero, intMesX)
    Call AtualizarCampoImpressoNF(lngNumero, intMesX)
  Next intAuxi
  
  Screen.MousePointer = vbDefault

  If MsgBox("Deseja imprimir os Ticktes? ", vbQuestion + vbYesNo, "Troca de Formulário") = vbYes Then
    Screen.MousePointer = vbHourglass
    Call PrintTicket
    Screen.MousePointer = vbDefault
  End If

End Sub

Private Sub cmdIniciarCancelamento_Click()
  'Habilitar os objetos para o cancelamento
  chk1.Enabled = True
  chk2.Enabled = True
  chk3.Enabled = True
  chk4.Enabled = True
  cmdConfirmarCancelamento.Enabled = True
  cmdDesistir.Enabled = True
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
Private Sub mskDataAssinaturaFin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataAssinaturaFin.Text = frmCalendario.gsDateCalender(mskDataAssinaturaFin.Text)
  End If
End Sub

Private Sub mskDataAssinaturaFin_LostFocus()
  mskDataAssinaturaFin.Text = Ajusta_Data(mskDataAssinaturaFin.Text)
End Sub

Private Sub mskDataAssinaturaIni_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataAssinaturaIni.Text = frmCalendario.gsDateCalender(mskDataAssinaturaIni.Text)
  End If
End Sub

Private Sub mskDataAssinaturaIni_LostFocus()
  mskDataAssinaturaIni.Text = Ajusta_Data(mskDataAssinaturaIni.Text)
End Sub

Private Function GetNomeCliFor(lngCodigo As Long) As String
  Dim rstCliFor As Recordset
  
  Set rstCliFor = db.OpenRecordset("SELECT Nome FROM Cli_For WHERE Código = " & lngCodigo, dbOpenDynaset)
  
  With rstCliFor
    GetNomeCliFor = IIf((.BOF And .EOF), "<_não_cadastrado>", .Fields("Nome").Value & "")
    .Close
  End With
  
  Set rstCliFor = Nothing
End Function

Private Function ExaminaSelecao() As Boolean
  If grdResultado.SelBookmarks.Count < 1 Then
    MsgBox "Favor selecionar alguma programação da grid.", vbExclamation, "Quick Store"
    ExaminaSelecao = True
  End If
End Function

Private Function VerificaQuantasLinhas() As Boolean
  If grdResultado.SelBookmarks.Count < 1 Then
    MsgBox "Favor selecionar uma programação da grid.", vbExclamation, "Quick Store"
    VerificaQuantasLinhas = True
  End If
  
  If grdResultado.SelBookmarks.Count > 1 Then
    MsgBox "Favor selecionar apenas uma programação da grid.", vbExclamation, "Quick Store"
    VerificaQuantasLinhas = True
  End If
End Function

Private Sub PrintNota(ByVal Numero As Long, ByVal MesX As Integer)
  'Copiado a Private do mesmo modo que existe em Saídas
  Dim strSQL                As String
  Dim intX                  As Integer
  Dim strFileNF             As String
  Dim intRet                As Integer
  Dim lngNotaFiscal         As Long
  Dim blnInTransaction      As Boolean
  Dim intRepeatUpdateLocked As Integer
  
  Dim rstSaidas             As Recordset
  Dim rstParametros         As Recordset
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
  
  'Abrir a tabela Parâmetros
  strSQL = "SELECT * FROM [Parâmetros Filial]"
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  
  Set rstParametros = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  strSQL = ""
  
  'Abrir a tabela Saídas
  strSQL = "SELECT * FROM Saídas "
  strSQL = strSQL & " WHERE [Num Autorizacao] = " & Numero
  strSQL = strSQL & " AND MesX = " & MesX
  
  Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstSaidas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If rstSaidas.Fields("Nota Impressa").Value <> 0 Then
        lngNotaFiscal = rstSaidas.Fields("Nota Impressa").Value
      
        If MsgBox("A NF " & (rstSaidas.Fields("Nota Impressa").Value) & " já foi impressa, deseja imprimir novamente?", _
          vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then
          Exit Sub
        End If
      Else
        lngNotaFiscal = 0
      End If
    End If
  End With
  
'  Call IsDataType(dtLong, rstSaidas.Fields("Nota Impressa").Value, lngNotaFiscal)
'  If lngNotaFiscal <> 0 Then
'    If MsgBox("A NF " & (rstSaidas.Fields("Nota Impressa").Value) & " já foi impressa, deseja imprimir novamente?", _
'      vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then
'      Exit Sub
'    End If
'  End If
  
  
  '--------------------------------------------------------------------------
  'Grava nova NF
  '--------------------------------------------------------------------------
  If lngNotaFiscal = 0 Then
    'Modificado leitura e gravação do número da última nota fiscal
    'Incluído transação durante gravação
    'lngNotaFiscal = rsParametros.Fields("Última Nota").Value + 1
    '
    ws.BeginTrans
    blnInTransaction = True
    
    lngNotaFiscal = g_lngNextNotaFiscal(rstSaidas.Fields("Filial").Value)

    With rstSaidas
      .LockEdits = True
      .Edit
      .Fields("Nota Impressa").Value = lngNotaFiscal
      
      .Update
      .LockEdits = False
    End With
    
    
    '05/05/2005 - mpdea
    'Atualiza a Nota Fiscal e Fatura do Contas a Receber
    Call StatusMsg("Verificando e atualizando contas a receber...")
    
    strSQL = "UPDATE [Contas a Receber] SET Nota = " & lngNotaFiscal
    strSQL = strSQL & ", Fatura = '" & lngNotaFiscal & "/ ' & Parcela"
    strSQL = strSQL & " WHERE Tipo = 'R'"
    strSQL = strSQL & " AND Filial = " & rstSaidas.Fields("Filial").Value
    strSQL = strSQL & " AND Sequência = " & rstSaidas.Fields("Sequência").Value
    
    db.Execute strSQL, dbFailOnError
    
    '10/09/2007 - Anderson
    'Gera arquivo log do sistema
    If g_bolSystemLog Then
      SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Alterar, strSQL, "frmFaturamentoAutomatico_PrintNota", "Contas a Receber", g_strArquivoSystemLog
    End If

    Call StatusMsg("")
    
    'Finaliza transação
    ws.CommitTrans
    blnInTransaction = False
  End If
  '--------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------
  'Imprime NF
  '--------------------------------------------------------------------------
  strFileNF = gsConfigPath + rstParametros.Fields("Nota Saída").Value + ".CNF"
  intRet = Imprime_Nota(strFileNF, rstSaidas.Fields("Filial").Value, rstSaidas.Fields("Sequência").Value)
  If intRet = 0 Then
    '14/04/2003 - mpdea
    'Atualiza a data da impressão da nota fiscal
    strSQL = "UPDATE Saídas SET DataEmissaoNota = #"
    strSQL = strSQL & Format(Date, "mm/dd/yyyy") & "# "
    strSQL = strSQL & "WHERE Filial = " & rstSaidas.Fields("Filial").Value
    strSQL = strSQL & " AND Sequência = " & rstSaidas.Fields("Sequência").Value
    db.Execute strSQL, dbFailOnError
    
    'DisplayMsg "Nota [" & lngNotaFiscal & "] impressa com sucesso."
  Else
    DisplayMsg "Houve o erro " & intRet & " durante a impressão da Nota."
  End If
  '--------------------------------------------------------------------------
  
  'Fechar os Recordsets
  rstParametros.Close
  rstSaidas.Close
  Set rstParametros = Nothing
  Set rstSaidas = Nothing
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3197, 3187, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Saídas - Imprimir Nota Fiscal") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          'Cancelamento da transação
'          If blnInTransaction Then ws.Rollback
'          Exit Sub
'        End If
      End If
    Case Else
      'Cancelamento da transação
      If blnInTransaction Then ws.Rollback
      'Outros Erros
      MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  End Select


End Sub

Private Sub AtualizarCampoImpressoNF(ByVal Numero As Long, ByVal MesX As Integer)
  Dim rstContrato As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT [Num Autorizacao], MesX, ImpressoNF "
  strSQL = strSQL & " FROM Programacao "
  strSQL = strSQL & " WHERE [Num Autorizacao] = " & Numero
  strSQL = strSQL & " AND MesX = " & MesX
  
  Set rstContrato = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstContrato
    If Not (.BOF And .EOF) Then
      .Edit
      .Fields("ImpressoNF").Value = True
      .Update
    End If
    .Close
  End With
  
  Set rstContrato = Nothing

End Sub

Private Sub PrintTicket()
  Dim intAuxi       As Integer
  Dim varBook       As Variant
  Dim lngNumero     As Long
  Dim intMesX       As Integer
  Dim strTicket     As String
  Dim blnSemTicket  As Boolean
  Dim strSQL        As String
  Dim rstSaidas     As Recordset
  Dim rstParametros As Recordset

  If ExaminaSelecao Then Exit Sub
  
  '----------------------------------------------------------
  'Verificar se há um Ticket padrão cadastrado
  '----------------------------------------------------------
  Set rstParametros = db.OpenRecordset("SELECT Filial, TicketPadrao FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("TicketPadrao").Value <> "" Then
        strTicket = gsConfigPath & .Fields("TicketPadrao").Value & "" & ".CTI"
      Else
        blnSemTicket = True
      End If
    End If
    .Close
  End With
  
  Set rstParametros = Nothing
  
  If blnSemTicket Then
    MsgBox "Cadastre em Parâmetros um Ticket Padrão.", vbExclamation, "Atenção"
    Exit Sub
  End If
  '----------------------------------------------------------

  For intAuxi = 0 To (grdResultado.SelBookmarks.Count - 1)
    varBook = grdResultado.SelBookmarks(intAuxi)
    grdResultado.Bookmark = varBook

    lngNumero = grdResultado.Columns("NumeroAutorizacao").CellValue(varBook)
    intMesX = grdResultado.Columns("MesX").CellValue(varBook)

    'Abrindo a table Saídas
    strSQL = "SELECT Filial, Sequência, [Num Autorizacao], MesX "
    strSQL = strSQL & " FROM Saídas "
    strSQL = strSQL & " WHERE [Num Autorizacao] = " & lngNumero
    strSQL = strSQL & " AND MesX = " & intMesX
    strSQL = strSQL & " AND Filial = " & gnCodFilial
    
    Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
    With rstSaidas
      If Not (.BOF And .EOF) Then
        .MoveFirst
        'Chamamos a impressão do Ticket
        Call Imprime_Ticket(strTicket, gnCodFilial, rstSaidas("Sequência").Value)
      End If
      .Close
    End With
    
    Set rstSaidas = Nothing
  
  Next intAuxi

End Sub
