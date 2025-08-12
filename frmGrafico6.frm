VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGrafico6 
   Caption         =   " Maiores Clientes no período"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrafico6.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   13920
      Begin VB.ComboBox cmb_uf 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmGrafico6.frx":4E95A
         Left            =   750
         List            =   "frmGrafico6.frx":4E9B2
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   690
         Width           =   4935
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
         Left            =   7350
         Picture         =   "frmGrafico6.frx":4EB78
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   202
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
         Left            =   9510
         Picture         =   "frmGrafico6.frx":4F45A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   202
         Width           =   465
      End
      Begin VB.ComboBox cmb_numClientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmGrafico6.frx":4FD3C
         Left            =   11460
         List            =   "frmGrafico6.frx":4FD64
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1290
      End
      Begin VB.ComboBox Combo_Filial 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmGrafico6.frx":4FDA1
         Left            =   750
         List            =   "frmGrafico6.frx":4FDA3
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   870
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   8340
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   255
         Width           =   1125
         _ExtentX        =   1984
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
         Left            =   6195
         TabIndex        =   5
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   255
         Width           =   1125
         _ExtentX        =   1984
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
         Caption         =   "Estado"
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
         Left            =   105
         TabIndex        =   14
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Até"
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
         Left            =   7935
         TabIndex        =   10
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "De"
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
         Left            =   5865
         TabIndex        =   9
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label7 
         Caption         =   "Filial"
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
         Left            =   105
         TabIndex        =   8
         Top             =   285
         Width           =   390
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
         Left            =   1665
         TabIndex        =   7
         Top             =   255
         Width           =   4005
      End
      Begin VB.Label Label2 
         Caption         =   "Visualizar (os)                            maiores clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10275
         TabIndex        =   6
         Top             =   285
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmd_pesquisar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   13920
   End
   Begin MSFlexGridLib.MSFlexGrid gridClientes 
      Height          =   6255
      Left            =   90
      TabIndex        =   11
      Top             =   1800
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
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
End
Attribute VB_Name = "frmGrafico6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrayFiliais(20, 2) As String
Dim iConta As Integer

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

Private Sub cmd_pesquisar_Click()
On Error GoTo Erro
 
  Dim rsSaidas As Recordset
  Dim strSQL As String
  Dim lngContadorRegGrid As Long
  Dim iIndice As Integer
 
  If Nome_Filial.Caption = "" Then
    DisplayMsg "Escolha a filial."
    Combo_Filial.SetFocus
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
   
  gridClientes.Rows = 1
  gridClientes.Row = 0
  
  If cmb_numClientes.Text = "TODOS" Then
    strSQL = "SELECT Sum(S.Total) AS SomaDeVendas, "
  Else
    strSQL = "SELECT top " & cmb_numClientes.Text & " Sum(S.Total) AS SomaDeVendas, "
  End If

  strSQL = strSQL & " S.Cliente, C.Nome, C.Estado "
  strSQL = strSQL & " From Saídas S, Cli_for C, [Operações Saída] Op "
  strSQL = strSQL & " where S.data >= CDATE('" & Data_Ini.Text & " 00:00:00') and "
  strSQL = strSQL & " S.data <= CDATE('" & Data_Fim.Text & " 00:00:00') and "
  strSQL = strSQL & " S.Filial=" & Combo_Filial.Text & " and "
  strSQL = strSQL & " S.Efetivada=1 and "
  strSQL = strSQL & " S.[Movimentação Desfeita]=0 and "
  strSQL = strSQL & " S.Cliente=C.Código and "
  
  If cmb_uf.Text <> "" And cmb_uf.Text <> "TODOS" Then
      iIndice = InStr(1, cmb_uf.Text, "(")
      strSQL = strSQL & " C.Estado = '" & Mid(cmb_uf.Text, iIndice + 1, 2) & "' and "
  End If
  
  strSQL = strSQL & " S.Operação=Op.Código and "
  strSQL = strSQL & " Op.Tipo='V' "
  strSQL = strSQL & " GROUP BY S.Cliente, C.Nome, C.Estado "
  strSQL = strSQL & " ORDER BY 1 DESC "

  Screen.MousePointer = vbHourglass
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  lngContadorRegGrid = 1
  
  If Not (rsSaidas.EOF And rsSaidas.BOF) Then
    rsSaidas.MoveFirst
  End If
  While Not rsSaidas.EOF
  
      gridClientes.AddItem lngContadorRegGrid & vbTab & FormataValorTexto(rsSaidas.Fields(0).Value, 2) & vbTab & _
                      rsSaidas.Fields(1).Value & vbTab & _
                      LTrim(RTrim(rsSaidas.Fields(2).Value) & vbTab & _
                      rsSaidas.Fields(3).Value)
                      
      rsSaidas.MoveNext
      lngContadorRegGrid = lngContadorRegGrid + 1
  Wend
  rsSaidas.Close
  Set rsSaidas = Nothing
  
  Screen.MousePointer = vbDefault
  Exit Sub
Erro:
  If Not (rsSaidas Is Nothing) Then
      rsSaidas.Close
      Set rsSaidas = Nothing
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

Private Sub Combo_Filial_LostFocus()
  Dim i As Integer
  Nome_Filial.Caption = ""
  
  If Combo_Filial.Text <> "" Then
      For i = 0 To iConta
        If Combo_Filial.Text = arrayFiliais(i, 0) Then
          Nome_Filial.Caption = arrayFiliais(i, 1)
          Exit For
        End If
      Next
  End If
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

Private Sub Form_Load()
On Error GoTo Erro:

  Dim rsParametros As Recordset
  Set rsParametros = db.OpenRecordset("select Filial, Nome from [Parâmetros Filial]", dbOpenDynaset)
  
  iConta = 0
  While Not rsParametros.EOF
      arrayFiliais(iConta, 0) = rsParametros.Fields(0).Value
      arrayFiliais(iConta, 1) = rsParametros.Fields(1).Value
      Combo_Filial.AddItem rsParametros.Fields(0).Value, iConta
      iConta = iConta + 1
      rsParametros.MoveNext
  Wend
  rsParametros.Close
  Set rsParametros = Nothing

  
  gridClientes.ColWidth(0) = 600
  gridClientes.ColWidth(1) = 2000
  gridClientes.ColWidth(2) = 2500
  gridClientes.ColWidth(3) = 8100
  gridClientes.ColWidth(4) = 400
  'gridClientes.ColAlignment(3) = Left
  
  gridClientes.Row = 0
  gridClientes.TextMatrix(0, 1) = "Qtde R$ faturado"
  gridClientes.TextMatrix(0, 2) = "Código do cliente"
  gridClientes.TextMatrix(0, 3) = "Nome do cliente"
  gridClientes.TextMatrix(0, 4) = "UF"
  
  Data_Fim.Text = Format(Now, "dd/mm/yyyy")
  Data_Ini.Text = Format(Now - 90, "dd/mm/yyyy")
  
  cmb_numClientes.ListIndex = 5
  
  Combo_Filial.ListIndex = 0
  Combo_Filial_LostFocus
  
  Exit Sub
Erro:
  MsgBox "Erro na carga da tela. Cod: " & Err.Number & " - Desc: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub gridClientes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridClientes.Redraw = False
End Sub

Private Sub gridClientes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  gridClientes.RowSel = gridClientes.Row
  gridClientes.Redraw = True
End Sub


