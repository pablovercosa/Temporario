VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDesconto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Desconto"
   ClientHeight    =   7965
   ClientLeft      =   2145
   ClientTop       =   2220
   ClientWidth     =   6300
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1850
   Icon            =   "Desconto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7965
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3540
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3540
      Width           =   3015
   End
   Begin VB.PictureBox picDescontoRateado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3345
      ScaleWidth      =   6045
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   6075
      Begin VB.TextBox txtPorcentagem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "0"
         Top             =   1575
         Width           =   675
      End
      Begin VB.TextBox txtDescontoSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1650
         TabIndex        =   2
         Text            =   "0"
         Top             =   1575
         Width           =   1935
      End
      Begin Threed.SSPanel sspInfo 
         Height          =   735
         Index           =   2
         Left            =   30
         TabIndex        =   17
         Top             =   60
         Width           =   5985
         _Version        =   65536
         _ExtentX        =   10557
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Digite o valor total de desconto ou o percentual desejado. O desconto será retirado do preço unitário."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
      End
      Begin VB.Label lblDescontoProgFidelidade2_SALDO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   3690
         TabIndex        =   23
         Top             =   1410
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "percentual %"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   4230
         TabIndex        =   18
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   21
         Top             =   1650
         Width           =   675
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desconto fornecido"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   20
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDescontoFornecido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1650
         TabIndex        =   19
         Top             =   885
         Width           =   1935
      End
      Begin VB.Label Total 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.PictureBox picDescontoSubTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3345
      ScaleWidth      =   6045
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.TextBox txtPorcentagem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5310
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "0"
         Top             =   2085
         Width           =   645
      End
      Begin VB.TextBox txtDescontoSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1470
         TabIndex        =   0
         Text            =   "0"
         Top             =   2085
         Width           =   2055
      End
      Begin Threed.SSPanel sspInfo 
         Height          =   735
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Digite o valor de desconto ou percentual desejado que será concedido como  desconto no sub total do cupom fiscal. "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
      End
      Begin Threed.SSPanel sspInfo 
         Height          =   675
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   5865
         _Version        =   65536
         _ExtentX        =   10345
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "Após conceder o desconto no SubTotal não será possível vender mais itens. A tela de Recebimento será acionada logo em seguida."
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   0
         BevelInner      =   1
      End
      Begin VB.Label lblDescontoProgFidelidade_SALDO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   3660
         TabIndex        =   22
         Top             =   1950
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblTotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   14
         Top             =   1485
         Width           =   2055
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Geral"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label lblNovoTotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   11
         Top             =   2715
         Width           =   2055
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Novo Total Geral"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   10
         Top             =   2790
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "percentual %"
         BeginProperty Font 
            Name            =   "WeblySleek UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   4290
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'30/04/2003 - mpdea
'Reformulado para suportar o desconto rateado (modo anterior)

'20/09/2002 - mpdea
'Totalmente reformulado para suportar o desconto no subtotal

Private mblnPressOK As Boolean
Private msngMaxDescPerc As Single
Private mcurTotalGeral As Currency
Private mcurDesconto As Currency
Private m_blnDescontoRateado As Boolean

Public Function Start(ByVal curTotalGeral As Currency, _
  ByVal sngMaxDescPerc As Single, ByRef curDescConcedido As Currency, _
  ByRef curNewTotal As Currency, ByVal blnDescontoRateado As Boolean, _
  Optional ByVal dblTotalDesconto As Double = 0) As Boolean
  
  mblnPressOK = False
  
  gnDesconto = 0
  
  mcurTotalGeral = curTotalGeral
  msngMaxDescPerc = sngMaxDescPerc
  m_blnDescontoRateado = blnDescontoRateado
  
  If blnDescontoRateado Then
    With picDescontoRateado
      .Visible = True
      .BorderStyle = 0
      .Top = 120
    End With
    lblDescontoFornecido.Caption = Format(dblTotalDesconto, FORMAT_VALUE)
    Call SelectAllText(txtDescontoSubTotal(1))
  Else
    With picDescontoSubTotal
      .Visible = True
      .BorderStyle = 0
      .Top = 120
    End With
    lblTotalGeral.Caption = Format(curTotalGeral, FORMAT_VALUE)
    lblNovoTotalGeral.Caption = Format(curTotalGeral, FORMAT_VALUE)
    Call SelectAllText(txtDescontoSubTotal(0))
  End If
  
  Me.Show vbModal
  
  
  'Retorna valor do desconto
  If blnDescontoRateado Then
    gnDesconto = CDbl(mcurDesconto)
  Else
    curDescConcedido = mcurDesconto
    curNewTotal = mcurTotalGeral
  End If
  
  Start = mblnPressOK
  
End Function

Private Sub cmdOK_Click()
  Dim curDesconto As Currency
  Dim curDescMax As Currency
  Dim intIndex As Integer
  
  
  intIndex = IIf(m_blnDescontoRateado, 1, 0)
  
  Call txtDescontoSubTotal_Validate(intIndex, False)
  Call txtPorcentagem_Validate(intIndex, False)
  
  curDesconto = CCur(txtDescontoSubTotal(intIndex).Text)
  
  If curDesconto <= 0 Or curDesconto >= mcurTotalGeral Then
    DisplayMsg "Valor incorreto."
    Call SelectAllText(txtDescontoSubTotal(intIndex), True)
    Exit Sub
  Else
    'Desconto máximo
    curDescMax = Format(mcurTotalGeral * msngMaxDescPerc / 100, "#0.00")
   
    If gParticipaProgramaFidelidade = 1 And gClienteEntregouResgatePontos = True And gSaldoCdGuidResgate > 0 Then
        '1-SIM PARTICIPA;
        '0-NÃO PARTICIPA Empresa/filial;
            
        If curDesconto - gSaldoCdGuidResgate > curDescMax Then
            DisplayMsg "Desconto superior ao permitido."
            Call SelectAllText(txtDescontoSubTotal(intIndex), True)
            Exit Sub
        End If
        
    Else
        'Fluxo normal sem o tratamento do prog. fidelidade...
        If curDesconto > curDescMax Then
          DisplayMsg "Desconto superior ao permitido."
          Call SelectAllText(txtDescontoSubTotal(intIndex), True)
          Exit Sub
        End If
    End If
   
   
  End If
  
  If Not m_blnDescontoRateado Then
    mcurTotalGeral = CCur(lblNovoTotalGeral.Caption)
  End If
  
  mcurDesconto = curDesconto
  
  mblnPressOK = True
  
  gSaldoCdGuidResgate_clicou_ok_telaDesconto = True
  Unload Me
  
End Sub

Private Sub cmdCancelar_Click()
  mblnPressOK = False
  Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo Erro:

  If gParticipaProgramaFidelidade = 1 Then
    '1-SIM PARTICIPA;
    '0-NÃO PARTICIPA Empresa/filial;
    If gClienteEntregouResgatePontos = True Then
        gSaldoCdGuidResgate_clicou_ok_telaDesconto = False

        lblDescontoProgFidelidade_SALDO.Visible = True
        lblDescontoProgFidelidade2_SALDO.Visible = True
        lblDescontoProgFidelidade_SALDO.Caption = "PROGRAMA FIDELIDADE DESC R$ " & CStr(Format(gSaldoCdGuidResgate, FORMAT_VALUE))
        lblDescontoProgFidelidade2_SALDO.Caption = "PROGRAMA FIDELIDADE DESC R$ " & CStr(Format(gSaldoCdGuidResgate, FORMAT_VALUE))
        txtDescontoSubTotal(0).Text = Format(gSaldoCdGuidResgate, FORMAT_VALUE)
        txtDescontoSubTotal(1).Text = Format(gSaldoCdGuidResgate, FORMAT_VALUE)

        txtPorcentagem(0).Visible = False
        txtPorcentagem(1).Visible = False
        lblTitle(3).Visible = False
        lblTitle(6).Visible = False
        sspInfo(0).Caption = "Digite o valor de desconto desejado que será concedido como desconto no sub total do cupom fiscal."
        sspInfo(2).Caption = "Digite o valor de desconto desejado que será concedido como desconto no sub total do cupom fiscal."

        Dim curDesconto As Variant
        If IsDataType(dtCurrency, gSaldoCdGuidResgate, curDesconto) Then
          lblNovoTotalGeral.Caption = Format(mcurTotalGeral - curDesconto, FORMAT_VALUE)
        End If
    Else
        lblDescontoProgFidelidade_SALDO.Visible = False
        lblDescontoProgFidelidade2_SALDO.Visible = False
    End If
  Else
    lblDescontoProgFidelidade_SALDO.Visible = False
    lblDescontoProgFidelidade2_SALDO.Visible = False
  End If
  
  Exit Sub

Erro:
  MsgBox "Erro na carga da tela (Sub Activate) " & Err.Number & " " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub Form_Load()
  
  Me.Width = 6355
  Me.Height = 4590
  
  lblTotalGeral.Caption = Format(0, FORMAT_VALUE)
  txtDescontoSubTotal(0).Text = Format(0, FORMAT_VALUE)
  txtDescontoSubTotal(1).Text = Format(0, FORMAT_VALUE)
  lblNovoTotalGeral.Caption = Format(0, FORMAT_VALUE)
  
End Sub

Private Sub txtDescontoSubTotal_Change(Index As Integer)
  Dim curDesconto As Currency
  
  If Index = 1 Then Exit Sub
  
  If IsDataType(dtCurrency, txtDescontoSubTotal(Index).Text, curDesconto) Then
    lblNovoTotalGeral.Caption = Format(mcurTotalGeral - curDesconto, FORMAT_VALUE)
  Else
    txtDescontoSubTotal(Index).Text = ""
    lblNovoTotalGeral.Caption = Format(mcurTotalGeral, FORMAT_VALUE)
  End If

End Sub

Private Sub txtDescontoSubTotal_GotFocus(Index As Integer)
  Call SelectAllText(txtDescontoSubTotal(Index))
End Sub

Private Sub txtDescontoSubTotal_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtDescontoSubTotal_Validate(Index As Integer, Cancel As Boolean)
  Dim curDesconto As Currency
  
  Call IsDataType(dtCurrency, txtDescontoSubTotal(Index).Text, curDesconto)
  txtDescontoSubTotal(Index).Text = Format(curDesconto, FORMAT_VALUE)
End Sub

Private Sub txtPorcentagem_Change(Index As Integer)
  Dim sngDescPerc As Single
  
  Call IsDataType(dtSingle, txtPorcentagem(Index).Text, sngDescPerc)
  If sngDescPerc < 0 Or sngDescPerc > 99.99 Then sngDescPerc = 0
  
  txtDescontoSubTotal(Index).Text = Format(mcurTotalGeral * sngDescPerc / 100, FORMAT_VALUE)
  
  If Index = 0 Then
    lblNovoTotalGeral.Caption = Format(mcurTotalGeral - CCur(txtDescontoSubTotal(Index).Text), FORMAT_VALUE)
  End If
  
End Sub

Private Sub txtPorcentagem_GotFocus(Index As Integer)
  Call SelectAllText(txtPorcentagem(Index))
End Sub

Private Sub txtPorcentagem_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

Private Sub txtPorcentagem_Validate(Index As Integer, Cancel As Boolean)
  Dim sngDescPerc As Single
  
  Call IsDataType(dtSingle, txtPorcentagem(Index).Text, sngDescPerc)
  If sngDescPerc < 0 Or sngDescPerc > 99.99 Then sngDescPerc = 0
  txtPorcentagem(Index).Text = sngDescPerc
End Sub
