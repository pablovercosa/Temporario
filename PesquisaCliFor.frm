VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPesquisaCliFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pesquisa de Clientes e Fornecedores"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PesquisaCliFor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11220
   Begin VB.CommandButton cmd_alterarTamanhoTela 
      Height          =   525
      Left            =   10470
      Picture         =   "PesquisaCliFor.frx":4E95A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2040
      Width           =   705
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipo "
      Height          =   2040
      Left            =   6930
      TabIndex        =   19
      Top             =   0
      Width           =   4275
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
         Height          =   390
         Left            =   3330
         Picture         =   "PesquisaCliFor.frx":4FA90
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1020
         Visible         =   0   'False
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
         Height          =   390
         Left            =   3330
         Picture         =   "PesquisaCliFor.frx":50372
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Marcado com Pendência, mas sem parcelas vencidas"
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   6
         Left            =   1620
         TabIndex        =   26
         Tag             =   "T"
         Top             =   1470
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Com parcelas vencidas e não pagas no período"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   1620
         TabIndex        =   25
         Tag             =   "T"
         Top             =   210
         Width           =   2460
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Tag             =   "T"
         Top             =   1620
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Outros"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Tag             =   "O"
         Top             =   1290
         Width           =   900
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Revendedor"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Tag             =   "R"
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F&ornecedores"
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Tag             =   "F"
         Top             =   615
         Width           =   1380
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clien&tes"
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Tag             =   "C"
         Top             =   272
         Width           =   960
      End
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   285
         Left            =   2130
         TabIndex        =   32
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   1050
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
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
         Height          =   285
         Left            =   2145
         TabIndex        =   33
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   660
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "De"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         TabIndex        =   35
         Top             =   675
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Até"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1860
         TabIndex        =   34
         Top             =   1065
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "Interromper"
      Enabled         =   0   'False
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
      Left            =   10320
      TabIndex        =   16
      Top             =   1830
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pesquisar"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2070
      Width           =   10335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2040
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   6825
      Begin VB.Frame fraEntrega 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buscar registros que"
         Height          =   1080
         Left            =   4230
         TabIndex        =   27
         Top             =   900
         Width           =   2565
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Contenham &TODAS as partes"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   29
            Top             =   270
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Contenham uma &OU outra parte"
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   1
            Left            =   60
            TabIndex        =   28
            Top             =   570
            Width           =   2265
         End
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CNPJ/CPF"
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   2
         Top             =   930
         Width           =   1065
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   1830
         TabIndex        =   3
         Top             =   885
         Width           =   2355
      End
      Begin VB.OptionButton optSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fa&x"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   5940
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fone &2"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   5100
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   4650
         TabIndex        =   14
         Top             =   540
         Width           =   1485
      End
      Begin VB.OptionButton optSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fone &1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   4260
         TabIndex        =   11
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1830
         TabIndex        =   8
         Top             =   1245
         Width           =   2355
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C&idade"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   7
         Top             =   1290
         Width           =   885
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1830
         TabIndex        =   10
         Top             =   1620
         Width           =   2355
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Bens Numeráveis"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   1665
         Width           =   1695
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Fantasia"
         Height          =   225
         Index           =   2
         Left            =   855
         TabIndex        =   5
         Top             =   585
         Width           =   975
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   1830
         TabIndex        =   6
         Top             =   532
         Width           =   2355
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1830
         TabIndex        =   1
         Top             =   172
         Width           =   2355
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Nome"
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   525
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Código"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   930
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultados 
      Height          =   2925
      Left            =   60
      TabIndex        =   18
      ToolTipText     =   "Selecione a linha e dê duplo-clique para posicionamento."
      Top             =   2580
      Width           =   11115
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   4
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      ForeColorEven   =   4210752
      BackColorOdd    =   16777152
      RowHeight       =   370
      ExtraHeight     =   265
      Columns.Count   =   4
      Columns(0).Width=   2090
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6879
      Columns(1).Caption=   "Descrição"
      Columns(1).Name =   "Descricao"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4789
      Columns(2).Caption=   "Descrição_2"
      Columns(2).Name =   "Descricao_2"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4789
      Columns(3).Caption=   "Descrição_3"
      Columns(3).Name =   "Descricao_3"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   19606
      _ExtentY        =   5159
      _StockProps     =   79
      BackColor       =   15066597
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "frmPesquisaCliFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsCliFor As Recordset
Private rsCliForNumeravel As Recordset
Private gsSql As String
Private gsCod() As String
Private gsDesc() As String
Private gsDesc2() As String
Private gsDesc3() As String
Private gbToAbort As Boolean
Public iOrigemSaidas As Boolean
Public iOrigemVendaRapida As Boolean

Dim arrayInadimplentes() As Variant
Dim contador_arrayInadimplentes As Long

Private Function AchaInadimplentes(pCliente As Long) As String
  Dim l As Long
  AchaInadimplentes = ""
  
  For l = 0 To contador_arrayInadimplentes - 1
      If arrayInadimplentes(l, 0) = pCliente Then
          AchaInadimplentes = arrayInadimplentes(l, 0)
          Exit For
      End If
  Next
End Function


Private Sub cmd_alterarTamanhoTela_Click()
    If Me.Height < 6000 Then
        Me.Top = Me.Top - 1700
        Me.Height = Me.Height + 2000
        grdResultados.Height = grdResultados.Height + 2000
    Else
        Me.Height = Me.Height - 2000
        grdResultados.Height = grdResultados.Height - 2000
    End If
End Sub

Private Sub cmd_calendarioDtFim_Click()
    Data_Fim.Text = frmCalendario.gsDateCalender(Data_Fim.Text)
End Sub

Private Sub cmd_calendarioDtIni_Click()
    Data_Ini.Text = frmCalendario.gsDateCalender(Data_Ini.Text)
End Sub

''''Private Sub cmdAbort_Click()
''''  gbToAbort = True
''''End Sub

Private Sub cmdSearch_Click()
On Error GoTo Erro
  Dim sSql As String
  
  If optTipo(5).Value = True Or optTipo(6).Value = True Then
        Dim rsContasReceber As Recordset
        
        grdResultados.RemoveAll
        
        If optTipo(5).Value = True Then
        
            If IsNull(Data_Ini.Text) Then
              DisplayMsg "Data incorreta, verifique."
              Data_Ini.SetFocus
              Exit Sub
            End If
  
            If Not IsDate(Data_Ini.Text) Then
              DisplayMsg "Data incorreta, verifique."
              Data_Ini.SetFocus
              Exit Sub
            End If
          
            If IsNull(Data_Fim.Text) Then
              DisplayMsg "Data incorreta, verifique."
              Data_Fim.SetFocus
              Exit Sub
            End If
              
            If Not IsDate(Data_Fim.Text) Then
              DisplayMsg "Data incorreta, verifique."
              Data_Fim.SetFocus
              Exit Sub
            End If
              
            If CDate(Data_Ini.Text) > CDate(Data_Fim.Text) Then
              DisplayMsg "Data inicial deve ser menor ou igual a data final."
              Data_Ini.SetFocus
              Exit Sub
            End If
            
            sSql = "Select distinct (R.Cliente), C.Nome, C.Fantasia, C.CGC "
            sSql = sSql & " From [Contas a Receber] R, Cli_For C "
            sSql = sSql & " Where R.Vencimento >= CDATE('" & Data_Ini & " 00:00:00') and "
            sSql = sSql & " R.Vencimento <= CDATE('" & Data_Fim & " 23:59:59') and "
            sSql = sSql & " R.Filial = " & gnCodFilial & " And "
            sSql = sSql & " R.[Valor Recebido] = 0 And "
            sSql = sSql & " R.Cliente = C.Código "
            sSql = sSql & " Order by C.Nome "
            
            Set rsContasReceber = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
  
            If Not (rsContasReceber.EOF And rsContasReceber.BOF) Then
              rsContasReceber.MoveFirst
            End If
  
            grdResultados.Redraw = False
            
            While Not rsContasReceber.EOF
                grdResultados.AddItem rsContasReceber.Fields("Cliente").Value & vbTab & _
                                    rsContasReceber.Fields("Nome").Value & vbTab & _
                                    rsContasReceber.Fields("Fantasia").Value & vbTab & _
                                    rsContasReceber.Fields("CGC").Value
                
                rsContasReceber.MoveNext
            Wend
            rsContasReceber.Close
            Set rsContasReceber = Nothing
            grdResultados.Redraw = True
        Else
            ' Cliente esta marcado com PENDENCIA (tipo SPC, SERASA, etc), mas não tem parcelas vencidas no quick
            Dim rsInadimplentes As Recordset
            Dim lContador       As Long
            lContador = 0
            Dim sClienteInad    As String
            
            'Busca inadimplentes
            sSql = "Select distinct (R.Cliente) "
            sSql = sSql & " From [Contas a Receber] R, Cli_For C "
            sSql = sSql & " Where R.Filial = " & gnCodFilial & " And "
            sSql = sSql & " R.[Valor Recebido] = 0 And "
            sSql = sSql & " R.Vencimento < CDATE('" & Data_Atual + 1 & "') and "
            'sSQL = sSQL & " R.Vencimento < CDATE('" & Data_Atual & " 23:59:59') and "
            sSql = sSql & " R.Cliente = C.Código "
            
            Set rsInadimplentes = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
            If Not (rsInadimplentes.EOF And rsInadimplentes.BOF) Then
                rsInadimplentes.MoveLast
                rsInadimplentes.MoveFirst
                
                ReDim arrayInadimplentes(rsInadimplentes.RecordCount, 2)
                contador_arrayInadimplentes = rsInadimplentes.RecordCount
                While Not rsInadimplentes.EOF
                    arrayInadimplentes(lContador, 0) = rsInadimplentes.Fields(0).Value
                    
                    lContador = lContador + 1
                    rsInadimplentes.MoveNext
                Wend
            End If
            rsInadimplentes.Close
            Set rsInadimplentes = Nothing
            
            'Busca marcados com PENDENCIA
            sSql = "Select Código, Nome, Fantasia, CGC From Cli_For Where Pendencia = -1 "
            
            Set rsContasReceber = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
  
            If Not (rsContasReceber.EOF And rsContasReceber.BOF) Then
              rsContasReceber.MoveFirst
            End If
  
            grdResultados.Redraw = False
            
            While Not rsContasReceber.EOF
            
                sClienteInad = AchaInadimplentes(rsContasReceber.Fields("Código").Value)
                
                If Trim(sClienteInad) = "" Then
                    'Mostra quem não esta inadimplente
                    grdResultados.AddItem rsContasReceber.Fields("Código").Value & vbTab & _
                                        rsContasReceber.Fields("Nome").Value & vbTab & _
                                        rsContasReceber.Fields("Fantasia").Value & vbTab & _
                                        rsContasReceber.Fields("CGC").Value
                End If
                
                rsContasReceber.MoveNext
            Wend
            rsContasReceber.Close
            Set rsContasReceber = Nothing
            grdResultados.Redraw = True
        End If
        
  Else
      If optSearch(1).Value = True Or optSearch(2).Value = True Then
        If Len(txtSearch(1).Text) < 3 Then
            DisplayMsg "Digite pelo menos 3 caracteres"
            txtSearch(1).SetFocus
            Exit Sub
        End If
      ElseIf optSearch(8).Value = True Then
        If Len(txtSearch(5).Text) < 3 Then
            DisplayMsg "Digite pelo menos 3 caracteres"
            txtSearch(5).SetFocus
            Exit Sub
        End If
      ElseIf optSearch(4).Value = True Then
        If Len(txtSearch(3).Text) < 3 Then
            DisplayMsg "Digite pelo menos 3 caracteres"
            txtSearch(3).SetFocus
            Exit Sub
        End If
      ElseIf optSearch(5).Value = True Or optSearch(6).Value = True Or optSearch(7).Value = True Then
        If Len(txtSearch(4).Text) < 3 Then
            DisplayMsg "Digite pelo menos 3 caracteres"
            txtSearch(4).SetFocus
            Exit Sub
        End If
      End If
  
      Call SearchString
    
      If grdResultados.Rows > 0 Then
          grdResultados.SetFocus
      Else
          If optSearch(1).Value = True Or optSearch(2).Value = True Then
              txtSearch(1).SetFocus
          ElseIf optSearch(8).Value = True Then
              txtSearch(5).SetFocus
          ElseIf optSearch(0).Value = True Then
              txtSearch(0).SetFocus
          End If
      End If
  End If
  
  Exit Sub
Erro:
    MsgBox "Erro " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub Form_Load()
  'Call CenterForm(Me)
  Screen.MousePointer = vbHourglass
  Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Código", dbReadOnly)
  Set rsCliForNumeravel = db.OpenRecordset("SELECT * FROM CliForNumeravel ORDER BY CodCliNumer", dbReadOnly)
  Screen.MousePointer = vbDefault
  cmdAbort.Enabled = False
  cmdSearch.Enabled = False
  
  Me.Left = Screen.Width - Me.Width
  
  If Screen.Height - (Me.Height + 350) > 0 Then
      Me.Top = Screen.Height - (Me.Height + 350)
  Else
      Me.Top = 500
  End If
  
  Me.Show
  txtSearch(1).SetFocus
  
End Sub

Private Function retiraCaracteresNaoNumericos(ByVal sDado As String) As String
    sDado = Replace(sDado, "-", "")
    sDado = Replace(sDado, ".", "")
    sDado = Replace(sDado, ",", "")
    sDado = Replace(sDado, "/", "")
    sDado = Replace(sDado, "\", "")
    sDado = Replace(sDado, "_", "")
    sDado = Replace(sDado, ";", "")
    sDado = Replace(sDado, " ", "")
    sDado = Replace(sDado, "|", "")
    sDado = Replace(sDado, ">", "")
    sDado = Replace(sDado, "<", "")
    sDado = Replace(sDado, "(", "")
    sDado = Replace(sDado, ")", "")
    sDado = Replace(sDado, "[", "")
    sDado = Replace(sDado, "]", "")
    sDado = Replace(sDado, "º", "")
    sDado = Replace(sDado, "ª", "")
    sDado = Replace(sDado, "=", "")
    sDado = Replace(sDado, "+", "")
    sDado = Replace(sDado, "!", "")
    sDado = Replace(sDado, "@", "")
    sDado = Replace(sDado, "#", "")
    sDado = Replace(sDado, "$", "")
    sDado = Replace(sDado, "%", "")
    sDado = Replace(sDado, "&", "")
    sDado = Replace(sDado, "*", "")
    
    sDado = Replace(sDado, vbCrLf, "")

    sDado = LTrim(RTrim(sDado))
    
    retiraCaracteresNaoNumericos = sDado
  
End Function

Private Sub SearchString()
  Dim nPos As Long
  Dim sCod As String
  Dim sText() As String
  Dim nI As Long
  Dim nK As Long
  Dim NF As Long
  Dim nItem As Long
  Dim nSum As Long
  Dim sTipo As String
  
  Screen.MousePointer = vbHourglass
  grdResultados.RemoveAll
  Erase gsCod
  Erase gsDesc
  Erase gsDesc2
  Erase gsDesc3
  
  
  nK = -1
  
  If optSearch(0).Value = True Then
    nK = 0
    NF = 0
  Else
    If optSearch(1).Value = True Then
      nK = 1
      NF = 1
    Else
      If optSearch(2).Value = True Then
        nK = 2
        NF = 1
      Else
        If optSearch(4).Value = True Then
          nK = 3
          NF = 3
         Else
           If optSearch(5).Value = True Then
              nK = 4
              NF = 4
           Else
              If optSearch(6).Value = True Then
                  nK = 5
                  NF = 4
              Else
                If optSearch(7).Value = True Then
                  nK = 6
                  NF = 4
                Else
                  If optSearch(8).Value = True Then '11/11/2004 - Daniel
                    nK = 8
                    NF = 5
                  End If
               End If
            End If
           End If
        End If
      End If
    End If
  End If
  
  nItem = -1
  gbToAbort = False
  cmdAbort.Enabled = True
  
  sTipo = "C"
  For nI = 0 To 4
    If optTipo(nI).Value = True Then
      sTipo = optTipo(nI).Tag
      Exit For
    End If
  Next nI
  
  If nK > -1 Then
    txtSearch(NF).Text = Trim(txtSearch(NF).Text & "")
    sText = Split(UCase(txtSearch(NF).Text), " ", -1, vbTextCompare)
    If Not rsCliFor.EOF Then
      With rsCliFor
        .MoveFirst
        Do While Not .EOF
          DoEvents
          If gbToAbort Then
            Exit Do
          End If
          nSum = 0
          If sTipo = "T" Or rsCliFor("Tipo").Value = sTipo Then
            For nI = 0 To UBound(sText)
              DoEvents
              If gbToAbort Then
                Exit For
              End If
              Select Case nK
                Case 0
                  nPos = InStr(CStr(.Fields("Código").Value), sText(nI))
                Case 1
                  nPos = InStr(UCase(.Fields("Nome").Value), sText(nI))
                Case 2
                  nPos = InStr(UCase(.Fields("Fantasia").Value & ""), sText(nI))
                Case 3
                  nPos = InStr(UCase(.Fields("Cidade").Value & ""), sText(nI))
                Case 4
                  nPos = InStr(UCase(.Fields("Fone 1").Value & ""), sText(nI))
                Case 5
                  nPos = InStr(UCase(.Fields("Fone 2").Value & ""), sText(nI))
                Case 6
                  nPos = InStr(UCase(.Fields("Fax").Value & ""), sText(nI))
                Case 8 '11/11/2004 - Daniel
                  If Not IsNull(.Fields("CGC").Value) Then
                      nPos = InStr(UCase(retiraCaracteresNaoNumericos(.Fields("CGC").Value) & ""), retiraCaracteresNaoNumericos(sText(nI)))
                  End If
              End Select
              If nPos > 0 Then
                nSum = nSum + 1
              End If
            Next nI
            If (optType(1).Value = True And nSum > 0) Or (nSum = UBound(sText) + 1) Then
              nItem = nItem + 1
              ReDim Preserve gsCod(nItem)
              ReDim Preserve gsDesc(nItem)
              ReDim Preserve gsDesc2(nItem)
              ReDim Preserve gsDesc3(nItem)
              gsCod(nItem) = .Fields("Código")
              gsDesc(nItem) = .Fields("Nome")
              
              If IsNull(.Fields("fantasia")) Then
                  gsDesc2(nItem) = ""
              Else
                  gsDesc2(nItem) = .Fields("Fantasia")
              End If
              
              If IsNull(.Fields("CGC")) Then
                  gsDesc3(nItem) = ""
              Else
                  gsDesc3(nItem) = .Fields("CGC")
              End If
              
            
            End If
          Else
            sTipo = sTipo
          End If
          .MoveNext
        Loop
      End With
    End If
  Else
    rsCliForNumeravel.FindFirst "CodNumer = '" & UCase(txtSearch(2).Text) & "'"
    If Not rsCliForNumeravel.NoMatch Then
      nItem = nItem + 1
      ReDim Preserve gsCod(nItem)
      ReDim Preserve gsDesc(nItem)
      ReDim Preserve gsDesc2(nItem)
      ReDim Preserve gsDesc3(nItem)
      gsCod(nItem) = rsCliForNumeravel.Fields("CodCliNumer")
      Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código = " & gsCod(nItem), dbReadOnly)
      If Not rsCliFor.EOF Then
        gsDesc(nItem) = rsCliFor.Fields("Nome")
        gsDesc2(nItem) = rsCliFor.Fields("Fantasia")
        gsDesc3(nItem) = rsCliFor.Fields("CGC")
      Else
        gsDesc(nItem) = ""
        gsDesc2(nItem) = ""
        gsDesc3(nItem) = ""
      End If
    End If
  End If
  
  If nItem > -1 Then
    Call LoadGrid
  Else
    gsTitle = "Resultados da Pesquisa"
    gsMsg = "Nenhum registro encontrado para as condições fornecidas."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  On Error Resume Next
  rsCliFor.MoveFirst
  rsCliForNumeravel.MoveFirst
  cmdAbort.Enabled = False
  Screen.MousePointer = vbDefault
  On Error GoTo 0
End Sub

Private Sub LoadGrid()
  Dim nI As Integer
  '''Call CenterForm(Me)
  '''frmPesquisaCliFor.Refresh
  grdResultados.Redraw = False
  For nI = 0 To UBound(gsCod)
    grdResultados.AddItem gsCod(nI) & vbTab & gsDesc(nI) & vbTab & gsDesc2(nI) & vbTab & gsDesc3(nI)
  Next nI
  grdResultados.Redraw = True
  Erase gsCod
  Erase gsDesc
  Erase gsDesc2
  Erase gsDesc3
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsCliFor.Close
  Set rsCliFor = Nothing
  rsCliForNumeravel.Close
  Set rsCliForNumeravel = Nothing
End Sub

Private Sub grdResultados_Click()
  If iOrigemSaidas = True Then
    frmSaidas.cboCliente = grdResultados.Columns(0).Text
    frmSaidas.Nome_Cliente = grdResultados.Columns(1).Text
  ElseIf iOrigemVendaRapida = True Then
    frmVendaRap2.Combo_Cliente = grdResultados.Columns(0).Text
    frmVendaRap2.Nome_Cliente = grdResultados.Columns(1).Text
  End If
End Sub

Private Sub grdResultados_DblClick()
  If iOrigemSaidas = True Then
    frmSaidas.cboCliente = grdResultados.Columns(0).Text
    frmSaidas.Nome_Cliente = grdResultados.Columns(1).Text
  ElseIf iOrigemVendaRapida = True Then
    frmVendaRap2.Combo_Cliente = grdResultados.Columns(0).Text
    frmVendaRap2.Nome_Cliente = grdResultados.Columns(1).Text
  Else
    'Veio da tela de Cadastro de Cliente
    frmCliFor.cboCodigo.Text = grdResultados.Columns(0).Text
    frmCliFor.cboCodigo.DoClick
    grdResultados.SetFocus
  End If
End Sub

Private Sub grdResultados_KeyPress(KeyAscii As Integer)
  Dim cCaracter As Variant
   cCaracter = Chr(KeyAscii)
   KeyAscii = Asc(UCase(cCaracter))
   
   If KeyAscii = 13 Then    'Tecla ENTER
         grdResultados_Click
   End If
End Sub

Private Sub optSearch_Click(Index As Integer)
  
  Select Case Index
    Case 0
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = True
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(1).Value = True
      txtSearch(0).SetFocus
    Case 1
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Enabled = True
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(1).Value = True
      txtSearch(1).SetFocus
    Case 2
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Enabled = True
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(1).Value = True
      txtSearch(1).SetFocus
    Case 3
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = True
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(1).Value = True
      txtSearch(2).SetFocus
    Case 4
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = True
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(0).Value = True
      txtSearch(3).SetFocus
   Case 5
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      '13/07/2007 - Anderson
      'Eliminada a opção que apaga o texto procurado quando a seleção é feita em Fone1, fone2 e fax, facilitando a pesquisa nestes campos e evitando a redigitação.
      'Solicitado: Supriprint
      'txtSearch(4).Text = ""
      txtSearch(4).Enabled = True
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(0).Value = True
      txtSearch(4).SetFocus
  Case 6
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      '13/07/2007 - Anderson
      'Eliminada a opção que apaga o texto procurado quando a seleção é feita em Fone1, fone2 e fax, facilitando a pesquisa nestes campos e evitando a redigitação.
      'Solicitado: Supriprint
      'txtSearch(4).Text = ""
      txtSearch(4).Enabled = True
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(0).Value = True
      txtSearch(4).SetFocus
   Case 7
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      '13/07/2007 - Anderson
      'Eliminada a opção que apaga o texto procurado quando a seleção é feita em Fone1, fone2 e fax, facilitando a pesquisa nestes campos e evitando a redigitação.
      'Solicitado: Supriprint
      'txtSearch(4).Text = ""
      txtSearch(4).Enabled = True
      txtSearch(5).Text = ""
      txtSearch(5).Enabled = False
      optType(0).Value = True
      txtSearch(4).SetFocus
   Case 8 '10/11/2004 - Daniel
      txtSearch(0).Text = ""
      txtSearch(0).Enabled = False
      txtSearch(1).Text = ""
      txtSearch(1).Enabled = False
      txtSearch(2).Text = ""
      txtSearch(2).Enabled = False
      txtSearch(3).Text = ""
      txtSearch(3).Enabled = False
      txtSearch(4).Text = ""
      txtSearch(4).Enabled = False
      txtSearch(5).Enabled = True
      optType(1).Value = True
      txtSearch(5).SetFocus
  End Select
End Sub

Private Sub optTipo_Click(Index As Integer)

  Screen.MousePointer = vbHourglass
  
  If optSearch(3).Value = False Then
  
    Select Case Index
      Case 0
        Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Tipo = 'C' And Código <> 0 ORDER BY Código", dbReadOnly)
      Case 1
        Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Tipo = 'F'  And Código <> 0 ORDER BY Código", dbReadOnly)
      Case 2
        Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Tipo = 'R' And Código <> 0 ORDER BY Código", dbReadOnly)
      Case 3
        Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Tipo = 'O'  And Código <> 0 ORDER BY Código", dbReadOnly)
      Case 4
        Set rsCliFor = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código <> 0 ORDER BY Código", dbReadOnly)
    End Select
    
  Else
      Select Case Index
          Case 0
            Set rsCliForNumeravel = db.OpenRecordset("SELECT * FROM CliForNumeravel WHERE TipoCliNumer = 'C' ORDER BY CodCliNumer", dbReadOnly)
          Case 1
            Set rsCliForNumeravel = db.OpenRecordset("SELECT * FROM CliForNumeravel WHERE TipoCliNumer = 'F'  ORDER BY CodCliNumer", dbReadOnly)
          Case 2
            Set rsCliForNumeravel = db.OpenRecordset("SELECT * FROM CliForNumeravel WHERE TipoCliNumer = 'R' ORDER BY CodCliNumer", dbReadOnly)
          Case 3
            Set rsCliForNumeravel = db.OpenRecordset("SELECT * FROM CliForNumeravel WHERE TipoCliNumer = 'O'  ORDER BY CodCliNumer", dbReadOnly)
      End Select
  End If
  
  If optTipo(5).Value = True Or optTipo(6).Value = True Then
      If optTipo(5).Value = True Then
          Label5.Visible = True
          Label2.Visible = True
          Data_Ini.Visible = True
          Data_Fim.Visible = True
          cmd_calendarioDtIni.Visible = True
          cmd_calendarioDtFim.Visible = True
      Else
          Label5.Visible = False
          Label2.Visible = False
          Data_Ini.Visible = False
          Data_Fim.Visible = False
          cmd_calendarioDtIni.Visible = False
          cmd_calendarioDtFim.Visible = False
      End If
      
      cmdSearch.Enabled = True
      
      optSearch(0).Enabled = False
      optSearch(1).Enabled = False
      optSearch(2).Enabled = False
      optSearch(3).Enabled = False
      optSearch(4).Enabled = False
      optSearch(5).Enabled = False
      optSearch(6).Enabled = False
      optSearch(7).Enabled = False
      optSearch(8).Enabled = False
      optType(0).Enabled = False
      optType(1).Enabled = False
      Frame1.Enabled = False
  Else
      Label5.Visible = False
      Label2.Visible = False
      Data_Ini.Visible = False
      Data_Fim.Visible = False
      cmd_calendarioDtIni.Visible = False
      cmd_calendarioDtFim.Visible = False
   
      optSearch(0).Enabled = True
      optSearch(1).Enabled = True
      optSearch(2).Enabled = True
      optSearch(3).Enabled = True
      optSearch(4).Enabled = True
      optSearch(5).Enabled = True
      optSearch(6).Enabled = True
      optSearch(7).Enabled = True
      optSearch(8).Enabled = True
      optType(0).Enabled = True
      optType(1).Enabled = True
      Frame1.Enabled = True
   
  End If
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub txtSearch_Change(Index As Integer)
  If Len(Trim(txtSearch(Index).Text)) > 0 Then
    cmdSearch.Enabled = True
    cmdSearch.Default = True
  Else
    cmdSearch.Enabled = False
    cmdSearch.Default = False
  End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    With txtSearch(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
