VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosCalculoVendaSimples 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cálculo do Preço de Venda Simplificado"
   ClientHeight    =   4710
   ClientLeft      =   4005
   ClientTop       =   1065
   ClientWidth     =   10755
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
   HelpContextID   =   1660
   Icon            =   "CalculaVendaSimples.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4710
   ScaleWidth      =   10755
   Begin VB.Data datPrecos 
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
      Height          =   375
      Left            =   3315
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   7575
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CheckBox Preço_Zero 
      Appearance      =   0  'Flat
      Caption         =   "&Não copiar produtos com preço original igual a 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5550
      TabIndex        =   3
      Top             =   1590
      Width           =   5235
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   60
      TabIndex        =   6
      Top             =   3000
      Width           =   10620
      Begin VB.OptionButton Arredonda_1000 
         Appearance      =   0  'Flat
         Caption         =   "10.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7620
         TabIndex        =   13
         Top             =   690
         Width           =   975
      End
      Begin VB.OptionButton Arredonda_500 
         Appearance      =   0  'Flat
         Caption         =   "5.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5490
         TabIndex        =   12
         Top             =   690
         Width           =   810
      End
      Begin VB.OptionButton Arredonda_100 
         Appearance      =   0  'Flat
         Caption         =   "1.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2610
         TabIndex        =   11
         Top             =   690
         Width           =   765
      End
      Begin VB.OptionButton Arredonda_050 
         Appearance      =   0  'Flat
         Caption         =   "0.50"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7620
         TabIndex        =   10
         Top             =   255
         Width           =   795
      End
      Begin VB.OptionButton Arredonda_010 
         Appearance      =   0  'Flat
         Caption         =   "0.10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5490
         TabIndex        =   9
         Top             =   255
         Width           =   915
      End
      Begin VB.OptionButton Arredonda_005 
         Appearance      =   0  'Flat
         Caption         =   "Arredondar para 0.05"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   8
         Top             =   270
         Width           =   2595
      End
      Begin VB.OptionButton O_Sem_Arredondamento 
         Appearance      =   0  'Flat
         Caption         =   "Não arrendondar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   2565
      End
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
      Left            =   1680
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Sub_Classe"
      Top             =   7470
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Sobre 
      Appearance      =   0  'Flat
      Caption         =   "&Sobrepõe preços existentes na tabela destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1590
      Width           =   4905
   End
   Begin VB.CommandButton B_Calcula 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Calcular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4170
      Width           =   10620
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
      Left            =   75
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1725
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Sub 
      Bindings        =   "CalculaVendaSimples.frx":4E95A
      DataSource      =   "Data2"
      Height          =   405
      Left            =   5550
      TabIndex        =   5
      ToolTipText     =   "Use 0 para todas as Subclasses"
      Top             =   2535
      Width           =   1185
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
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8202
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2090
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   2090
      _ExtentY        =   714
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "CalculaVendaSimples.frx":4E96E
      DataSource      =   "Data1"
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   2535
      Width           =   1095
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
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9208
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1852
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1931
      _ExtentY        =   714
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabOrig 
      Bindings        =   "CalculaVendaSimples.frx":4E982
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   1050
      Width           =   4845
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      BackColorOdd    =   12648384
      Columns(0).Width=   3200
      _ExtentX        =   8546
      _ExtentY        =   714
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboTabDest 
      Bindings        =   "CalculaVendaSimples.frx":4E99A
      Height          =   405
      Left            =   5550
      TabIndex        =   1
      Top             =   1050
      Width           =   5115
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      BackColorOdd    =   12648384
      Columns(0).Width=   3200
      _ExtentX        =   9022
      _ExtentY        =   714
      _StockProps     =   93
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "O preço é aplicado nos produtos com base nos valores e percentuais registrados na ABA cálculos do cadastro de produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   60
      TabIndex        =   21
      Top             =   0
      Width           =   10620
   End
   Begin VB.Label Nome_Sub 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6810
      TabIndex        =   20
      Top             =   2535
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "SubClasse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5550
      TabIndex        =   19
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   18
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Nome_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1230
      TabIndex        =   17
      Top             =   2535
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Tabela destino dos preços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5550
      TabIndex        =   16
      Top             =   660
      Width           =   2985
   End
   Begin VB.Label Tabela 
      Appearance      =   0  'Flat
      Caption         =   "Tabela origem dos preços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      TabIndex        =   15
      Top             =   660
      Width           =   3075
   End
End
Attribute VB_Name = "frmPrecosCalculoVendaSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num_Registro As Variant
Dim rsPreços As Recordset
Dim rsPreços2 As Recordset
Dim rsClasses As Recordset
Dim rsSubclasses As Recordset
Dim rsProdutos As Recordset
Dim rsTabelas As Recordset
Private rsConta_Cli As Recordset

Private Sub Combo_Classe_CloseUp()
  Combo_Classe.Text = Combo_Classe.Columns(1).Text
  Combo_Classe_LostFocus
End Sub

'-----------------------------------------------------------------------------------
'05/07/2002 - mpdea
'Implementado o suporte a transação com tratamento a erro
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Calcula_Click()
  Dim Produto As Variant
  Dim Preço As Variant
  Dim Copiados As Long
  Dim Aux As Integer
  Dim i As Integer
  Dim nTempCopiados As Long
  
  Dim Str_Arredonda As String
  Dim Novo_Preço As Double
  
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler
  
  Copiados = 0
  Produto = 0

  Call StatusMsg("")
  
  If IsNull(cboTabOrig.Text) Or cboTabOrig.Text = "" Then
    DisplayMsg "Tabela de Origem inválida !"
    cboTabOrig.SetFocus
    Exit Sub
  End If

  If IsNull(cboTabDest.Text) Or cboTabDest.Text = "" Then
    DisplayMsg "Tabela Destino inválida !"
    cboTabDest.SetFocus
    Exit Sub
  End If

  cboTabDest.Text = Trim(cboTabDest.Text)

  If cboTabDest.Text = cboTabOrig.Text Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja efetuar as alterações na mesma tabela de preços?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
    Sobre.Value = vbChecked
  End If

  Screen.MousePointer = vbHourglass
  ws.BeginTrans
  blnOnTransaction = True
  
  If IsNull(Nome_Classe.Caption) Or Nome_Classe.Caption = "" Then Combo_Classe.Text = 0
  If IsNull(Nome_Sub.Caption) Or Nome_Sub.Caption = "" Then Combo_Sub.Text = 0


  Rem Começa a copiar as tabelas
  rsProdutos.Index = "Código"
  rsPreços.Index = "Tabela"
  rsPreços2.Index = "Tabela"
  
  Str_Arredonda = "000"
  If Arredonda_005.Value = True Then Str_Arredonda = "005"
  If Arredonda_010.Value = True Then Str_Arredonda = "010"
  If Arredonda_050.Value = True Then Str_Arredonda = "050"
  If Arredonda_100.Value = True Then Str_Arredonda = "100"
  If Arredonda_500.Value = True Then Str_Arredonda = "500"
  If Arredonda_1000.Value = True Then Str_Arredonda = "1000"
  

Lp1:
  If nTempCopiados <> Copiados Then
    nTempCopiados = Copiados
    Call StatusMsg("Foram copiados " & Copiados & " registros.")
  End If
  rsPreços.Seek ">", cboTabOrig.Text, Produto
  If rsPreços.NoMatch Then
    Aux = 1
    GoTo Fim
  End If
  If rsPreços("Tabela") <> cboTabOrig.Text Then
    Aux = 2
    GoTo Fim
  End If

  Produto = rsPreços("Produto")
  
  rsProdutos.Seek "=", Produto
  If rsProdutos.NoMatch Then GoTo Lp1

  If Preço_Zero.Value = 1 Then
    If rsPreços("Preço") = 0 Then
      GoTo Lp1
    End If
  End If


  Rem Verifica se é da classe desejada
  If Val(Combo_Classe.Text) <> 0 Then
    If rsProdutos("Classe") <> Val(Combo_Classe.Text) Then
      GoTo Lp1
    End If
  End If

  Rem Verifica se é da sub classe desejada
  If Val(Combo_Sub.Text) <> 0 Then
    If rsProdutos("Sub Classe") <> Val(Combo_Sub.Text) Then
      GoTo Lp1
    End If
  End If
  
  'Verificações para cálculo do preço
  If rsPreços("Preço") <= 0 Or rsProdutos("Custo Lucro Perc") <= 0 Then
    GoTo Lp1
  End If

  Novo_Preço = rsPreços("Preço") * (1 + rsProdutos("Custo Lucro Perc") / 100)
  Novo_Preço = Arredonda_Valor(Novo_Preço, Str_Arredonda)

  rsPreços2.Seek "=", cboTabDest.Text, rsPreços("Produto")
  If Not rsPreços2.NoMatch Then
    If Sobre.Value = 0 Then
      GoTo Lp1
    End If

    rsPreços2.Edit
    rsPreços2("Preço") = Format(Novo_Preço, "#############0.00")
    rsPreços2("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsPreços2.Update
  
    'Atualiza o sincronismo para o produto WEB alterado
    Call WEB_SynchronizeProduct(rsPreços("Produto").Value)
        
    Copiados = Copiados + 1
    GoTo Lp1
  End If


  rsPreços2.AddNew
  
  rsPreços2("Tabela") = cboTabDest.Text
  rsPreços2("Produto") = rsPreços("Produto")
  rsPreços2("Preço") = Format(Novo_Preço, "#############0.00")
  rsPreços2("Data Alteração") = Format(Date, "dd/mm/yyyy")
  
  rsPreços2.Update
  
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(rsPreços("Produto").Value)
  
  Copiados = Copiados + 1

  GoTo Lp1

Fim:
 
  'Cria configuração da tabela
  Call CheckConfigTablePrice(cboTabDest.Text)
  
  ws.CommitTrans
  blnOnTransaction = False
  
  datPrecos.Refresh
  cboTabOrig.Refresh
  cboTabDest.Refresh
  
  cboTabDest.Text = ""
  
  Screen.MousePointer = vbDefault
  
  DisplayMsg "Final de processo. Copiados " & Copiados & " registros."
  
  Call StatusMsg("")

  Exit Sub

ErrHandler:
  Screen.MousePointer = vbDefault
  If blnOnTransaction Then ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Combo_Classe_LostFocus()
  Nome_Classe.Caption = ""
  If IsNull(Combo_Classe.Text) Then Exit Sub
  If Not IsNumeric(Combo_Classe.Text) Then Exit Sub

  rsClasses.Index = "Código"
  rsClasses.Seek "=", Combo_Classe.Text
  If Not rsClasses.NoMatch Then
     Nome_Classe.Caption = rsClasses("Nome")
  Else
     Combo_Classe.Text = 0
  End If

End Sub

Private Sub Combo_Sub_CloseUp()
 Combo_Sub.Text = Combo_Sub.Columns(1).Text
 Combo_Sub_LostFocus
End Sub

Private Sub Combo_Sub_LostFocus()
  Nome_Sub.Caption = ""
  If IsNull(Combo_Sub.Text) Then Exit Sub
  If Not IsNumeric(Combo_Sub.Text) Then Exit Sub

  rsSubclasses.Index = "Código"
  rsSubclasses.Seek "=", Combo_Sub.Text
  If Not rsSubclasses.NoMatch Then
     Nome_Sub.Caption = rsSubclasses("Nome")
  Else
     Combo_Sub.Text = 0
  End If
End Sub

Private Sub Form_Load()

  Call CenterForm(Me)
  
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsPreços2 = db.OpenRecordset("Preços")
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Preços")
  Set rsConta_Cli = db.OpenRecordset("SELECT * FROM [Conta Cliente]", dbOpenDynaset)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  datPrecos.DatabaseName = gsQuickDBFileName

End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsPreços.Close
  rsPreços2.Close
  rsClasses.Close
  rsSubclasses.Close
  rsProdutos.Close
  rsTabelas.Close
  rsConta_Cli.Close
  Set rsPreços = Nothing
  Set rsPreços2 = Nothing
  Set rsClasses = Nothing
  Set rsSubclasses = Nothing
  Set rsProdutos = Nothing
  Set rsTabelas = Nothing
  Set rsConta_Cli = Nothing
End Sub

Private Sub Multiplicador_KeyPress(KeyAscii As Integer)
  KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub cboTabDest_KeyPress(KeyAscii As Integer)
  KeyAscii = gnLimitKeyPress(cboTabDest, 15, KeyAscii)
  If KeyAscii <> 0 Then
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub cboTabDest_LostFocus()
  If IsNull(cboTabDest.Text) Then Exit Sub
  cboTabDest.Text = UCase$(cboTabDest.Text)
'  If cboTabDest.Text = cboTabOrig.Text And Len(cboTabOrig.Text) > 0 Then
'    DisplayMsg "Aviso: As alterações serão realizadas na mesma tabela e não existe um desfaz automático."
'  End If
End Sub

