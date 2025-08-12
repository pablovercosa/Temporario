VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManContas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Realizar Recebimento de Contas do Cliente"
   ClientHeight    =   7170
   ClientLeft      =   540
   ClientTop       =   690
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1700
   Icon            =   "ManContasCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   11550
   Begin SSDataWidgets_B.SSDBGrid grdContaClientes 
      Height          =   4170
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   11340
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
      Col.Count       =   8
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   16777152
      RowHeight       =   423
      ExtraHeight     =   53
      Columns.Count   =   8
      Columns(0).Width=   1746
      Columns(0).Caption=   "Data"
      Columns(0).Name =   "Data"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   10
      Columns(0).Mask =   "##/##/####"
      Columns(0).PromptInclude=   -1  'True
      Columns(0).PromptChar=   32
      Columns(1).Width=   3413
      Columns(1).Caption=   "Produto"
      Columns(1).Name =   "Produto"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4842
      Columns(2).Caption=   "Descrição"
      Columns(2).Name =   "Descrição"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   979
      Columns(3).Caption=   "Qtde"
      Columns(3).Name =   "Qtde"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1640
      Columns(4).Caption=   "Valor"
      Columns(4).Name =   "Valor"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).NumberFormat=   "##,###,##0.00"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1826
      Columns(5).Caption=   "Valor Pago"
      Columns(5).Name =   "Valor Pago"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).NumberFormat=   "##,###,##0.00"
      Columns(5).FieldLen=   256
      Columns(6).Width=   1746
      Columns(6).Caption=   "Pagamento"
      Columns(6).Name =   "Data Pagamento"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   10
      Columns(6).Mask =   "##/##/####"
      Columns(6).PromptInclude=   -1  'True
      Columns(6).PromptChar=   32
      Columns(7).Width=   2672
      Columns(7).Caption=   "Tab Preços"
      Columns(7).Name =   "Tab Preços"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   20002
      _ExtentY        =   7355
      _StockProps     =   79
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
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Left            =   2370
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Caixas"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   7245
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton B_Cancela 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2310
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton B_Prossegue 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Prosse&guir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1845
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Recebe_Valor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Receber &Parcial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1380
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Recebe_Todas 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Receber &Tudo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Recebe todas as compras"
      Top             =   1380
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton B_Monta 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Pesquisar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   675
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar"
      Height          =   675
      Left            =   120
      TabIndex        =   15
      Top             =   525
      Width           =   7695
      Begin VB.OptionButton O_Não_Pagas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Compras não pagas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4860
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton O_Pagas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Compras já pagas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2730
         TabIndex        =   2
         Top             =   270
         Width           =   1785
      End
      Begin VB.OptionButton O_todas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todas as compras"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   510
         TabIndex        =   1
         Top             =   270
         Width           =   1785
      End
   End
   Begin SSDataWidgets_B.SSDBCombo cboCodigo 
      Bindings        =   "ManContasCliente.frx":4E95A
      DataSource      =   "Data1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   105
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
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   9816
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
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   12648447
   End
   Begin VB.Frame Quadro 
      Caption         =   "Selecionar Caixa"
      Height          =   1470
      Left            =   120
      TabIndex        =   18
      Top             =   1260
      Visible         =   0   'False
      Width           =   4350
      Begin SSDataWidgets_B.SSDBCombo Combo_Caixa 
         Bindings        =   "ManContasCliente.frx":4E96E
         DataSource      =   "Data3"
         Height          =   375
         Left            =   660
         TabIndex        =   7
         Top             =   300
         Width           =   915
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
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   9578
         Columns(0).Caption=   "Descrição"
         Columns(0).Name =   "Descrição"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Descrição"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1614
         Columns(1).Caption=   "Caixa"
         Columns(1).Name =   "Caixa"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Caixa"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         _ExtentX        =   1614
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.OptionButton O_cheque 
         Caption         =   "Caixa - cheque (Oculto)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.OptionButton O_Dinheiro 
         Caption         =   "Caixa - dinheiro (Coluto)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.OptionButton O_Indeterminado 
         Caption         =   "Indeterminado (Oculto)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox Valor_Receber 
         Height          =   375
         Left            =   1605
         TabIndex        =   22
         Top             =   780
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
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
         Format          =   "###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label L_Receber 
         Caption         =   "A Receber"
         Height          =   255
         Left            =   690
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Nome_Caixa 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1605
         TabIndex        =   20
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Caixa"
         Height          =   225
         Left            =   150
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Total a Receber"
      Height          =   255
      Left            =   8340
      TabIndex        =   17
      Top             =   165
      Width           =   1230
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9585
      TabIndex        =   16
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label Nome_Cli 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1845
      TabIndex        =   14
      Top             =   105
      Width           =   5970
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   150
      TabIndex        =   13
      Top             =   165
      Width           =   585
   End
End
Attribute VB_Name = "frmManContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsClientes As Recordset
Private rsConta_Cli As Recordset
Private rsCaixa As Recordset
Private rsCaixas As Recordset
Private rsParametros As Recordset '29/01/2007 - Anderson - Utilizado para obter a quantidade default de parcelas na tela de recebimento
Private rsCliFor As Recordset '29/01/2007 - Anderson - Utilizado para obter a quantidade default de parcelas na tela de recebimento
Private Rec_SQL As Recordset
Private Tot As Single

Private Sub Atualiza_Caixa(ByVal Valor As Double, ByVal Tipo As Integer)

  Dim Sem_Caixa As Integer
  Dim Tot_Dinheiro, Tot_Cheques, Tot_Pré, Tot_Cartões, Tot_Vale, Saldo_Ant, Tot_Vales As Double
  Dim Ordem As Long
 
  'tipo = 1 dinheiro
  'tipo = 2 caixa
 
  'procura para ver se existe caixa no dia
  Sem_Caixa = False
  rsCaixa.Index = "Data"
  rsCaixa.Seek ">", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 0
  If rsCaixa.NoMatch Then Sem_Caixa = True
  If Sem_Caixa = False Then If gnCodFilial <> rsCaixa("Filial") Then Sem_Caixa = True
  If Sem_Caixa = False Then If Data_Atual <> rsCaixa("Data") Then Sem_Caixa = True
  If Sem_Caixa = False Then If Val(Combo_Caixa.Text) <> rsCaixa("Caixa") Then Sem_Caixa = True
  
  If Sem_Caixa = True Then
     'Acha o último dia
     Sem_Caixa = False
     rsCaixa.Seek "<", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 0
     If rsCaixa.NoMatch Then Sem_Caixa = True
     If Sem_Caixa = False Then If rsCaixa("Filial") <> gnCodFilial Then Sem_Caixa = True
     If Sem_Caixa = False Then If Val(Combo_Caixa.Text) <> rsCaixa("Caixa") Then Sem_Caixa = True
     
     If Sem_Caixa = True Then  'Caixa zerado
       rsCaixa.AddNew
         rsCaixa("Filial") = gnCodFilial
         rsCaixa("Data") = Data_Atual
         rsCaixa("Hora") = Format(Time, "hh:mm:ss")
         rsCaixa("Caixa") = Val(Combo_Caixa.Text)
         rsCaixa("Ordem") = 1
         rsCaixa("Descrição") = "Início do dia"
       rsCaixa.Update
     End If
     If Sem_Caixa = False Then  'pega último caixa
       Tot_Dinheiro = rsCaixa("Total Dinheiro")
       Tot_Cheques = rsCaixa("Total Cheques")
       Tot_Pré = rsCaixa("Total Cheques Pré")
       Tot_Cartões = rsCaixa("Total Cartões")
       Tot_Vales = rsCaixa("Total Vales")
       Saldo_Ant = rsCaixa("Final")
     
       rsCaixa.AddNew
         rsCaixa("Filial") = gnCodFilial
         rsCaixa("Data") = Data_Atual
         rsCaixa("Hora") = Format(Time, "hh:mm:ss")
         rsCaixa("Caixa") = Val(Combo_Caixa.Text)
         rsCaixa("Ordem") = 1
         rsCaixa("Descrição") = "Início do dia"
         rsCaixa("Dinheiro") = Tot_Dinheiro
         rsCaixa("Total Dinheiro") = Tot_Dinheiro
         rsCaixa("Cheques") = Tot_Cheques
         rsCaixa("Total Cheques") = Tot_Cheques
         rsCaixa("Cheques Pré") = Tot_Pré
         rsCaixa("Total Cheques Pré") = Tot_Pré
         rsCaixa("Cartões") = Tot_Cartões
         rsCaixa("Total Cartões") = Tot_Cartões
         rsCaixa("Vales") = Tot_Vales
         rsCaixa("Total Vales") = Tot_Vales
         rsCaixa("Saldo Anterior") = Saldo_Ant
         rsCaixa("Final") = Saldo_Ant
       rsCaixa.Update
     End If
  End If
  
 
  'Acha o último caixa
  rsCaixa.Index = "Data"
  rsCaixa.Seek "<", gnCodFilial, Val(Combo_Caixa.Text), Data_Atual, 9999#
  If rsCaixa.NoMatch Then
       DisplayMsg "Caixa não encontrado."
       Exit Sub
  End If
       
  If rsCaixa("Filial") <> gnCodFilial Then Exit Sub
  If rsCaixa("Data") <> Data_Atual Then Exit Sub
  If rsCaixa("Caixa") <> Val(Combo_Caixa.Text) Then Exit Sub
  
  Ordem = rsCaixa("Ordem") + 1
  
  Tot_Dinheiro = rsCaixa("Total Dinheiro")
  Tot_Cheques = rsCaixa("Total Cheques")
  Tot_Pré = rsCaixa("Total Cheques Pré")
  Tot_Cartões = rsCaixa("Total Cartões")
  Tot_Vales = rsCaixa("Total Vales")
  Saldo_Ant = rsCaixa("Final")
     
  rsCaixa.AddNew
    rsCaixa("Filial") = gnCodFilial
    rsCaixa("Data") = Data_Atual
    rsCaixa("Hora") = Format(Time, "hh:mm:ss")
    rsCaixa("Caixa") = Val(Combo_Caixa.Text)
    rsCaixa("Ordem") = Ordem
    rsCaixa("Descrição") = Left(("Conta recebida - " + Nome_cli.Caption), 30)
    
    rsCaixa("Total Dinheiro") = Tot_Dinheiro
    If Tipo = 1 Then
       rsCaixa("Dinheiro") = Valor
       rsCaixa("Total Dinheiro") = Tot_Dinheiro + Valor
    End If
    
    rsCaixa("Total Cheques") = Tot_Cheques
    If Tipo = 2 Then
       rsCaixa("Cheques") = Valor
       rsCaixa("Total Cheques") = Tot_Cheques + Valor
    End If
    
    rsCaixa("Cheques Pré") = 0
    rsCaixa("Total Cheques Pré") = Tot_Pré
    rsCaixa("Cartões") = 0
    rsCaixa("Total Cartões") = Tot_Cartões
    rsCaixa("Vales") = 0
    rsCaixa("Total Vales") = Tot_Vales
    rsCaixa("Saldo Anterior") = Saldo_Ant
    rsCaixa("Final") = Saldo_Ant + rsCaixa("Dinheiro") + rsCaixa("Cheques")
    rsCaixa("Final") = rsCaixa("Final") + rsCaixa("Cheques Pré") + rsCaixa("Cartões") + rsCaixa("Vales")
  rsCaixa.Update

   DisplayMsg "Caixa Atualizado."

End Sub

'07/03/2007 - ANDERSON
'Retirada a função devido a mudanças na forma de pagamento do recebimento
'Private Sub Atualiza_Todas()
' Dim Resp As Integer
' Dim I_Produto As String
' Dim I_Contador As Long
' Dim I_Data As Variant
' Dim I_Filial As Integer
'
'
' Resp = MsgBox("Deseja realmente indicar o recebimento do valor digitado ?  Esta operação não poderá ser anulada posteriormente.", 1, "ATENÇÃO")
' If Resp = 2 Then
'    B_Monta.Enabled = True
'    Recebe_Valor.Enabled = True
'    Recebe_Todas.Enabled = True
'    L_Receber.Visible = False
'    Valor_Receber.Visible = False
'    B_Prossegue.Visible = False
'    Quadro.Visible = False
'    Exit Sub
' End If
'
'
'  rsConta_Cli.Index = "Produto"
'  I_Data = CDate("01/01/80")
'  I_Produto = ""
'  I_Contador = 0
'  I_Filial = 0
'
'  On Error GoTo ErrorHandler
'
'  Call ws.BeginTrans
'
'Lp1:
'  rsConta_Cli.Seek ">", gnCodFilial, Val(cboCodigo.Text), I_Data, I_Produto, I_Contador
'  If rsConta_Cli.NoMatch Then GoTo Fim
'  If rsConta_Cli("Cliente") <> Val(cboCodigo.Text) Then GoTo Fim
'  If rsConta_Cli("Filial") <> gnCodFilial Then GoTo Fim
'
'  'I_Filial = rsConta_Cli("Filial")
'  I_Data = rsConta_Cli("Data")
'  I_Produto = rsConta_Cli("Produto")
'  I_Contador = rsConta_Cli("Contador")
'
'  rsConta_Cli.Edit
'  rsConta_Cli("Valor Pago") = rsConta_Cli("Valor")
'  rsConta_Cli("Data Pagamento") = Data_Atual
'  rsConta_Cli("Data Alteração") = Format(Date, "dd/mm/yyyy")
'  rsConta_Cli.Update
'  GoTo Lp1
'
'
'  'se chegou aqui é porque já está tudo pago
'Fim:
'
'  '31/01/2007 - Anderson - Retirado atualização do caixa, devido a mudança na forma de pagamento da conta do cliente
'  'If O_Dinheiro.Value = True Then
'  '  Call Atualiza_Caixa(CDbl(Total.Caption), 1)
'  'End If
'
'  'If O_cheque.Value = True Then
'  '  Call Atualiza_Caixa(CDbl(Total.Caption), 2)
'  'End If
'
'  Call ws.CommitTrans
'
'  Valor_Receber.Text = ""
'  B_Monta.Enabled = True
'  Recebe_Valor.Visible = False    '
'  Recebe_Todas.Visible = False    '
'  B_Cancela.Visible = False       '
'  L_Receber.Visible = False
'  Valor_Receber.Visible = False
'  B_Prossegue.Visible = False
'  Quadro.Visible = False
'
'  Total.Caption = "0,00"
'  grdContaClientes.RemoveAll
'  grdContaClientes.Visible = True
'
'
'  Exit Sub
'
'ErrorHandler:
'  gsTitle = LoadResString(201)
'  gsMsg = "Erro durante a atualização de contas."
'  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
'  gnStyle = vbOKOnly + vbExclamation
'  On Error Resume Next
'  Call ws.Rollback
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'  Exit Sub
'
'End Sub

Private Sub B_Cancela_Click()
  L_Receber.Visible = False
  Valor_Receber.Visible = False
  Valor_Receber.Text = ""
  
  B_Prossegue.Visible = False
  Quadro.Visible = False
  B_Cancela.Visible = False
    
    
  Recebe_Valor.Enabled = True
  Recebe_Todas.Enabled = True
  B_Monta.Enabled = True
  
'  DisplayMsg "Baixa cancelada."
  
End Sub

Private Sub B_Monta_Click()
 
  Call StatusMsg("")
  
  If Nome_cli.Caption = "" Then
    DisplayMsg "Encontre um cliente antes."
    cboCodigo.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde, montando tabela...")
  DoEvents
  
  Call LoadGridContaClientes
  
  Call StatusMsg("")

End Sub


'09/10/2003 - mpdea
'Corrigido tipo das variáveis Double -> Single
'
Private Sub B_Prossegue_Click()
  Dim Resp As Integer
  Dim A_Receber As Single
  Dim A_Pagar As Single
  Dim I_Data As Variant
  Dim I_Produto As String
  Dim I_Contador As Long
  Dim I_Filial As Integer
  
  '-------------------------------------------------------------------
  '29/01/2007 - Anderson - Parametros para utilização da tela de recebimentos
  Dim typRecebimento As tpPaymentType
  Dim strTipoParcelamento As String
  Dim bytNrContaCC As Byte

  '31/01/2007 - Anderson - Variáveis utilizadas para a implementação da tela de recebimentos
  Dim rstContasReceber As Recordset
  Dim rstCaixa As Recordset
  'Dim rstCliFor As Recordset
  Dim rstCartoes As Recordset
  Dim strSQL As String
  Dim blnInTransaction As Boolean
  
  Dim typTotalizadores As tpPaymentType
  Dim intOrdem As Integer
  'Dim varBookmark As Variant
  Dim dblSaldoAnterior As Double
  'Dim dblValorPago As Double
  Dim bytCaixa As Byte
  'Dim dteDataPgto As Date
  Dim intX As Integer
  'Dim lngContador As Long
  Dim lngCodCliente As Long
  Dim intRet As Integer
 
  Dim intBanco As Integer
  Dim strCheque As String
  Dim strData As String
  Dim dblValor As Double
  Dim intCount As Integer
  Dim intParcelas As Integer
  
  'Dim bytNrContaCC As Byte
  
  Dim intCartaoAdministradora As Integer
  Dim bytCartaoQtdeParcelas As Byte
  Dim dblCartaoVlrParcela As Double
  
  'Dim dblValorPagar As Double
  
  'Dim intRepeatUpdateLocked As Integer
  '-------------------------------------------------------------------
  
  Dim dblValorPago As Double

  Call StatusMsg("")
  
  '-----------------------------------------------------------------------------
  'Validação
  '-----------------------------------------------------------------------------
  '31/01/2007 - Anderson - retirada a validação do caixa, devido a mudança na forma de pagamento da conta do cliente
  'If O_Dinheiro.Value = True Or O_cheque.Value = True Then
    If Nome_Caixa.Caption = "" Then
      DisplayMsg "Digite o caixa aonde o dinheiro / cheque será colocado."
      Exit Sub
    End If
  'End If
  
  'Pega o código do Caixa selecionado
  bytCaixa = Combo_Caixa.Text
  
  If Valor_Receber.Visible = False Then
    Valor_Receber.Text = Total.Caption
    '07/03/2007 - Anderson
    'Retirada essa operação devida as mudanças na forma de pagamento
    'Atualiza_Todas
    'Exit Sub
  End If
  
  If IsNull(Valor_Receber.Text) Then Exit Sub
  If Valor_Receber.Text = "" Then Exit Sub
  If Not IsNumeric(Valor_Receber.Text) Then Exit Sub
  
  If CSng(Valor_Receber.Text) = 0 Then Exit Sub
  
  If CSng(Valor_Receber.Text) > CSng(Format(Tot, "0.00")) Then
    Beep
    DisplayMsg "Valor não pode ser maior que total a receber."
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja realmente indicar o recebimento do valor digitado? Esta operação não poderá ser anulada posteriormente."
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Valor_Receber.Text = ""
    B_Monta.Enabled = True
    Recebe_Valor.Enabled = True
    Recebe_Todas.Enabled = True
    L_Receber.Visible = False
    Valor_Receber.Visible = False
    B_Prossegue.Visible = False
    Quadro.Visible = False
    Exit Sub
  End If
  '-----------------------------------------------------------------------------
  'Recebimento
  '-----------------------------------------------------------------------------
  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", Val(cboCodigo.Text)
  If rsCliFor.NoMatch Then Exit Sub

  dblValorPago = Valor_Receber.Text

  With frmRecebimento
    .Limpa_Tela (0)
    .Só_Leitura.Value = 0
    .Conta.Visible = False
    .L_Sequência.Caption = "-1"
    .Receber.Caption = Format(dblValorPago, FORMAT_VALUE)
    .Intervalo_Parc.Caption = rsParametros.Fields("VR Intervalo Parc").Value
    .Combo_Banco.Text = rsCliFor("Conta Cobrança")
    .Conta.Enabled = rsCliFor.Fields("Tem Conta").Value
    .Max_Cheques.Caption = 0
    .Max_Parcelas.Caption = 0
    If Not rsCliFor.Fields("Faturado").Value Then
      .Max_Cheques.Caption = "1"
      .Max_Parcelas.Caption = "1"
    Else
      .Max_Cheques.Caption = "9999"
      .Max_Parcelas.Caption = "9999"
    End If

    .Show vbModal
    If .Retorno.Caption <> "OK" Then
      Unload frmRecebimento
      Exit Sub
    End If

    .Conta.Visible = True

    typRecebimento.dblDinheiro = Format(CDbl(frmRecebimento.Dinheiro.Text), FORMAT_VALUE)
    typRecebimento.dblCartao = Format(CDbl(frmRecebimento.Cartão.Text), FORMAT_VALUE)
    typRecebimento.dblVale = Format(CDbl(frmRecebimento.Vale.Text), FORMAT_VALUE)
    typRecebimento.dblCheque = Format(frmRecebimento.Pega_Total_Cheques_Separado(False), FORMAT_VALUE)
    typRecebimento.dblChequePre = Format(frmRecebimento.Pega_Total_Cheques_Separado(True), FORMAT_VALUE)
    typRecebimento.dblParcelamento = Format(frmRecebimento.Pega_Total_Parcelas, FORMAT_VALUE)


    If .O_Banco.Value Then
      strTipoParcelamento = "B"
      Call IsDataType(dtByte, .Combo_Banco.Text, bytNrContaCC)
    ElseIf .O_Carteira.Value Then
      strTipoParcelamento = "C"
    ElseIf .O_Carnet.Value Then
      strTipoParcelamento = "T"
    End If

  End With
  
  'LOG *****************
  Dim sSQL_Log As String
  sSQL_Log = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "#, '"
  sSQL_Log = sSQL_Log & Left("Usu:" & gnUserCode & " Fil:" & gnCodFilial & " Cli:" & cboCodigo.Text & " VrTot:" & Total.Caption & " VrPar:" & Valor_Receber.Text, 80) & "', 'RECEBE CONTA CLIENTE')"
  db.Execute sSQL_Log, dbFailOnError
  'fim *******************
  
  ' Processa o recebimento
  '02/02/2007 - Anderson
  'Alteração realizada para aceitar pagamentos através da tela de recebimentos
  'A_Receber = CSng(Valor_Receber.Text)
  A_Receber = typRecebimento.dblDinheiro + typRecebimento.dblCartao + typRecebimento.dblVale + typRecebimento.dblCheque + typRecebimento.dblChequePre + typRecebimento.dblParcelamento
  
  rsConta_Cli.Index = "Produto"
  I_Data = CDate("01/01/80")
  I_Produto = 0
  I_Contador = 0
  I_Filial = gnCodFilial
  
  On Error GoTo ErrorHandler
  
  Call WaitSeconds(1, True) 'Aguarda um segundo para o refresh
  Me.Refresh
  
  Call StatusMsg("Aguarde...")
  Screen.MousePointer = vbHourglass
  
  'Inicia transação
  ws.BeginTrans
  blnInTransaction = True
  
  
  'Abre recordset de Contas a Receber
  strSQL = "SELECT * FROM [Contas a Receber] ORDER BY Contador;"
  Set rstContasReceber = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  
  '-----------------------------------------------------------------------------
  'Caixa
  '-----------------------------------------------------------------------------
  'Verifica o início do caixa, abertura do dia e retorna os últimos valores
    
  If Not gbCheckOpenCaixa(bytCaixa, gnUserCode, dblSaldoAnterior, intOrdem, typTotalizadores) Then
    'Ocorreu erro e a mensagem é exibida pela função
    ws.Rollback
    blnInTransaction = False
    Exit Sub
  End If
  '
  'Atualiza caixa
  Set rstCaixa = db.OpenRecordset("Caixa", dbOpenDynaset)
  With rstCaixa
    intOrdem = intOrdem + 1
    .AddNew
    .Fields("Filial").Value = gnCodFilial
    .Fields("Data").Value = Data_Atual
    .Fields("Caixa").Value = bytCaixa
    .Fields("Funcionário").Value = gnUserCode
    .Fields("Hora").Value = Format(Time, "hh:mm:ss")
    .Fields("Ordem").Value = intOrdem
    .Fields("Descrição").Value = Left("Conta recebida - " & rsCliFor("Nome"), 30)
    
    .Fields("Dinheiro").Value = typRecebimento.dblDinheiro
    .Fields("Total Dinheiro").Value = typTotalizadores.dblDinheiro + .Fields("Dinheiro").Value
    .Fields("Cheques").Value = typRecebimento.dblCheque
    .Fields("Total Cheques").Value = typTotalizadores.dblCheque + .Fields("Cheques").Value
    .Fields("Cheques Pré").Value = typRecebimento.dblChequePre
    .Fields("Total Cheques Pré").Value = typTotalizadores.dblChequePre + .Fields("Cheques Pré").Value
    .Fields("Cartões").Value = typRecebimento.dblCartao
    .Fields("Total Cartões").Value = typTotalizadores.dblCartao + .Fields("Cartões").Value
    .Fields("Vales").Value = typRecebimento.dblVale
    .Fields("Total Vales").Value = typTotalizadores.dblVale + .Fields("Vales").Value
    .Fields("Parcelamento").Value = typRecebimento.dblParcelamento
    .Fields("Total Parcelamento").Value = .Fields("Parcelamento").Value
    .Fields("Saldo Anterior").Value = dblSaldoAnterior
    .Fields("Final").Value = Format(.Fields("Saldo Anterior").Value + _
      .Fields("Dinheiro").Value + .Fields("Cheques").Value + _
      .Fields("Cheques Pré").Value + .Fields("Cartões").Value + _
      .Fields("Vales").Value, FORMAT_VALUE)
    .Update
  End With
  '-----------------------------------------------------------------------------
  
  
  '01, 04 e 07/08/2003 - mpdea
  '-----------------------------------------------------------------------------
  'Atualiza contas a receber
  '-----------------------------------------------------------------------------
  lngCodCliente = rsCliFor("Código")
  With rstContasReceber
        
    '---------------------------------------------------------------------------
    'Cheque
    '---------------------------------------------------------------------------
    ' alteração parametro cheque (Pablo)
    'For intX = 1 To 50
    For intX = 1 To pab_VR_Qtde_Cheques
      intRet = frmRecebimento.Pega_Banco(intX, intBanco, strCheque, strData, dblValor)
      If intRet = 1 Then

        .AddNew
        .Fields("Tipo").Value = "C"
        .Fields("Filial").Value = gnCodFilial
        .Fields("Sequência").Value = 0
        .Fields("Cliente").Value = lngCodCliente
        .Fields("Banco").Value = intBanco
        .Fields("Cheque").Value = strCheque
        .Fields("Vencimento").Value = Format(strData, "dd/mm/yyyy")
        .Fields("Valor").Value = Format(dblValor, FORMAT_VALUE)
        .Fields("Vendedor").Value = 0
        .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
        If CDate(strData) = CDate(Data_Atual) Then
          .Fields("Processado").Value = True
          .Fields("Valor Recebido").Value = .Fields("Valor").Value
          .Fields("Data Recebimento").Value = .Fields("Vencimento").Value
        End If
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rstContasReceber("Cliente") & "- Seq:" & rstContasReceber("Sequência") & "- NF:" & rstContasReceber("Nota") & "- Venc:" & rstContasReceber("Vencimento") & "- Valor:" & rstContasReceber("Valor"), _
          "frmManContas_B_Prossegue_Click (Cheque)", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
        .Update
      End If
    Next intX
    '---------------------------------------------------------------------------
      
      
    '---------------------------------------------------------------------------
    'Parcelamento
    '---------------------------------------------------------------------------
    intCount = 0
    ' alteração parametro parcela (Pablo)
    'For intX = 1 To 50
    For intX = 1 To pab_VR_Qtde_Parcela
      intRet = frmRecebimento.Pega_Parcela(intX, strData, dblValor, intParcelas)
      If intRet = 1 Then
        intCount = intCount + 1
        .AddNew
        .Fields("Tipo").Value = "R"
        .Fields("Filial").Value = gnCodFilial
        .Fields("Cliente").Value = lngCodCliente
        .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Parcela").Value = intCount
        .Fields("Descrição").Value = "Parcela " & intCount & "/" & intParcelas
        .Fields("Vencimento").Value = Format(strData, "dd/mm/yyyy")
        .Fields("Valor").Value = Format(dblValor, FORMAT_VALUE)
        .Fields("Sequência").Value = 0
        .Fields("Nota").Value = 0
        .Fields("Vendedor").Value = 0
        .Fields("Tipo Parcelamento").Value = strTipoParcelamento
        .Fields("Conta Boleto").Value = bytNrContaCC
        .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rstContasReceber("Cliente") & "- Seq:" & rstContasReceber("Sequência") & "- NF:" & rstContasReceber("Nota") & "- Venc:" & rstContasReceber("Vencimento") & "- Valor:" & rstContasReceber("Valor"), _
          "frmManContas_B_Prossegue_Click (Parcelamento)", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
        .Update
      End If
    Next intX
    '---------------------------------------------------------------------------
  
  
    '---------------------------------------------------------------------------
    'Cartão
    '---------------------------------------------------------------------------
    If typRecebimento.dblCartao > 0 Then
      'Administradora
      Call IsDataType(dtInteger, frmRecebimento.Combo_Empresa.Text, intCartaoAdministradora)
      'Quantidade de parcelas
      Call IsDataType(dtByte, frmRecebimento.Label_Cartão2.Caption, bytCartaoQtdeParcelas)
      'Valor da parcela
      Call IsDataType(dtDouble, frmRecebimento.Label_Cartão4.Caption, dblCartaoVlrParcela)
      
      strSQL = "SELECT * FROM Cartões WHERE Código = " & _
        intCartaoAdministradora
      Set rstCartoes = db.OpenRecordset(strSQL, dbOpenSnapshot)
      If Not (rstCartoes.BOF And rstCartoes.EOF) Then
        For intX = 1 To bytCartaoQtdeParcelas
          .AddNew
          .Fields("Tipo").Value = "O"
          .Fields("Filial").Value = gnCodFilial
          .Fields("Sequência").Value = 0
          .Fields("Cliente").Value = lngCodCliente
          .Fields("Administradora").Value = intCartaoAdministradora
          .Fields("Cartão").Value = frmRecebimento.Num_Cartão.Text
          .Fields("Vencimento").Value = (CDate(Data_Atual) + rstCartoes.Fields("Dias Pagar").Value + ((intX - 1) * 30))
          .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
          
          If bytCartaoQtdeParcelas = 1 Then
            .Fields("Valor Cartão").Value = typRecebimento.dblCartao
            .Fields("Valor").Value = Round(CDbl(typRecebimento.dblCartao * ((1 - rstCartoes.Fields("Taxa").Value / 100))), 2)
          Else
            .Fields("Valor Cartão").Value = dblCartaoVlrParcela
            .Fields("Valor").Value = Round(CDbl(dblCartaoVlrParcela * ((1 - rstCartoes.Fields("Taxa").Value / 100))), 2)
          End If
          
          .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
          '10/09/2007 - Anderson
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
            "Cli:" & rstContasReceber("Cliente") & "- Seq:" & rstContasReceber("Sequência") & "- NF:" & rstContasReceber("Nota") & "- Venc:" & rstContasReceber("Vencimento") & "- Valor:" & rstContasReceber("Valor"), _
            "frmManContas_B_Prossegue_Click (Cartões)", _
            "Contas a Receber", g_strArquivoSystemLog
          End If
          .Update
        Next intX
      End If
      rstCartoes.Close
      Set rstCartoes = Nothing
    End If
    '---------------------------------------------------------------------------
    
    
    '---------------------------------------------------------------------------
    'Vendas à vista
    '---------------------------------------------------------------------------
    If typRecebimento.dblDinheiro + typRecebimento.dblVale Then
      If rsParametros.Fields("Gerar Conta Paga").Value Then
        .AddNew
        .Fields("Tipo").Value = "R"
        .Fields("Filial").Value = gnCodFilial
        .Fields("Cliente").Value = lngCodCliente
        .Fields("Sequência").Value = 0
        .Fields("Nota").Value = 0
        .Fields("Vendedor").Value = 0
        .Fields("Descrição").Value = "Pagamento à vista"
        .Fields("Valor").Value = Format(typRecebimento.dblDinheiro + typRecebimento.dblVale, FORMAT_VALUE)
        .Fields("Valor Recebido").Value = .Fields("Valor").Value
        .Fields("Data Recebimento").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Data Emissão").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Vencimento").Value = Format(Data_Atual, "dd/mm/yyyy")
        .Fields("Data Alteração").Value = Format(Data_Atual, "dd/mm/yyyy")
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rstContasReceber("Cliente") & "- Seq:" & rstContasReceber("Sequência") & "- NF:" & rstContasReceber("Nota") & "- Venc:" & rstContasReceber("Vencimento") & "- Valor:" & rstContasReceber("Valor"), _
          "frmManContas_B_Prossegue_Click (Venda a Vista)", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
        .Update
      End If
    End If
    
    .Close
  End With
  
   
Lp1:
  rsConta_Cli.Seek ">", I_Filial, Val(cboCodigo.Text), I_Data, I_Produto, I_Contador
  If rsConta_Cli.NoMatch Then GoTo Fim
  If rsConta_Cli("Cliente") <> Val(cboCodigo.Text) Then GoTo Fim
  
  I_Data = rsConta_Cli("Data")
  I_Produto = rsConta_Cli("Produto")
  I_Contador = rsConta_Cli("Contador")
  I_Filial = rsConta_Cli("Filial")
  If rsConta_Cli("Valor") = rsConta_Cli("Valor Pago") Then GoTo Lp1
  
  A_Pagar = rsConta_Cli("Valor") - rsConta_Cli("Valor Pago")
  
  If A_Receber > A_Pagar Then
    rsConta_Cli.Edit
      rsConta_Cli("Valor Pago") = rsConta_Cli("Valor")
      rsConta_Cli("Data Pagamento") = Data_Atual
      rsConta_Cli("Data Alteração") = Data_Atual
      A_Receber = A_Receber - A_Pagar
    rsConta_Cli.Update
    GoTo Lp1
  End If
  
  rsConta_Cli.Edit
    rsConta_Cli("Valor Pago") = rsConta_Cli("Valor Pago") + A_Receber
    rsConta_Cli("Data Pagamento") = Data_Atual
    rsConta_Cli("Data Alteração") = Data_Atual
  rsConta_Cli.Update
  
  'se chegou aqui é porque já está tudo pago
Fim:

  '31/01/2007 - Anderson - Retirado atualização do caixa, devido a mudança na forma de pagamento da conta do cliente
  'If O_Dinheiro.Value = True Then
  '  Call Atualiza_Caixa(CDbl(Valor_Receber.Text), 1)
  'End If
  
  'If O_cheque.Value = True Then
  '  Call Atualiza_Caixa(CDbl(Valor_Receber.Text), 2)
  'End If

  ws.CommitTrans

  Set rstContasReceber = Nothing
  
  'Descarrega a tela de recebimento
  Unload frmRecebimento
  
  'Atualiza a tela
  B_Cancela_Click
  B_Monta_Click
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault

  Valor_Receber.Text = ""
  B_Monta.Enabled = True
  Recebe_Valor.Enabled = True
  Recebe_Todas.Enabled = True
  L_Receber.Visible = False
  Valor_Receber.Visible = False
  B_Prossegue.Visible = False
  Quadro.Visible = False
  B_Monta_Click
  
  Exit Sub
  
ErrorHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro durante a atualização de contas."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  On Error Resume Next
  Call ws.Rollback
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Private Sub Combo_Caixa_CloseUp()

  Combo_Caixa.Text = Combo_Caixa.Columns(1).Text


  
  Combo_Caixa_LostFocus
  
  
End Sub

Private Sub Combo_Caixa_LostFocus()


  Nome_Caixa.Caption = ""
  If IsNull(Combo_Caixa.Text) Then Exit Sub
  If Combo_Caixa.Text = "" Then Exit Sub
  If Not IsNumeric(Combo_Caixa.Text) Then Exit Sub
  If Val(Combo_Caixa.Text) < 1 Then Exit Sub
  If Val(Combo_Caixa.Text) > 99 Then Exit Sub
 
  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", Val(Combo_Caixa.Text)
  If rsCaixas.NoMatch Then Exit Sub
  Nome_Caixa.Caption = rsCaixas("Descrição") & ""
  
End Sub

Private Sub cboCodigo_CloseUp()
 cboCodigo.Text = cboCodigo.Columns(1).Text
 cboCodigo_LostFocus
End Sub

Private Sub cboCodigo_LostFocus()
 Nome_cli.Caption = ""
 
 If IsNull(cboCodigo.Text) Then Exit Sub
 If Not IsNumeric(cboCodigo.Text) Then Exit Sub
 If Val(cboCodigo.Text) <= 0 Then Exit Sub
 If Val(cboCodigo.Text) > 99999999 Then Exit Sub
 
 rsClientes.Index = "Código"
 rsClientes.Seek "=", Val(cboCodigo.Text)
 If rsClientes.NoMatch Then Exit Sub
 
 Nome_cli.Caption = rsClientes("Nome")
 

End Sub


Private Sub Form_Load()
 
  Call CenterForm(Me)
  
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsConta_Cli = db.OpenRecordset("Conta Cliente")
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  
  '-------------------------------------------------------------------
  '29/01/2007 - Anderson - Abertura dos parametros da empresa fiial para obter o valor default das parcelas
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial")

  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then
    MsgBox "Filial não encontrada", vbCritical, "Erro"
    Exit Sub
  End If
  '-------------------------------------------------------------------
  
  Data1.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
   
  If gbCaixas = False Then
    Combo_Caixa.Text = 1
    Combo_Caixa_LostFocus
    Combo_Caixa.Enabled = False
  End If
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

 
 rsClientes.Close
 rsConta_Cli.Close
 rsCaixa.Close
 rsCaixas.Close
 rsParametros.Close  '29/01/2007 - Anderson - Parametros da empresa fiial para obter o valor default das parcelas
 rsCliFor.Close '29/01/2007 - Anderson - Parametros da empresa fiial para obter o valor default das parcelas
 
 Set rsParametros = Nothing '29/01/2007 - Anderson - Parametros da empresa fiial para obter o valor default das parcelas
 Set rsCliFor = Nothing '29/01/2007 - Anderson - Parametros da empresa fiial para obter o valor default das parcelas
 Set rsClientes = Nothing
 Set rsConta_Cli = Nothing
 Set rsCaixa = Nothing
 Set rsCaixas = Nothing
  
End Sub

Private Sub grdContaClientes_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  End If
End Sub

'31/01/2007 - Anderson - retirada a funcionalidade desse recurso devido a mudança na forma de pagamento da conta do cliente
'Private Sub O_Indeterminado_Click()
'  If Not frmGerente.gbSenhaGerente Then
'    O_Dinheiro.Value = True
'    Exit Sub
'  End If
'End Sub

Private Sub Recebe_Todas_Click()
  Call StatusMsg("")
  Quadro.Visible = True
  B_Prossegue.Visible = True
  B_Cancela.Visible = True
  Recebe_Valor.Enabled = False
  Recebe_Todas.Enabled = False
  B_Monta.Enabled = False

End Sub

Private Sub Recebe_Valor_Click()
  Call StatusMsg("")
  

  L_Receber.Visible = True
  Valor_Receber.Visible = True
  B_Prossegue.Visible = True
  Quadro.Visible = True
  B_Cancela.Visible = True
    
  Valor_Receber.SetFocus
  
  Recebe_Valor.Enabled = False
  Recebe_Todas.Enabled = False
  B_Monta.Enabled = False
End Sub

Private Sub Valor_Receber_KeyPress(KeyAscii As Integer)
 KeyAscii = gnGotCurrency(KeyAscii)
End Sub

Private Sub LoadGridContaClientes()
  Dim rsContas_Cli As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sSql As String
  Dim sCod As String
  
  On Error GoTo ErrContatos
  
  bAllow = grdContaClientes.AllowAddNew
  grdContaClientes.AllowAddNew = True
  grdContaClientes.AllowUpdate = True
  
  If Len(Trim(cboCodigo.Text)) > 0 Then
    sCod = cboCodigo.Text
  Else
    sCod = "-1"
  End If
  
  sSql = "SELECT * FROM [Conta Cliente] WHERE Filial = " & gnCodFilial & " And Cliente = " & sCod
  If O_Pagas.Value = True Then
    sSql = sSql + " AND Valor = [Valor Pago]"
  Else
    If O_Não_Pagas.Value = True Then
      sSql = sSql + " AND Valor <> [Valor Pago]"
    End If
  End If
    
  sSql = sSql & " ORDER BY Data, Produto"
  Set rsContas_Cli = db.OpenRecordset(sSql, dbOpenDynaset)

  grdContaClientes.RemoveAll
  grdContaClientes.Redraw = False
  
  If Not rsContas_Cli.EOF Then
    With rsContas_Cli
      .MoveFirst
      Tot = 0
      Do While Not .EOF
        sRecord = .Fields("Data") & vbTab & _
          .Fields("Produto") & vbTab & _
          .Fields("Descrição") & vbTab & _
          .Fields("Qtde") & vbTab & _
          .Fields("Valor") & vbTab & _
          .Fields("Valor Pago") & vbTab & _
          .Fields("Data Pagamento") & vbTab & _
          .Fields("TabPrecos")
        grdContaClientes.AddItem sRecord
        Tot = Tot + (.Fields("Valor") - .Fields("Valor Pago"))
        .MoveNext
      Loop
      .MoveFirst
      Total.Caption = Format(Tot, "###,###,##0.00")
      If Tot <> 0 Then
        If O_Pagas.Value = False Then
          Recebe_Todas.Visible = True
          Recebe_Valor.Visible = True
          '22/07/2004 - Daniel
          'Este bug foi encontrado pela BIC Amazônia de Manaus
          'Ao dar baixa em uma conta em seguida carregar as
          'contas do próximo cliente os botões ficavam visíveis
          'mas desabilitados
          Recebe_Todas.Enabled = True
          Recebe_Valor.Enabled = True
        Else
          Recebe_Todas.Visible = False
          Recebe_Valor.Visible = False
        End If
      Else
        Recebe_Todas.Visible = False
        Recebe_Valor.Visible = False
      End If
    End With
    grdContaClientes.Scroll -99, -99
  Else
    DisplayMsg "Nenhuma conta encontrada segundo os critérios fornecidos."
    '07/03/2007 - Anderson
    'Limpa tela
    Total.Caption = ""
    Recebe_Todas.Visible = False
    Recebe_Valor.Visible = False
  End If

  grdContaClientes.Redraw = True
  grdContaClientes.AllowAddNew = bAllow
  grdContaClientes.AllowUpdate = bAllow

  rsContas_Cli.Close
  Set rsContas_Cli = Nothing
  Exit Sub

ErrContatos:
  Exit Sub

End Sub

'Private Sub WriteGridContaClientes()
'  Dim sSql As String
'  Dim bm As Variant
'  Dim nRow As Long
'  Dim rsContas_Cli As Recordset
'
'  On Error GoTo ErrHandler
'
'  grdContaClientes.Update
'
'  Call ws.BeginTrans
'
'  Set rsContas_Cli = db.OpenRecordset("SELECT * FROM [Contatos Efetuados] WHERE Cliente = " & cboCodigo.Text, dbOpenDynaset)
'
'  With rsContas_Cli
'
'    If Not .EOF Then
'      Do While Not .EOF
'        .Delete
'        .MoveNext
'      Loop
'    End If
'
'    For nRow = 0 To grdContaClientes.Rows - 1
'      bm = grdContaClientes.AddItemBookmark(nRow)
'      If IsDate(grdContaClientes.Columns("Data").CellText(bm)) Then
'        .AddNew
'        .Fields("Cliente") = cboCodigo.Text
'        .Fields("Data") = grdContaClientes.Columns("Data").CellText(bm)
'        .Fields("Descrição") = grdContaClientes.Columns("Descricao").CellText(bm)
'        .Fields("Pendência") = Not (grdContaClientes.Columns("Pendencia").CellText(bm) = "")
'        If IsDate(grdContaClientes.Columns("DataAviso").CellText(bm)) Then
'          .Fields("Data Aviso") = grdContaClientes.Columns("DataAviso").CellText(bm)
'        End If
'        .Update
'      End If
'    Next nRow
'
'  End With
'
'  rsContas_Cli.Close
'  Set rsContas_Cli = Nothing
'
'  Call ws.CommitTrans
'  Exit Sub
'
'ErrHandler:
'  gsTitle = LoadResString(201)
'  gsMsg = "Erro ao Atualizar Contas Clientes."
'  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
'  gnStyle = vbOKOnly + vbExclamation
'  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'  Exit Sub
'
'End Sub

