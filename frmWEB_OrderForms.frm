VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmWEB_OrderForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Pedidos da Loja Virtual"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frmWEB_OrderForms.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9615
   Begin VB.TextBox txtSequencia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtPasso 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   645
      Left            =   120
      TabIndex        =   61
      Top             =   690
      Width           =   9375
      Begin VB.TextBox txtShopperCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtShopperName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.TextBox txtBoleto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin TabDlg.SSTab sstOrderForm 
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Acompanhamento Logístico"
      TabPicture(0)   =   "frmWEB_OrderForms.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sspWarning"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNextStep"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelOrderForm"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Dados do &Pedido"
      TabPicture(1)   =   "frmWEB_OrderForms.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTitle(12)"
      Tab(1).Control(1)=   "lblTitle(10)"
      Tab(1).Control(2)=   "lblTitle(9)"
      Tab(1).Control(3)=   "lblTitle(8)"
      Tab(1).Control(4)=   "lblTitle(7)"
      Tab(1).Control(5)=   "lblTitle(6)"
      Tab(1).Control(6)=   "lblTitle(4)"
      Tab(1).Control(7)=   "lblTitle(3)"
      Tab(1).Control(8)=   "lblTitle(1)"
      Tab(1).Control(9)=   "lblChageStatus(0)"
      Tab(1).Control(10)=   "lblChageStatus(1)"
      Tab(1).Control(11)=   "lblTitle(15)"
      Tab(1).Control(12)=   "lblTitle(14)"
      Tab(1).Control(13)=   "lblTitle(11)"
      Tab(1).Control(14)=   "lblTitle(16)"
      Tab(1).Control(15)=   "txtOrderID"
      Tab(1).Control(16)=   "txtFormaPagamento"
      Tab(1).Control(17)=   "txtSubTotal"
      Tab(1).Control(18)=   "txtBonusUtilizado"
      Tab(1).Control(19)=   "txtBonusTotal"
      Tab(1).Control(20)=   "txtShippingMethod"
      Tab(1).Control(21)=   "txtStatusShopper"
      Tab(1).Control(22)=   "txtStatusAdmin"
      Tab(1).Control(23)=   "txtTotal"
      Tab(1).Control(24)=   "txtComentario"
      Tab(1).Control(25)=   "txtTraceCode"
      Tab(1).Control(26)=   "txtShippingTotal"
      Tab(1).Control(27)=   "txtSeguro"
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Dados para &Entrega"
      TabPicture(2)   =   "frmWEB_OrderForms.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtShipField(10)"
      Tab(2).Control(1)=   "txtShipField(9)"
      Tab(2).Control(2)=   "txtShipField(8)"
      Tab(2).Control(3)=   "txtShipField(7)"
      Tab(2).Control(4)=   "txtShipField(6)"
      Tab(2).Control(5)=   "txtShipField(5)"
      Tab(2).Control(6)=   "txtShipField(4)"
      Tab(2).Control(7)=   "txtShipField(3)"
      Tab(2).Control(8)=   "txtShipField(2)"
      Tab(2).Control(9)=   "txtShipField(1)"
      Tab(2).Control(10)=   "txtShipField(0)"
      Tab(2).Control(11)=   "lblShip(10)"
      Tab(2).Control(12)=   "lblShip(9)"
      Tab(2).Control(13)=   "lblShip(8)"
      Tab(2).Control(14)=   "lblShip(7)"
      Tab(2).Control(15)=   "lblShip(6)"
      Tab(2).Control(16)=   "lblShip(5)"
      Tab(2).Control(17)=   "lblShip(4)"
      Tab(2).Control(18)=   "lblShip(3)"
      Tab(2).Control(19)=   "lblShip(2)"
      Tab(2).Control(20)=   "lblShip(1)"
      Tab(2).Control(21)=   "lblShip(0)"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Dados para &Cobrança"
      TabPicture(3)   =   "frmWEB_OrderForms.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtBillField(10)"
      Tab(3).Control(1)=   "txtBillField(9)"
      Tab(3).Control(2)=   "txtBillField(8)"
      Tab(3).Control(3)=   "txtBillField(7)"
      Tab(3).Control(4)=   "txtBillField(6)"
      Tab(3).Control(5)=   "txtBillField(5)"
      Tab(3).Control(6)=   "txtBillField(4)"
      Tab(3).Control(7)=   "txtBillField(3)"
      Tab(3).Control(8)=   "txtBillField(2)"
      Tab(3).Control(9)=   "txtBillField(1)"
      Tab(3).Control(10)=   "txtBillField(0)"
      Tab(3).Control(11)=   "lblBill(10)"
      Tab(3).Control(12)=   "lblBill(9)"
      Tab(3).Control(13)=   "lblBill(8)"
      Tab(3).Control(14)=   "lblBill(7)"
      Tab(3).Control(15)=   "lblBill(6)"
      Tab(3).Control(16)=   "lblBill(5)"
      Tab(3).Control(17)=   "lblBill(4)"
      Tab(3).Control(18)=   "lblBill(3)"
      Tab(3).Control(19)=   "lblBill(2)"
      Tab(3).Control(20)=   "lblBill(1)"
      Tab(3).Control(21)=   "lblBill(0)"
      Tab(3).ControlCount=   22
      TabCaption(4)   =   "&Itens"
      TabPicture(4)   =   "frmWEB_OrderForms.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdItens"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Histórico do Status"
      TabPicture(5)   =   "frmWEB_OrderForms.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdHistoric"
      Tab(5).ControlCount=   1
      Begin VB.TextBox txtSeguro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2520
         Width           =   1400
      End
      Begin VB.TextBox txtShippingTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   1400
      End
      Begin VB.TextBox txtTraceCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   2900
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   -70200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   -70440
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   -67200
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   -70440
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   -67200
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelOrderForm 
         BackColor       =   &H008080FF&
         Caption         =   "Cancelar Pedido"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   4335
      End
      Begin VB.CommandButton cmdNextStep 
         BackColor       =   &H0080C0FF&
         Caption         =   "Confirmar próximo passo >>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   240
         TabIndex        =   73
         Top             =   480
         Width           =   4335
         Begin VB.Image imgCheckPicture 
            Height          =   315
            Left            =   3960
            Picture         =   "frmWEB_OrderForms.frx":0632
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgCheck 
            Height          =   315
            Index           =   0
            Left            =   240
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lblPassoName 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Recebido"
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   77
            Top             =   360
            Width           =   1230
         End
         Begin VB.Image imgCheck 
            Height          =   315
            Index           =   1
            Left            =   240
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblPassoName 
            AutoSize        =   -1  'True
            Caption         =   "Pagamento Confirmado"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   76
            Top             =   840
            Width           =   1650
         End
         Begin VB.Image imgCheck 
            Height          =   315
            Index           =   2
            Left            =   240
            Top             =   1200
            Width           =   285
         End
         Begin VB.Label lblPassoName 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Embalado (Recibo, Etiqueta)"
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   75
            Top             =   1320
            Width           =   2565
         End
         Begin VB.Image imgCheck 
            Height          =   315
            Index           =   3
            Left            =   240
            Top             =   1680
            Width           =   285
         End
         Begin VB.Label lblPassoName 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Enviado"
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   74
            Top             =   1800
            Width           =   1125
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2520
         Width           =   1400
      End
      Begin VB.TextBox txtStatusAdmin 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   -70200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txtStatusShopper 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txtShippingMethod 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   4405
      End
      Begin VB.TextBox txtBonusTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtBonusUtilizado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -68640
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2520
         Width           =   1400
      End
      Begin VB.TextBox txtFormaPagamento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtOrderID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   4405
      End
      Begin SSDataWidgets_B.SSDBGrid grdItens 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   8895
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   6
         AllowUpdate     =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         BackColorOdd    =   14737632
         RowHeight       =   423
         ExtraHeight     =   26
         Columns.Count   =   6
         Columns(0).Width=   2699
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4101
         Columns(1).Caption=   "Descrição"
         Columns(1).Name =   "Descricao"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).VertScrollBar=   -1  'True
         Columns(2).Width=   979
         Columns(2).Caption=   "Qtde."
         Columns(2).Name =   "Quantidade"
         Columns(2).Alignment=   1
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         Columns(3).Width=   2434
         Columns(3).Caption=   "Preço"
         Columns(3).Name =   "Preco"
         Columns(3).Alignment=   1
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   6
         Columns(3).NumberFormat=   "CURRENCY"
         Columns(3).FieldLen=   256
         Columns(4).Width=   2143
         Columns(4).Caption=   "Desconto"
         Columns(4).Name =   "Desconto"
         Columns(4).Alignment=   1
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   6
         Columns(4).NumberFormat=   "CURRENCY"
         Columns(4).FieldLen=   256
         Columns(5).Width=   2355
         Columns(5).Caption=   "Preço Final"
         Columns(5).Name =   "Total"
         Columns(5).Alignment=   1
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   6
         Columns(5).NumberFormat=   "CURRENCY"
         Columns(5).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   15690
         _ExtentY        =   5530
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
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   -68520
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   -70440
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2520
         Width           =   4575
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2520
         Width           =   4095
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1320
         Width           =   7455
      End
      Begin VB.TextBox txtBillField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   -68520
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   -70440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2520
         Width           =   4575
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2520
         Width           =   4095
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1320
         Width           =   7455
      End
      Begin VB.TextBox txtShipField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   8895
      End
      Begin SSDataWidgets_B.SSDBGrid grdHistoric 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   8895
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   4
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         BackColorOdd    =   14737632
         RowHeight       =   423
         ExtraHeight     =   212
         Columns.Count   =   4
         Columns(0).Width=   2302
         Columns(0).Caption=   "Data"
         Columns(0).Name =   "Data"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   7
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3519
         Columns(1).Caption=   "Passo (Status)"
         Columns(1).Name =   "Passo"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   4419
         Columns(2).Caption=   "Status visto pelo comprador"
         Columns(2).Name =   "StatusShopper"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).VertScrollBar=   -1  'True
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   4419
         Columns(3).Caption=   "Status visto pelo administrador"
         Columns(3).Name =   "StatusAdmin"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).VertScrollBar=   -1  'True
         Columns(3).Locked=   -1  'True
         TabNavigation   =   1
         _ExtentX        =   15690
         _ExtentY        =   5530
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
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   240
         TabIndex        =   80
         Top             =   2760
         Width           =   4335
         Begin VB.Image imgCancelPicture 
            Height          =   285
            Left            =   3960
            Picture         =   "frmWEB_OrderForms.frx":0A9A
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgCheck 
            Height          =   315
            Index           =   4
            Left            =   240
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lblPassoName 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Cancelado"
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   81
            Top             =   360
            Width           =   1305
         End
      End
      Begin Threed.SSPanel sspWarning 
         Height          =   1455
         Left            =   4800
         TabIndex        =   83
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   2566
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   6
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   120
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ATENÇÃO"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   $"frmWEB_OrderForms.frx":0EAE
            Height          =   735
            Left            =   120
            TabIndex        =   85
            Top             =   600
            Width           =   4095
         End
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Seguro"
         Height          =   195
         Index           =   16
         Left            =   -74760
         TabIndex        =   97
         Top             =   2280
         Width           =   510
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Valor da entrega"
         Height          =   195
         Index           =   11
         Left            =   -74760
         TabIndex        =   96
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Código de rastreamento"
         Height          =   195
         Index           =   14
         Left            =   -73200
         TabIndex        =   95
         Top             =   1680
         Width           =   1680
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Comentário"
         Height          =   195
         Index           =   15
         Left            =   -70200
         TabIndex        =   94
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "DDD"
         Height          =   195
         Index           =   10
         Left            =   -69600
         TabIndex        =   93
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Index           =   9
         Left            =   -70440
         TabIndex        =   92
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   91
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Index           =   7
         Left            =   -67200
         TabIndex        =   90
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "DDD"
         Height          =   195
         Index           =   10
         Left            =   -69600
         TabIndex        =   89
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Index           =   9
         Left            =   -70440
         TabIndex        =   88
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   87
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Index           =   7
         Left            =   -67200
         TabIndex        =   86
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblChageStatus 
         AutoSize        =   -1  'True
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -66360
         MouseIcon       =   "frmWEB_OrderForms.frx":0F4D
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label lblChageStatus 
         AutoSize        =   -1  'True
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -70920
         MouseIcon       =   "frmWEB_OrderForms.frx":1817
         MousePointer    =   99  'Custom
         TabIndex        =   78
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Index           =   1
         Left            =   -71760
         TabIndex        =   72
         Top             =   2280
         Width           =   360
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Status visível ao comprador"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   71
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Status visível ao administrador"
         Height          =   195
         Index           =   4
         Left            =   -70200
         TabIndex        =   70
         Top             =   2880
         Width           =   2160
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pagamento"
         Height          =   195
         Index           =   6
         Left            =   -70200
         TabIndex        =   69
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Bônus gerado"
         Height          =   195
         Index           =   7
         Left            =   -70200
         TabIndex        =   68
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Bônus utilizado"
         Height          =   195
         Index           =   8
         Left            =   -68640
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "SubTotal"
         Height          =   195
         Index           =   9
         Left            =   -73200
         TabIndex        =   66
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Método de entrega"
         Height          =   195
         Index           =   10
         Left            =   -74760
         TabIndex        =   65
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Código chave identificador"
         Height          =   195
         Index           =   12
         Left            =   -74760
         TabIndex        =   64
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Index           =   6
         Left            =   -68520
         TabIndex        =   60
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Index           =   5
         Left            =   -72720
         TabIndex        =   59
         Top             =   2880
         Width           =   330
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   58
         Top             =   2880
         Width           =   315
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Index           =   3
         Left            =   -70440
         TabIndex        =   57
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   56
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   55
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   54
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Index           =   6
         Left            =   -68520
         TabIndex        =   53
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Index           =   5
         Left            =   -72720
         TabIndex        =   52
         Top             =   2880
         Width           =   330
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   51
         Top             =   2880
         Width           =   315
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Index           =   3
         Left            =   -70440
         TabIndex        =   50
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   49
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   48
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblShip 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   420
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Sequência"
      Height          =   195
      Index           =   13
      Left            =   5640
      TabIndex        =   82
      Top             =   120
      Width           =   765
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Index           =   5
      Left            =   7320
      TabIndex        =   63
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   62
      Top             =   120
      Width           =   450
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   9000
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmWEB_OrderForms.frx":20E1
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmWEB_OrderForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'06/11/2002 - mpdea
'Comentado referências a Bônus da pasta Dados do Pedido
'(aguardando definições para implementação futura)

Option Explicit

Private mrsWEBOrderForms As Recordset
Private mstrToolCaption As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub PosRecordset(ByVal lngID As Long)
  With mrsWEBOrderForms
    .FindFirst "ID = " & lngID
    If .NoMatch Then
      MsgBox "Registro não localizado.", vbExclamation, "Atenção"
    Else
      Call ClearScreen
      Call ShowRecord
    End If
  End With
End Sub

Private Sub ClearScreen()
  Dim objControl As control
  Dim intX As Integer
  
  For Each objControl In Me.Controls
    If TypeOf objControl Is TextBox Then
      objControl.Text = ""
    End If
  Next objControl
  
  For intX = imgCheck.LBound To imgCheck.UBound
    imgCheck(intX).Picture = LoadPicture()
  Next intX
  
  cmdNextStep.Enabled = False
  cmdCancelOrderForm.Enabled = False
  
  lblChageStatus(0).Enabled = False
  lblChageStatus(1).Enabled = False
  
  With grdItens
    .Redraw = False
    .RemoveAll
    .Redraw = True
  End With
  
  With grdHistoric
    .Redraw = False
    .RemoveAll
    .RowHeight = 850
    .Redraw = True
  End With
  
End Sub

Private Sub ShowRecord()
  Dim rsItens As Recordset
  Dim rsHistoric As Recordset
  Dim strSQL As String
  Dim strCodProd As String
  Dim intErro As Integer
  Dim lngCodigo As Long
  Dim strNome As String
  Dim enuStep As enWEB_OrderFormStep
  
  Dim strDescrProduto As String
  
  
  On Error GoTo ErrHandler
  
  With mrsWEBOrderForms
  
    txtBoleto.Text = .Fields("Boleto").Value
    txtSequencia.Text = .Fields("Sequencia").Value & ""
    
    Call WEB_GetShopperData(.Fields("ShopperID").Value, lngCodigo, strNome)
    txtShopperCodigo.Text = lngCodigo
    txtShopperName.Text = strNome
    
    'Passo (Status)
    enuStep = .Fields("Passo").Value
    txtPasso.Text = gstrWEB_GetDescPasso(enuStep)
    
    'Pedido cancelado
    If enuStep = ofsCanceled Then
      imgCheck(4).Picture = imgCancelPicture.Picture
    Else
      If enuStep >= ofsReceived Then
        imgCheck(0).Picture = imgCheckPicture.Picture
      End If
      If enuStep >= ofsConfirmedPayment Then
        imgCheck(1).Picture = imgCheckPicture.Picture
      End If
      If enuStep >= ofsPacked Then
        imgCheck(2).Picture = imgCheckPicture.Picture
      End If
      If enuStep >= ofsHasSent Then
        imgCheck(3).Picture = imgCheckPicture.Picture
      End If
    End If
    
    sspWarning.Visible = (enuStep = ofsReceived) And _
                         (CLng("0" & .Fields("Sequencia").Value) = 0)
    
    'Realiza verificação referente ao passo atual
    'Não pode haver alterações após o Pedido ser Enviado
    cmdNextStep.Enabled = enuStep < ofsHasSent
    cmdCancelOrderForm.Enabled = enuStep < ofsHasSent
    
    lblChageStatus(0).Enabled = True
    lblChageStatus(1).Enabled = True
    
    txtOrderID.Text = .Fields("OrderID").Value
    txtData.Text = .Fields("Data").Value
    txtShippingMethod.Text = mstrGetDescShipping(.Fields("ShippingMethod").Value)
    txtTraceCode.Text = .Fields("TraceCode").Value & ""
    txtFormaPagamento.Text = mstrGetDescPayment(.Fields("CodPagamento").Value)
    txtBonusTotal.Text = .Fields("BonusTotal").Value
    txtBonusUtilizado.Text = .Fields("BonusUtilizado").Value
    txtShippingTotal.Text = Format(.Fields("ShippingTotal").Value, "Currency")
    '12/04/2005 - Daniel
    'Adicionado o campo Seguro
    txtSeguro.Text = Format(IIf(IsNumeric(.Fields("Seguro").Value), (.Fields("Seguro").Value), 0), "Currency")
    '-------------------------
    txtSubTotal.Text = Format(.Fields("SubTotal").Value, "Currency")
    txtTotal.Text = Format(.Fields("Total").Value, "Currency")
    txtStatusAdmin.Text = .Fields("StatusAdmin").Value & ""
    txtStatusShopper.Text = .Fields("StatusShopper").Value & ""
    
    '21/05/2004 - mpdea
    'Novos campos
    txtComentario.Text = .Fields("Comentario").Value & ""
    
    
    'Ship
    txtShipField(0).Text = .Fields("ShipName").Value & ""
    txtShipField(1).Text = .Fields("ShipAddress").Value & ""
    txtShipField(2).Text = .Fields("ShipCity").Value & ""
    txtShipField(3).Text = .Fields("ShipState").Value & ""
    txtShipField(4).Text = .Fields("ShipZip").Value & ""
    txtShipField(5).Text = .Fields("ShipCountry").Value & ""
    txtShipField(6).Text = .Fields("ShipPhone").Value & ""
    
    '21/05/2004 - mpdea
    'Novos campos
    txtShipField(7).Text = .Fields("ShipStreetNumber").Value & ""
    txtShipField(8).Text = .Fields("ShipStreetCompl").Value & ""
    txtShipField(9).Text = .Fields("ShipDistrict").Value & ""
    txtShipField(10).Text = .Fields("ShipDDDPhone").Value & ""
    
    
    'Bill
    txtBillField(0).Text = .Fields("BillName").Value & ""
    txtBillField(1).Text = .Fields("BillAddress").Value & ""
    txtBillField(2).Text = .Fields("BillCity").Value & ""
    txtBillField(3).Text = .Fields("BillState").Value & ""
    txtBillField(4).Text = .Fields("BillZip").Value & ""
    txtBillField(5).Text = .Fields("BillCountry").Value & ""
    txtBillField(6).Text = .Fields("BillPhone").Value & ""
    
    '21/05/2004 - mpdea
    'Novos campos
    txtBillField(7).Text = .Fields("BillStreetNumber").Value & ""
    txtBillField(8).Text = .Fields("BillStreetCompl").Value & ""
    txtBillField(9).Text = .Fields("BillDistrict").Value & ""
    txtBillField(10).Text = .Fields("BillDDDPhone").Value & ""
    
    
    strSQL = "SELECT * FROM WEB_OrderItens WHERE OrderFormID = " & .Fields("ID").Value
    Set rsItens = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rsItens
      If Not .BOF And Not .EOF Then
        Do Until .EOF
          
          Call Acha_Produto(.Fields("sku").Value, strCodProd, 0, 0, 0, 0, intErro)
          
          '23/06/2004 - mpdea
          'Modificado para exibir a descrição de produto não localizado
          If intErro = 0 Then
            strDescrProduto = gsGetNameProduto(strCodProd)
          Else
            strDescrProduto = "Produto Inexistente / Apagado"
          End If
          
          'Adiciona o produto
          grdItens.AddItem .Fields("sku").Value & vbTab & _
                           strDescrProduto & vbTab & _
                           .Fields("Quantity").Value & vbTab & _
                           .Fields("ListPrice").Value & vbTab & _
                           .Fields("Discount").Value & vbTab & _
                           .Fields("Total").Value
          
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsItens = Nothing
    
    strSQL = "SELECT * FROM WEB_OrderStatusHistoric WHERE OrderFormID = " & _
             .Fields("ID").Value & " ORDER BY Data DESC"
    Set rsHistoric = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rsHistoric
      If Not .BOF And Not .EOF Then
        Do Until .EOF
          grdHistoric.AddItem .Fields("Data").Value & vbTab & _
                              gstrWEB_GetDescPasso(.Fields("Passo").Value) & vbTab & _
                              .Fields("StatusShopper").Value & vbTab & _
                              .Fields("StatusAdmin").Value
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsHistoric = Nothing
  
  End With
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub GoToLinkManager()
  'Endereço do gerenciador da Loja Virtual
  Const LINK_MANAGER As String = "https://www.ciashop.com.br/REPLACEME/manager"
  'Texto a ser substituído
  Const REPLACEME As String = "REPLACEME"
  Dim rsGet As Recordset
  Dim strStore As String
  Dim strLink As String
  Dim lngRet As Long
  Dim strSQL As String
  
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  
  strSQL = "SELECT CNX_Store FROM WEB_Config WHERE ID = 1"
  Set rsGet = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rsGet
    If Not .BOF And Not .EOF Then
      strStore = .Fields("CNX_Store").Value & ""
      If strStore <> "" Then
'        If MsgBox("A ação a seguir requer conexão com a internet. Deseja continuar?", vbQuestion + vbOKCancel, "Atenção") = vbOK Then
          strLink = Replace(LINK_MANAGER, REPLACEME, strStore)
          lngRet = ShellExecute(hwnd, "open", strLink, "", "", vbNormalFocus)
'        End If
      Else
        MsgBox "Aplicativo Quick Web não configurado.", vbExclamation, "Aviso"
      End If
    Else
      MsgBox "Aplicativo Quick Web não configurado.", vbExclamation, "Aviso"
    End If
    .Close
  End With
  
  Screen.MousePointer = vbDefault
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  Select Case Tool.Name
    Case "miOpFirst"
      Call Navigate(navFirst)
      
    Case "miOpPrevious"
      Call Navigate(navPrevious)
      
    Case "miOpNext"
      Call Navigate(navNext)
      
    Case "miOpLast"
      Call Navigate(navLast)
      
    Case "miRefresh"
      Screen.MousePointer = vbHourglass
      Call ClearScreen
      mrsWEBOrderForms.Requery
      If mrsWEBOrderForms.RecordCount > 0 Then
        Call ShowRecord
      End If
      Screen.MousePointer = vbDefault
      
    Case "miOpFind"
      frmWEB_OFFind.Show ', Me
    
    Case "miManagerStore"
      Call GoToLinkManager
    
  End Select
  
End Sub

Private Sub Navigate(ByVal lngType As enNavigate)
    
  On Error GoTo ErrHandler
    
  Call ClearScreen
  
  With mrsWEBOrderForms
    If .RecordCount > 0 Then
      Select Case lngType
        Case navFirst
          .MoveFirst
        Case navNext
          If .EOF Then
            .MoveLast
          Else
            .MoveNext
            If .EOF Then .MoveLast
          End If
        Case navPrevious
          If .BOF Then
            .MoveFirst
          Else
            .MovePrevious
            If .BOF Then .MoveFirst
          End If
        Case navLast
          .MoveLast
      End Select
      Call ShowRecord
    End If
  End With
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub ActiveBar1_ComboDrop(ByVal Tool As ActiveBarLibraryCtl.Tool)
  mstrToolCaption = Tool.Text
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Dim strFilter As String
  Dim strSQL As String
  
  If Tool.Name = "miFilter" Then
    Select Case Tool.CBListIndex
      Case 0 'Todos
        strFilter = "Passo >= " & ofsReceived
      Case 1 'Novo Pedidos
        strFilter = "Passo = " & ofsReceived
      Case 2 'Pagamento confirmado
        strFilter = "Passo = " & ofsConfirmedPayment
      Case 3 'Produto embalado
        strFilter = "Passo = " & ofsPacked
      Case 4 'Pedido enviado
        strFilter = "Passo = " & ofsHasSent
      Case 5 'Pedido cancelado
        strFilter = "Passo = " & ofsCanceled
      Case 6 'Personalizar
        If frmWEB_OFPersonalizeFilter.Personalize(strFilter) Then
          Tool.Text = "Personalizado"
        Else
          Tool.Text = mstrToolCaption
          Exit Sub
        End If
    End Select
    
    strSQL = "SELECT * FROM WEB_OrderForms WHERE Filial = " & _
             gnCodFilial & " AND " & strFilter & " ORDER BY ID DESC"
    
    Set mrsWEBOrderForms = db.OpenRecordset(strSQL, dbOpenDynaset)
    Call Navigate(navFirst)
  End If
  
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  Cancel = True
End Sub

'05/03/2003 - mpdea
'Corrigido invalid use of null ao passar a sequência
Private Sub cmdCancelOrderForm_Click()
  Dim enuStep As enWEB_OrderFormStep
  Dim strText As String
  
  With mrsWEBOrderForms
    enuStep = .Fields("Passo").Value
    
    If enuStep >= ofsHasSent Then
      MsgBox "Operação não permitida.", vbExclamation, "Atenção"
      Exit Sub
    End If
    
    strText = "Confirma o 'Cancelamento do Pedido'?"
    
    If MsgBox(strText, vbQuestion + vbYesNo, "Confirmação de Status") = vbYes Then
      If frmWEB_OFChangeStatus.ChangeStatus( _
        .Fields("ID").Value, ofsCanceled, True, enuStep, _
        .Fields("Filial").Value, _
        CLng("0" & .Fields("Sequencia").Value)) Then
        
        Call ClearScreen
        Call ShowRecord
      End If
    End If
  End With

End Sub

Private Sub cmdNextStep_Click()
  Dim enuStep As enWEB_OrderFormStep
  Dim strText As String
  
  enuStep = mrsWEBOrderForms.Fields("Passo").Value
  
  Select Case enuStep
    Case ofsReceived
      If CLng("0" & mrsWEBOrderForms.Fields("Sequencia").Value) = 0 Then
        If MsgBox("Confirma a Criação da Venda?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
          Call CreateSale
        End If
      Else
        MsgBox "Efetue o recebimento na tela de Saídas para Confirmar o Pagamento.", vbExclamation, "Atenção"
      End If
      Exit Sub
      
    Case ofsConfirmedPayment
      strText = "Pedido Embalado (Recibo, Etiqueta)"
      enuStep = ofsPacked
    Case ofsPacked
      strText = "Pedido Enviado"
      enuStep = ofsHasSent
    Case Else
      MsgBox "Operação não permitida.", vbExclamation, "Atenção"
      Exit Sub
  End Select
  
  strText = "Confirma o seguinte passo: '" & strText & "'?"
  
  If MsgBox(strText, vbQuestion + vbYesNo, "Confirmação de Status") = vbYes Then
    If frmWEB_OFChangeStatus.ChangeStatus( _
      mrsWEBOrderForms.Fields("ID").Value, enuStep, True) Then
      Call ClearScreen
      Call ShowRecord
    End If
  End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Dim strSQL As String
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("Aguarde...")
  
  Call CenterForm(Me)
  
  Call ActiveBarLoadToolTips(Me)
  
  With ActiveBar1
    With .Tools("miFilter").CBList
      .Clear
      .InsertItem 0, "Todos"
      .InsertItem 1, "Pedido recebido"
      .InsertItem 2, "Pedido com pagamento confirmado"
      .InsertItem 3, "Pedido com produto embalado"
      .InsertItem 4, "Pedido enviado"
      .InsertItem 5, "Pedido cancelado"
      .InsertItem 6, "Personalizar..."
    End With
    With .Tools("miFilter")
      .Text = .CBList(0)
    End With
    .RecalcLayout
  End With
  
  strSQL = "SELECT * FROM WEB_OrderForms WHERE Filial = " & _
           gnCodFilial & " ORDER BY ID DESC"
  
  Set mrsWEBOrderForms = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  Call Navigate(navFirst)
  
  Call StatusMsg("")
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mrsWEBOrderForms.Close
  Set mrsWEBOrderForms = Nothing
End Sub

Private Sub lblChageStatus_Click(Index As Integer)
  If frmWEB_OFChangeStatus.ChangeStatus(mrsWEBOrderForms.Fields("ID").Value, _
    mrsWEBOrderForms.Fields("Passo").Value) Then
    Call ClearScreen
    Call ShowRecord
  End If
End Sub

Private Sub CreateSale()
  Dim rsSaidas As Recordset
  Dim rsSaidasProdutos As Recordset
  Dim rsItens As Recordset
  Dim rsListPrice As Recordset
  Dim strSQL As String
  Dim lngSequence As Long
  Dim bytFilial As Byte
  Dim intX As Integer
  Dim strCodProd As String
  Dim intErro As Integer
  Dim strListPrice As String
  Dim curDiscount As Currency
  Dim intCodOpVenda As Integer
  
  On Error GoTo ErrHadler
  
  Call StatusMsg("Aguarde...")
  Screen.MousePointer = vbHourglass
  
  'Inicia transação
  ws.BeginTrans
  
  'Filial
  bytFilial = mrsWEBOrderForms.Fields("Filial").Value
  'Nova sequência
  lngSequence = gnGetNextSequencia(CInt(bytFilial))
  
  'Atualiza a sequência do Pedido
  strSQL = "UPDATE WEB_OrderForms SET Sequencia = " & lngSequence & _
           " WHERE ID = " & mrsWEBOrderForms.Fields("ID").Value
  Call db.Execute(strSQL)
  
  'Tabela de preços
  strListPrice = Replace(LIST_PRICE_WEB, REPLACE_TQW, _
                         Format(mrsWEBOrderForms.Fields("ID").Value, _
                         String(Len(REPLACE_TQW), "0")))
  
  'Apaga qualquer produto que esteja utilizando a sequencia
  Call EraseTypeMoviment(tmSaidasProdutos, CInt(bytFilial), lngSequence)
  
  'Itens do pedido
  strSQL = "SELECT * FROM WEB_OrderItens WHERE OrderFormID = " & mrsWEBOrderForms.Fields("ID").Value
  Set rsItens = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  Set rsSaidasProdutos = db.OpenRecordset("Saídas - Produtos", dbOpenDynaset)
  Set rsListPrice = db.OpenRecordset("Preços", dbOpenDynaset)
  
  intX = 0
  With rsItens
    If Not .BOF And Not .EOF Then
      Do Until .EOF
        Call Acha_Produto(.Fields("sku").Value, strCodProd, 0, 0, 0, 0, intErro)
        If intErro = 0 Then
          'Inclui item na venda
          intX = intX + 1
          With rsSaidasProdutos
            .AddNew
            .Fields("Filial").Value = bytFilial
            .Fields("Sequência").Value = lngSequence
            .Fields("Linha").Value = intX
            .Fields("Código").Value = rsItens.Fields("sku").Value
            .Fields("Qtde").Value = rsItens.Fields("Quantity").Value
            .Fields("Preço").Value = rsItens.Fields("ListPrice").Value
            'Desconto
            If CCur("0" & rsItens.Fields("Discount").Value) > 0 Then
              .Fields("Desconto").Value = rsItens.Fields("Discount").Value / rsItens.Fields("ListPrice").Value * 100
              .Fields("Desconto Valor").Value = rsItens.Fields("Discount").Value
              curDiscount = curDiscount + .Fields("Desconto Valor").Value
            End If
            .Fields("ICM").Value = 0
            .Fields("IPI").Value = 0
            .Fields("Preço Final").Value = rsItens.Fields("Total").Value
            .Fields("Etiqueta").Value = False
            .Fields("Código sem Grade").Value = strCodProd
            .Fields("InGeradoViaConsig").Value = False
            .Fields("Situação Tributária").Value = " "
            .Fields("Unidade Venda").Value = " "
            .Fields("Descricao Adicional").Value = ""
            .Update
          End With
          'Inclui preço do item
          With rsListPrice
            .FindFirst "Tabela = '" & strListPrice & _
                       "' AND Produto = '" & strCodProd & "'"
            If .NoMatch Then
              .AddNew
              .Fields("Tabela").Value = strListPrice
              .Fields("Produto").Value = strCodProd
              .Fields("Preço").Value = rsItens.Fields("ListPrice").Value
              .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
              .Update
            End If
          End With
        Else
          ws.Rollback
          Call StatusMsg("Erro")
          Screen.MousePointer = vbDefault
          MsgBox "Produto [" & .Fields("sku").Value & "] não cadastrado.", _
            vbCritical, "Erro"
          Exit Sub
        End If
        .MoveNext
      Loop
    End If
  End With
  
  'Obtém o código para a operação de venda
  Call GetWEBCod_Op(0, intCodOpVenda, 0)
  
  'Venda (main)
  Set rsSaidas = db.OpenRecordset("Saídas", dbOpenDynaset)
  With rsSaidas
    .AddNew
    .Fields("Filial").Value = bytFilial
    .Fields("Data").Value = Date
    .Fields("Sequência").Value = lngSequence
    .Fields("Operação").Value = intCodOpVenda
    .Fields("Caixa").Value = gbytFirstCaixa()
    .Fields("Tabela").Value = strListPrice
    .Fields("Digitador").Value = 0
    .Fields("Operador").Value = 0
    .Fields("Cliente").Value = CLng(txtShopperCodigo.Text)
    '-----------------------------------------------------------------
    '19/04/2005 - Daniel
    'Solicitação: Aura Prata
    'Atualizar o campo "Faturado" do cadastro de clientes
      Call AtualizarFieldFaturadoCliFor(CLng(txtShopperCodigo.Text))
    '-----------------------------------------------------------------
    .Fields("Observações").Value = "Venda da Loja Virtual"
    .Fields("Produtos").Value = mrsWEBOrderForms.Fields("SubTotal").Value + curDiscount
    .Fields("Desconto").Value = curDiscount
    .Fields("Serviços").Value = 0
    .Fields("Base ISS").Value = 0
    .Fields("Valor ISS").Value = 0
    .Fields("Perc IR Sobre ISS").Value = 0
    .Fields("Valor IR Sobre ISS").Value = 0
    .Fields("IPI").Value = 0
    .Fields("Frete").Value = mrsWEBOrderForms.Fields("ShippingTotal").Value
    '-----------------------------------------------------------------
    '12/04/2005 - Daniel
    'Adicionado Seguro
    .Fields("Seguro").Value = IIf(IsNumeric(mrsWEBOrderForms.Fields("Seguro").Value), mrsWEBOrderForms.Fields("Seguro").Value, 0)
    '-----------------------------------------------------------------
    .Fields("Base ICM").Value = 0
    .Fields("Valor ICM").Value = 0
    .Fields("Base ICM Subs").Value = 0
    .Fields("Valor ICM Subs").Value = 0
    .Fields("Total").Value = mrsWEBOrderForms.Fields("Total").Value
    .Fields("Efetivada").Value = False
    .Fields("Recebimento").Value = False
    .Fields("Nota Impressa").Value = 0
    .Fields("Recebe - Conta").Value = 0
    .Fields("Recebe - Dinheiro").Value = 0
    .Fields("Recebe - Emp Cartão").Value = 0
    .Fields("Recebe - Num Cartão").Value = 0
    .Fields("Recebe - Cartão").Value = 0
    .Fields("Recebe - Vale").Value = 0
    .Fields("Referência").Value = ""
    .Fields("Nota Cancelada").Value = False
    .Fields("Movimentação Desfeita").Value = False
    .Fields("Total Vista").Value = 0
    .Fields("Total Prazo").Value = 0
    .Fields("Tipo Parcela").Value = ""
    .Fields("Conta").Value = 0
    .Fields("Prometido Para").Value = ""
    .Fields("Orçamento Aprovado").Value = ""
    .Fields("Data Acerto Empréstimo").Value = Null
    .Fields("Técnico").Value = 0
    .Fields("Cupom Fiscal Impresso").Value = False
    .Fields("Parcela Cartão").Value = ""
    .Fields("Qtde Parcelas").Value = 0
    .Fields("Valor Parcela").Value = 0
    .Fields("WEBOrderFormID").Value = mrsWEBOrderForms.Fields("ID").Value
    .Update
  End With
  
  'Finaliza transação
  ws.CommitTrans
  
  rsSaidas.Close
  Set rsSaidas = Nothing
  
  rsItens.Close
  Set rsItens = Nothing
  
  rsSaidasProdutos.Close
  Set rsSaidasProdutos = Nothing
  
  rsListPrice.Close
  Set rsListPrice = Nothing
  
  Call ClearScreen
  Call ShowRecord
  
  Call StatusMsg("Venda criada como sequência nº " & lngSequence)
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
ErrHadler:
  ws.Rollback
  Call StatusMsg("Erro")
  Screen.MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Function mstrGetDescShipping(ByVal bytID As Byte) As String
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Name FROM WEB_ShippingMethods WHERE ID = " & bytID
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      mstrGetDescShipping = .Fields("Name").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

Private Function mstrGetDescPayment(ByVal bytID As Byte) As String
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Name FROM WEB_PaymentMethods WHERE ID = " & bytID
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      mstrGetDescPayment = .Fields("Name").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

Private Sub AtualizarFieldFaturadoCliFor(ByVal lngCliente As Long)
  '19/04/2005 - Daniel
  Dim rstCliFor As Recordset
  Dim rstParame As Recordset
  Dim strSQL    As String
  
  On Error GoTo TratarErro
  
  strSQL = "SELECT CliWebComprarPrazo FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial
  
  Set rstParame = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstParame
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("CliWebComprarPrazo").Value Then
        strSQL = ""
        strSQL = "SELECT Faturado FROM Cli_For WHERE Código = " & lngCliente
        
        Set rstCliFor = db.OpenRecordset(strSQL, dbOpenDynaset)
        
        If Not (rstCliFor.BOF And rstCliFor.EOF) Then
          rstCliFor.MoveFirst
          rstCliFor.Edit
          rstCliFor.Fields("Faturado").Value = True
          rstCliFor.Update
          rstCliFor.Close
        End If
        
        Set rstCliFor = Nothing
        
      End If
      
    End If
    .Close
  End With
  
  Set rstParame = Nothing
  
  Exit Sub
  
TratarErro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Exit Sub

End Sub
