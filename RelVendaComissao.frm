VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendaComissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relat�rio de Comiss�es por Vendedor"
   ClientHeight    =   3975
   ClientLeft      =   3630
   ClientTop       =   3120
   ClientWidth     =   6120
   Icon            =   "RelVendaComissao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   6120
   Begin VB.Data Data3 
      Appearance      =   0  'Flat
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   4080
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Apelido, C�digo FROM Funcion�rios WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   855
      Left            =   4560
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
      Begin VB.OptionButton optTipoAnalitico 
         Caption         =   "Anal�tico"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optTipoSintetico 
         Caption         =   "Sint�tico"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Per�odo"
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   5895
      Begin MSMask.MaskEdBox Data_Fim 
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
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
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Pressione F2 para Calend�rio"
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Op��es"
      Height          =   855
      Left            =   1560
      TabIndex        =   17
      Top             =   2280
      Width           =   2895
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optPedidos 
         Caption         =   "Pedidos"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   225
         Width           =   1215
      End
      Begin VB.OptionButton optOrcamentos 
         Caption         =   "Or�amentos"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   495
         Width           =   1215
      End
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   75
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Par�metro"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   2025
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Cliente"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sa�da"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
      Begin VB.OptionButton B_V�deo 
         Caption         =   "V�deo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton B_Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton B_Imprime 
      Caption         =   "Imprimir"
      Height          =   400
      Left            =   4680
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo Combo 
      Bindings        =   "RelVendaComissao.frx":058A
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   735
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
      Columns(0).Width=   8229
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1588
      Columns(1).Caption=   "Filial"
      Columns(1).Name =   "Filial"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Filial"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Cliente 
      Bindings        =   "RelVendaComissao.frx":059E
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   735
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
      Columns(0).Width=   5080
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2011
      Columns(1).Caption=   "C�digo"
      Columns(1).Name =   "C�digo"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "C�digo"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin Crystal.CrystalReport Rel1 
      Left            =   600
      Top             =   0
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Vendedor 
      Bindings        =   "RelVendaComissao.frx":05B2
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   855
      Width           =   735
      DataFieldList   =   "Apelido"
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
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "Apelido"
      Columns(0).Name =   "Apelido"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Apelido"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5080
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2011
      Columns(2).Caption=   "C�digo"
      Columns(2).Name =   "C�digo"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "C�digo"
      Columns(2).DataType=   3
      Columns(2).FieldLen=   256
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   3720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   915
      Width           =   855
   End
   Begin VB.Label Nome_Vendedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   22
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filial:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Nome_Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   15
      Top             =   60
      Width           =   3975
   End
   Begin VB.Label Nome_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   435
      Width           =   3975
   End
End
Attribute VB_Name = "frmRelVendaComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsParametros As Recordset
Private rsFuncionarios As Recordset
Private rsClientes As Recordset

Private Sub B_Imprime_Click()
 Dim Erro As Boolean
 Dim Str_Rel As String
 Dim strNomeArquivo As String
 Dim Data1 As Variant
 
 Dim strSQL As String
 Dim rsRelatorioComissao As Recordset
 Dim bolHaving As Boolean
 
 Call StatusMsg("")
 
 If Combo_Cliente.Text = "" Then Combo_Cliente.Text = 0
 If Combo_Vendedor.Text = "" Then Combo_Vendedor.Text = 0

 Rem Verifica empresa
 If IsNull(Nome_Empresa.Caption) Or Nome_Empresa.Caption = "" Then
   DisplayMsg "Escolha a filial."
   Combo.SetFocus
   Exit Sub
 End If

 If Filial_Liberada <> 0 Then
   If Val(Combo.Text) <> Filial_Liberada Then
     DisplayMsg "Funcion�rio n�o tem acesso a esta filial."
     Exit Sub
   End If
 End If


 Rem Verifica cliente
 If Nome_Cliente.Caption = "" And Val(Combo_Cliente.Text) <> 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Cliente incorreto, verifique."
   Combo_Cliente.SetFocus
   Exit Sub
 End If
 
  Rem Verifica Vendedor
 If Nome_Vendedor.Caption = "" And Val(Combo_Vendedor.Text) <> 0 Then Erro = True
 If Erro = True Then
   DisplayMsg "Vendedor incorreto, verifique."
   Combo_Vendedor.SetFocus
   Exit Sub
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
 
 dbTemp.Execute "DELETE * FROM RelVendaComissao"
 B_Imprime.Enabled = False
 
 If optTipoAnalitico Then

   strSQL = "SELECT [Sa�das - Produtos].Filial, [Sa�das - Produtos].Sequ�ncia, Sa�das.Data, Funcion�rios.C�digo, Funcion�rios.Nome, Funcion�rios.Comiss�o, Cli_For.C�digo, Cli_For.Nome, Sa�das.[Nota Impressa], Sum([Sa�das - Produtos].[Pre�o Final]) AS [PrecoFinal], Sum([Qtde]*[PrecoCusto]) AS CustoFinal, Sa�das.[Nota Cancelada], [Opera��es Sa�da].C�digo, [Opera��es Sa�da].Nome, [Opera��es Sa�da].Tipo, Sa�das.Efetivada "
   strSQL = strSQL & "FROM ((([Sa�das - Produtos] INNER JOIN Sa�das ON ([Sa�das - Produtos].Sequ�ncia = Sa�das.Sequ�ncia) AND ([Sa�das - Produtos].Filial = Sa�das.Filial)) INNER JOIN Funcion�rios ON Sa�das.Digitador = Funcion�rios.C�digo) INNER JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
   
   strSQL = strSQL & "WHERE (Sa�das.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) "
   strSQL = strSQL & " AND (Sa�das.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "
   
   If Combo_Vendedor.Text <> 0 Then
     strSQL = strSQL & "AND Funcion�rios.C�digo=" & Combo_Vendedor.Text & " "
   End If
   
   If Combo_Cliente.Text <> 0 Then
     strSQL = strSQL & "AND Cli_For.C�digo=" & Combo_Cliente.Text & " "
   End If
   
   strSQL = strSQL & "GROUP BY [Sa�das - Produtos].Filial, [Sa�das - Produtos].Sequ�ncia, Sa�das.Data, Funcion�rios.C�digo, Funcion�rios.Nome, Funcion�rios.Comiss�o, Cli_For.C�digo, Cli_For.Nome, Sa�das.[Nota Impressa], Sa�das.[Nota Cancelada], [Opera��es Sa�da].C�digo, [Opera��es Sa�da].Nome, [Opera��es Sa�da].Tipo, Sa�das.Efetivada "
   strSQL = strSQL & "Having Sa�das.[Nota Cancelada] <> -1 AND [Sa�das - Produtos].Filial =" & Val(Combo.Text) & " "
  
   If optOrcamentos Then
     strSQL = strSQL & "AND [Opera��es Sa�da].Tipo='O' "
   End If
   
   If optPedidos Then
     strSQL = strSQL & "AND [Opera��es Sa�da].Tipo='V' and Sa�das.Efetivada = -1 "
   End If
   
   If optTodos Then
     strSQL = strSQL & "AND ([Opera��es Sa�da].Tipo='V' and Sa�das.Efetivada = -1) OR [Opera��es Sa�da].Tipo='O' "
   End If
   
   strSQL = strSQL & "ORDER BY [Sa�das - Produtos].Filial, [Sa�das - Produtos].Sequ�ncia;  "
  
   Set rsRelatorioComissao = db.OpenRecordset(strSQL)
   
    With rsRelatorioComissao
      If (.BOF And .EOF) Then
        pgbProgress.min = 0
        pgbProgress.Max = 1
        pgbProgress.Value = 0
        MsgBox "N�o h� informa��es para serem exibidas no relat�rio. Verifique se os filtros foram preenchidos corretamente.", vbInformation, "Quick Store"
        B_Imprime.Enabled = True
        Exit Sub
      End If
      
      .MoveLast
      .MoveFirst
      
      pgbProgress.min = 0
      pgbProgress.Max = .RecordCount + 1
   
       Do Until .EOF
       
        DoEvents
       
        strSQL = "INSERT INTO RelVendaComissao "
        strSQL = strSQL & "( CodigoVendedor, NomeVendedor, Comissao, CodigoCliente, NomeCliente, Data, Sequencia, NotaFiscal, Custo, ValorFinal, Lucro, Indice, ComissaoValor, Tipo, NomeOperacao ) "
        strSQL = strSQL & "VALUES (" & .Fields("Funcion�rios.C�digo") & ","
        strSQL = strSQL & """" & .Fields("Funcion�rios.Nome") & ""","
        strSQL = strSQL & "" & Replace(.Fields("Comiss�o"), ",", ".") & ","
        strSQL = strSQL & "" & .Fields("Cli_For.C�digo") & ","
        strSQL = strSQL & """" & Replace(.Fields("Cli_For.Nome"), """", "'") & ""","
        strSQL = strSQL & "#" & Format(.Fields("Data"), "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "" & .Fields("Sequ�ncia") & ","
        strSQL = strSQL & "" & .Fields("Nota Impressa") & ","
        strSQL = strSQL & "" & Replace(.Fields("CustoFinal"), ",", ".") & ","
        strSQL = strSQL & "" & Replace(.Fields("PrecoFinal"), ",", ".") & ","
        strSQL = strSQL & "" & Replace(.Fields("PrecoFinal") - .Fields("CustoFinal"), ",", ".") & ","
        
        If .Fields("CustoFinal") = 0 Then
          strSQL = strSQL & "0,"
        Else
          strSQL = strSQL & "" & Replace(.Fields("PrecoFinal") / .Fields("CustoFinal"), ",", ".") & ","
        End If
        
        strSQL = strSQL & "" & Replace(.Fields("PrecoFinal") * (.Fields("Comiss�o") / 100), ",", ".") & ","
        strSQL = strSQL & "'" & .Fields("Tipo") & "',"
        strSQL = strSQL & "'" & .Fields("Opera��es Sa�da.Nome") & "')"
            
        dbTemp.Execute strSQL, dbFailOnError
          
        pgbProgress.Value = .AbsolutePosition
        .MoveNext
          
       Loop
       
   End With
   
   rsRelatorioComissao.Close
   Set rsRelatorioComissao = Nothing
 Else
   strSQL = "SELECT Funcion�rios.C�digo, Funcion�rios.Nome, Funcion�rios.Comiss�o, Cli_For.C�digo, Cli_For.Nome, Sum([Sa�das - Produtos].[Pre�o Final]) AS PrecoFinal, Sum([Qtde]*[PrecoCusto]) AS CustoFinal "
   strSQL = strSQL & "FROM ((([Sa�das - Produtos] INNER JOIN Sa�das ON ([Sa�das - Produtos].Filial = Sa�das.Filial) AND ([Sa�das - Produtos].Sequ�ncia = Sa�das.Sequ�ncia)) INNER JOIN Funcion�rios ON Sa�das.Digitador = Funcion�rios.C�digo) INNER JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
   strSQL = strSQL & "WHERE Sa�das.[Nota Cancelada] <> -1 AND [Sa�das - Produtos].Filial =" & Val(Combo.Text) & " "
   strSQL = strSQL & " AND (Sa�das.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) "
   strSQL = strSQL & " AND (Sa�das.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#) "
  
   If optOrcamentos Then
     strSQL = strSQL & "AND [Opera��es Sa�da].Tipo='O' "
   End If
   
   If optPedidos Then
     strSQL = strSQL & "AND [Opera��es Sa�da].Tipo='V' and Sa�das.Efetivada = -1 "
   End If
   
   If optTodos Then
     strSQL = strSQL & "AND ([Opera��es Sa�da].Tipo='V' and Sa�das.Efetivada = -1) OR (((Sa�das.Data >= #" & Format(Data_Ini.Text, "mm/dd/yyyy") & "#) AND (Sa�das.Data <= #" & Format(Data_Fim.Text, "mm/dd/yyyy") & "#)) AND [Opera��es Sa�da].Tipo='O') "
   End If
   
   strSQL = strSQL & "GROUP BY Funcion�rios.C�digo, Funcion�rios.Nome, Funcion�rios.Comiss�o, Cli_For.C�digo, Cli_For.Nome "
   
   bolHaving = False
   
   If Combo_Vendedor.Text <> 0 Then
     strSQL = strSQL & "Having Funcion�rios.C�digo=" & Combo_Vendedor.Text & " "
     bolHaving = True
   End If
   
   If Combo_Cliente.Text <> 0 Then
     If bolHaving Then
        strSQL = strSQL & "AND Cli_For.C�digo=" & Combo_Cliente.Text & " "
     Else
        strSQL = strSQL & "Having Cli_For.C�digo=" & Combo_Cliente.Text & " "
     End If
   End If
  
   Set rsRelatorioComissao = db.OpenRecordset(strSQL)
   
    With rsRelatorioComissao
      If (.BOF And .EOF) Then
        pgbProgress.min = 0
        pgbProgress.Max = 1
        pgbProgress.Value = 0
        MsgBox "N�o h� informa��es para serem exibidas no relat�rio. Verifique se os filtros foram preenchidos corretamente.", vbInformation, "Quick Store"
        B_Imprime.Enabled = True
        Exit Sub
      End If
      
      .MoveLast
      .MoveFirst
      
      pgbProgress.min = 0
      pgbProgress.Max = .RecordCount + 1
   
       Do Until .EOF
       
        DoEvents
       
        strSQL = "INSERT INTO RelVendaComissao "
        strSQL = strSQL & "( CodigoVendedor, NomeVendedor, Comissao, CodigoCliente, NomeCliente, Custo, ValorFinal, Lucro, Indice, ComissaoValor ) "
        strSQL = strSQL & "VALUES (" & .Fields("Funcion�rios.C�digo") & ","
        strSQL = strSQL & """" & .Fields("Funcion�rios.Nome") & ""","
        strSQL = strSQL & "" & Replace(.Fields("Comiss�o"), ",", ".") & ","
        strSQL = strSQL & "" & .Fields("Cli_For.C�digo") & ","
        strSQL = strSQL & """" & Replace(.Fields("Cli_For.Nome"), """", "'") & ""","
        strSQL = strSQL & "" & Replace(.Fields("CustoFinal"), ",", ".") & ","
        strSQL = strSQL & "" & Replace(.Fields("PrecoFinal"), ",", ".") & ","
        strSQL = strSQL & "" & Replace(.Fields("PrecoFinal") - .Fields("CustoFinal"), ",", ".") & ","
        
        If .Fields("CustoFinal") = 0 Then
          strSQL = strSQL & "0,"
        Else
          strSQL = strSQL & "" & Replace(.Fields("PrecoFinal") / .Fields("CustoFinal"), ",", ".") & ","
        End If
        
        strSQL = strSQL & "" & Replace(.Fields("PrecoFinal") * (.Fields("Comiss�o") / 100), ",", ".") & ")"
 
        dbTemp.Execute strSQL, dbFailOnError
          
        pgbProgress.Value = .AbsolutePosition
        .MoveNext
          
       Loop
       
   End With
   
   rsRelatorioComissao.Close
   Set rsRelatorioComissao = Nothing
 End If
 
 B_Imprime.Enabled = True
 
  pgbProgress.min = 0
  pgbProgress.Max = 1
  pgbProgress.Value = 0
  
 Rem  Nome do BD
 Rel1.DataFiles(0) = gsQuickDBFileName
 Rel1.DataFiles(1) = gsTempDBFileName
 
 Rem Sa�da
 If B_V�deo = True Then Rel1.Destination = 0
 If B_Impressora = True Then Rel1.Destination = 1
 
 Rem Nome do arquivo .rpt
 If optTipoSintetico.Value Then
   strNomeArquivo = gsReportPath & "rptVendasComissoesSintetico.rpt"
 ElseIf optTipoAnalitico.Value Then
   strNomeArquivo = gsReportPath & "rptVendasComissoesAnalitico.rpt"
 End If
 Rel1.ReportFileName = strNomeArquivo
 
 ' Modelo 1 ou 2
 'SetPrinterModeloPwd2 Rel1

 Str_Rel = "nome_empresa = '"
 Str_Rel = Str_Rel + gsNomeEmpresa + "'"
 Rel1.Formulas(1) = Str_Rel

 Rem data inicial
 Str_Rel = "Periodo = '"
 Str_Rel = Str_Rel + Data_Ini.Text + " - " + Data_Fim.Text + "'"
 Rel1.Formulas(2) = Str_Rel

 Rem data final
 Str_Rel = "Ordenacao = '"
 If optTodos Then
   Str_Rel = Str_Rel + "Todos (Pedidos de Venda e Or�amentos de Venda) '"
 ElseIf optPedidos Then
   Str_Rel = Str_Rel + "Pedidos de Venda'"
 ElseIf optOrcamentos Then
   Str_Rel = Str_Rel + "Or�amentos de Venda'"
 End If
 Rel1.Formulas(3) = Str_Rel


 Call StatusMsg("Aguarde, imprimindo...")
 MousePointer = vbHourglass
 
  
  '25/07/2003 - mpdea
  'Seta a impressora para relat�rio
  Call SetPrinterName("REL", Rel1)
  

 Rel1.Action = 1

 Call StatusMsg("")
 MousePointer = vbDefault

End Sub


Private Sub Combo_Vendedor_CloseUp()
  Combo_Vendedor.Text = Combo_Vendedor.Columns(2).Text
  Combo_Vendedor_LostFocus
End Sub

Private Sub Combo_Vendedor_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Vendedor_LostFocus()
  Call StatusMsg("")
  Nome_Vendedor.Caption = ""
  If IsNull(Combo_Vendedor.Text) Then Exit Sub
  If Not IsNumeric(Combo_Vendedor.Text) Then Exit Sub

  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Combo_Vendedor.Text
  If Not rsFuncionarios.NoMatch Then
    Nome_Vendedor.Caption = rsFuncionarios("Apelido")
  Else
    Combo_Vendedor.Text = 0
  End If

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
  
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  Data2.DatabaseName = gsQuickDBFileName
  Data3.DatabaseName = gsQuickDBFileName
  
  Combo.Text = gnCodFilial
  
End Sub
Private Sub Combo_Cliente_GotFocus()
  Call StatusMsg(LoadResString(50))
End Sub

Private Sub Combo_Cliente_CloseUp()
  Combo_Cliente.Text = Combo_Cliente.Columns(1).Text
  Combo_Cliente_LostFocus
End Sub

Private Sub Combo_Cliente_LostFocus()
  Call StatusMsg("")
  Nome_Cliente.Caption = ""
  If IsNull(Combo_Cliente.Text) Then Exit Sub
  If Not IsNumeric(Combo_Cliente.Text) Then Exit Sub

  rsClientes.Index = "C�digo"
  rsClientes.Seek "=", Combo_Cliente.Text
  If Not rsClientes.NoMatch Then
    Nome_Cliente.Caption = rsClientes("Nome")
  Else
    Combo_Cliente.Text = 0
  End If

End Sub

