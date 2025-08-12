VERSION 5.00
Begin VB.Form frmApagaMovim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apagar Movimenta��o ou Zerar Estoque de Produtos"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   ControlBox      =   0   'False
   Icon            =   "ApagaMovim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   10950
   Begin VB.Frame frm_zeraEstoqueProdutos 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5910
      TabIndex        =   5
      Top             =   330
      Visible         =   0   'False
      Width           =   4995
      Begin VB.OptionButton opt_frameNegativos 
         Appearance      =   0  'Flat
         Caption         =   "Produtos com estoque negativo"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2130
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton opt_frameTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos produtos"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   540
         TabIndex        =   6
         Top             =   180
         Width           =   1425
      End
   End
   Begin VB.OptionButton opt_zerarEstoque 
      Appearance      =   0  'Flat
      Caption         =   "Zerar Estoque de Produtos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5910
      TabIndex        =   4
      Top             =   90
      Width           =   2745
   End
   Begin VB.OptionButton opt_apagaMov 
      Appearance      =   0  'Flat
      Caption         =   "Apagar Movimenta��o"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3030
      TabIndex        =   3
      Top             =   90
      Value           =   -1  'True
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
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
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3270
      Width           =   10845
   End
   Begin VB.CommandButton cmdZero 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Apagar Movimenta��o"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2700
      Width           =   10845
   End
   Begin VB.Label lblMsg 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   60
      TabIndex        =   2
      Top             =   945
      Width           =   10815
   End
End
Attribute VB_Name = "frmApagaMovim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdZero_Click()
  On Error GoTo Erro
  
  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If


  If opt_apagaMov.Value = True Then
      gsTitle = LoadResString(201)
      gsMsg = "Tem certeza que deseja apagar todas as movimenta��es ?"
      gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      If gnResponse = vbNo Then
        Unload Me
        Exit Sub
      End If
      
      Screen.MousePointer = vbHourglass
      Call GetNumberOfUsers
      Screen.MousePointer = vbDefault
      If gnCtCurrentUsers > 1 Then
        Beep
        gsTitle = LoadResString(201)
        gsMsg = "Esta opera��o somente poder� ser feita ap�s todas as demais esta��es em rede fecharem suas respectivas se��es."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Exit Sub
      End If
    
      Call ZeroMovim
      
      Unload Me
  Else
  
      Dim rsEstoqueFinalProd As Recordset
      Dim sTamanho As String
      Dim sCor As String
      Dim sCodigoProduto As String
      Dim sSQL As String
      
      If opt_frameNegativos.Value = True Then
          ' Zerar estoque de produtos com ESTOQUE NEGATIVO
          
          If MsgBox("Tem certeza que deseja zerar estoque dos produtos com ESTOQUE NEGATIVO?", vbYesNo + vbQuestion, "Aten��o") = vbYes Then
              Set rsEstoqueFinalProd = db.OpenRecordset("SELECT * FROM [Estoque Final] Where filial = " & gnCodFilial & " and [Estoque Atual] < 0 ", dbReadOnly)
              
              rsEstoqueFinalProd.MoveLast
              rsEstoqueFinalProd.MoveFirst
              
              While Not rsEstoqueFinalProd.EOF
                  ws.BeginTrans
                  
                  sSQL = "Update Estoque set [Estoque Anterior]=0, Vendas=0, [Valor Vendas]=0, Compras=0, [Valor Compras]=0, "
                  sSQL = sSQL & " [Transf Sa�da]=0, [Valor T Sa�da]=0,[Transf Entra]=0,[Valor T Entra]=0,"
                  sSQL = sSQL & " [Ajuste Sa�da]=0, [Valor Ajuste Sa�da]=0,[Ajuste Entra]=0,[Valor Ajuste Entra]=0,"
                  sSQL = sSQL & " [Gr�tis Sa�da]=0, [Valor Gr�tis Sa�da]=0,[Gr�tis Entra]=0,[Valor Gr�tis Entra]=0,"
                  sSQL = sSQL & " [Quebras]=0, [Valor Quebras]=0,[Empre Sa�da]=0,[Valor Empre Sa�da]=0,"
                  sSQL = sSQL & " [Empre Entra]=0, [Valor Empre Entra]=0,[Devolu��o]=0,[Valor Devolu��o]=0,"
                  sSQL = sSQL & " [Estoque Final]=0 "
                  sSQL = sSQL & " where filial = " & gnCodFilial & " and Data = (Select max(data) from Estoque where filial = " & gnCodFilial
                  sSQL = sSQL & " and Produto='" & rsEstoqueFinalProd.Fields("Produto") & "' "
                  sSQL = sSQL & " and Tamanho=" & rsEstoqueFinalProd.Fields("Tamanho")
                  sSQL = sSQL & " and Cor=" & rsEstoqueFinalProd.Fields("Cor") & " )"
                  db.Execute sSQL
              
                  ws.CommitTrans
                  rsEstoqueFinalProd.MoveNext
              Wend
              rsEstoqueFinalProd.Close
              Set rsEstoqueFinalProd = Nothing
              
              ws.BeginTrans
              db.Execute "Update [Estoque Final] set [Estoque Atual]=0 where filial = " & gnCodFilial & " and [Estoque Atual] < 0 "
              ws.CommitTrans
          
              MsgBox "Zerado com sucesso o estoque dos produtos que tinham ESTOQUE NEGATIVO.", vbInformation, "Sucesso"
          End If
      Else
          ' Zerar estoque de TODOS os produtos cadastrados
          
          If MsgBox("Tem certeza que deseja zerar o estoque de TODOS os produtos cadastrados?", vbYesNo + vbQuestion, "Aten��o") = vbYes Then
              Set rsEstoqueFinalProd = db.OpenRecordset("SELECT * FROM [Estoque Final] Where filial = " & gnCodFilial, dbReadOnly)
              
              rsEstoqueFinalProd.MoveLast
              rsEstoqueFinalProd.MoveFirst
              
              While Not rsEstoqueFinalProd.EOF
                  ws.BeginTrans
                  
                  sSQL = "Update Estoque set [Estoque Anterior]=0, Vendas=0, [Valor Vendas]=0, Compras=0, [Valor Compras]=0, "
                  sSQL = sSQL & " [Transf Sa�da]=0, [Valor T Sa�da]=0,[Transf Entra]=0,[Valor T Entra]=0,"
                  sSQL = sSQL & " [Ajuste Sa�da]=0, [Valor Ajuste Sa�da]=0,[Ajuste Entra]=0,[Valor Ajuste Entra]=0,"
                  sSQL = sSQL & " [Gr�tis Sa�da]=0, [Valor Gr�tis Sa�da]=0,[Gr�tis Entra]=0,[Valor Gr�tis Entra]=0,"
                  sSQL = sSQL & " [Quebras]=0, [Valor Quebras]=0,[Empre Sa�da]=0,[Valor Empre Sa�da]=0,"
                  sSQL = sSQL & " [Empre Entra]=0, [Valor Empre Entra]=0,[Devolu��o]=0,[Valor Devolu��o]=0,"
                  sSQL = sSQL & " [Estoque Final]=0 "
                  sSQL = sSQL & " where filial = " & gnCodFilial & " and Data = (Select max(data) from Estoque where filial = " & gnCodFilial
                  sSQL = sSQL & " and Produto='" & rsEstoqueFinalProd.Fields("Produto") & "' "
                  sSQL = sSQL & " and Tamanho=" & rsEstoqueFinalProd.Fields("Tamanho")
                  sSQL = sSQL & " and Cor=" & rsEstoqueFinalProd.Fields("Cor") & " )"
                  db.Execute sSQL
              
              
                  ws.CommitTrans
                  rsEstoqueFinalProd.MoveNext
              Wend
              rsEstoqueFinalProd.Close
              Set rsEstoqueFinalProd = Nothing
              
              ws.BeginTrans
              db.Execute "Update [Estoque Final] set [Estoque Atual]=0 where filial = " & gnCodFilial
              ws.CommitTrans
          
              MsgBox "Zerado com sucesso o estoque de TODOS os produtos cadastrados.", vbInformation, "Sucesso"
          End If
      
      End If
      
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro " & Err.Number & " - " & Err.Description, vbInformation, "Aten��o"
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  lblMsg.Caption = "ATEN��O: Toda a movimento de dados do BANCO DE DADOS ser� apagada. S�o elas: Caixa, Comiss�o, Comiss�o Servi�os, Consigna��o Entrada, Consigna��o Sa�da, Conta Cliente, Contas a Pagar e Contas a Receber."
  lblMsg.Caption = lblMsg.Caption & vbCrLf & "Entretanto, os dados cadastrais ser�o mantidos."
  lblMsg.Caption = lblMsg.Caption & vbCrLf & "� recomend�vel fazer um backup ANTES DE APAG�-LAS, caso estas movimenta��es n�o sejam de treinamento."
  lblMsg.Caption = lblMsg.Caption & vbCrLf & vbCrLf & "Esta opera��o poder� levar alguns minutos, caso seu banco de dados estiver com grande volume de informa��o."
End Sub

Private Sub ZeroMovim()
  Dim sSQL As String
  
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  
  Call ws.BeginTrans
  
  sSQL = "DELETE * FROM [Caixa]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Comiss�o]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Comiss�o Servi�os]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Consigna��o Entrada]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Consigna��o Sa�da]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Conta Cliente]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Contas a Pagar]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Contas a Receber]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  '10/09/2007 - Anderson
  'Gera arquivo log do sistema
  If g_bolSystemLog Then
    SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, sSQL, "frmApagaMovim_ZeroMovim", "Contas a Receber", g_strArquivoSystemLog
  End If
  
  sSQL = "DELETE * FROM [Entradas]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Entradas - Produtos]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Estoque]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Estoque Final]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Etiquetas]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Etiquetas - Tempo]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Grade - Tempo]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Lan�amentos Banc�rios]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Livro Ponto]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Mala Direta - Tempo]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Movimento - Cheques]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Movimento - Parcelas]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Pre�os - Tempo]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Resumo Clientes]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Resumo Di�rio]"
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Sa�das]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Sa�das - Produtos]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "DELETE * FROM [Sa�das - Servi�os]"
  Call StatusMsg("Apagando " & Mid(sSQL, InStr(sSQL, "["), 100) & " ...")
  Call db.Execute(sSQL, dbFailOnError)
  
  sSQL = "UPDATE [Par�metros Filial] SET [�ltima Movimenta��o] = 0, [�ltima Nota] = 0;"
  Call db.Execute(sSQL, dbFailOnError)
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
  gsTitle = LoadResString(201)
  gsMsg = "*** CONFIRMAR OPERA��O DE ""Apagar Movimentos""?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Call ws.Rollback
    DisplayMsg "Opera��o cancelada. Movimenta��o mantida."
    Exit Sub
  End If
  
  Call ws.CommitTrans
  
  DisplayMsg "Movimenta��o zerada. Recomenda-se realizar mais tarde uma compacta��o da base de dados."
  
  Exit Sub
  
  
ErrHandler:
  DisplayMsg "Erro ao Apagar Movimentos." & vbCrLf & CStr(Err.Number) & " - " & Err.Description
  Exit Sub
  
End Sub

Private Sub opt_apagaMov_Click()
    If opt_apagaMov.Value = True Then
        lblMsg.Caption = "ATEN��O: Toda a movimento de dados do BANCO DE DADOS ser� apagada. S�o elas: Caixa, Comiss�o, Comiss�o Servi�os, Consigna��o Entrada, Consigna��o Sa�da, Conta Cliente, Contas a Pagar e Contas a Receber."
        lblMsg.Caption = lblMsg.Caption & vbCrLf & "Entretanto, os dados cadastrais ser�o mantidos."
        lblMsg.Caption = lblMsg.Caption & vbCrLf & "� recomend�vel fazer um backup ANTES DE APAG�-LAS, caso estas movimenta��es n�o sejam de treinamento."
        lblMsg.Caption = lblMsg.Caption & vbCrLf & vbCrLf & "Esta opera��o poder� levar alguns minutos, caso seu banco de dados estiver com grande volume de informa��o."
          
        frm_zeraEstoqueProdutos.Visible = False
        cmdZero.Caption = "Apagar Movimenta��o"
    End If
End Sub

Private Sub opt_zerarEstoque_Click()
    If opt_zerarEstoque.Value = True Then
        lblMsg.Caption = "ATEN��O: O estoque dos produtos ser� ZERADO."
        lblMsg.Caption = lblMsg.Caption & vbCrLf & "Escolha se deseja zerar o estoque dos produtos COM ESTOQUE NEGATIVO ou DE TODOS OS PRODUTOS CADASTRADOS."
    
        frm_zeraEstoqueProdutos.Visible = True
        cmdZero.Caption = "Zerar Estoque de produtos"
    End If
End Sub
