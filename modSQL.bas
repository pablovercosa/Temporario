Attribute VB_Name = "modSQL"
Option Explicit

Public Const SQL_CONS_PRODUTO_ATIVO As String = "SELECT * " & _
  "FROM Produtos WHERE C�digo <> '0' AND NOT [Desativado] " & _
  "ORDER BY Nome;"

Public Const SQL_CONS_PRODUTO_NORMAL As String = "SELECT Nome, C�digo " & _
  "FROM Produtos WHERE C�digo <> '0' AND NOT [Desativado] " & _
  "AND Tipo = 'N' ORDER BY Nome;"

Public Const SQL_CONS_PRODUTO_EDICAO As String = "SELECT Nome, C�digo " & _
  "FROM Produtos WHERE C�digo <> '0' AND NOT [Desativado] " & _
  "AND Tipo = 'E' ORDER BY Nome;"

Public Const SQL_CONS_PRODUTO_GRADE As String = "SELECT Nome, C�digo " & _
  "FROM Produtos WHERE C�digo <> '0' AND NOT [Desativado] " & _
  "AND Tipo = 'G' ORDER BY Nome;"

Public Const SQL_CONS_COR As String = "SELECT DISTINCTROW Nome, C�digo " & _
  "FROM Cores WHERE C�digo <> 0 ORDER BY Nome;"

Public Const SQL_CONS_TAB_PRECO_T1 As String = "SELECT Tabela FROM [Tabela de Pre�os] " & _
  "ORDER BY Tabela;"

Public Const SQL_CONS_TAB_PRECO_T2 As String = "SELECT DISTINCT Tabela FROM Pre�os " & _
  "ORDER BY Tabela;"

'---------------------------------------------------------------------------------
'07/05/2002 - mpdea
'
'Nova SQL para exibi��o das tabelas de pre�os
'>>-------------------------------------------------------------------------------
Public Const SQL_CONS_TAB_PRECO_SHOW As String = "SELECT DISTINCT " & _
  "[Tabela de Pre�os].Tabela FROM [Tabela de Pre�os] " & _
  "INNER JOIN Pre�os ON [Tabela de Pre�os].Tabela = Pre�os.Tabela WHERE " & _
  "[Tabela de Pre�os].Tabela <> 'CUSTO' ORDER BY [Tabela de Pre�os].Tabela"
'-------------------------------------------------------------------------------<<


