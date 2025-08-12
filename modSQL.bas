Attribute VB_Name = "modSQL"
Option Explicit

Public Const SQL_CONS_PRODUTO_ATIVO As String = "SELECT * " & _
  "FROM Produtos WHERE Código <> '0' AND NOT [Desativado] " & _
  "ORDER BY Nome;"

Public Const SQL_CONS_PRODUTO_NORMAL As String = "SELECT Nome, Código " & _
  "FROM Produtos WHERE Código <> '0' AND NOT [Desativado] " & _
  "AND Tipo = 'N' ORDER BY Nome;"

Public Const SQL_CONS_PRODUTO_EDICAO As String = "SELECT Nome, Código " & _
  "FROM Produtos WHERE Código <> '0' AND NOT [Desativado] " & _
  "AND Tipo = 'E' ORDER BY Nome;"

Public Const SQL_CONS_PRODUTO_GRADE As String = "SELECT Nome, Código " & _
  "FROM Produtos WHERE Código <> '0' AND NOT [Desativado] " & _
  "AND Tipo = 'G' ORDER BY Nome;"

Public Const SQL_CONS_COR As String = "SELECT DISTINCTROW Nome, Código " & _
  "FROM Cores WHERE Código <> 0 ORDER BY Nome;"

Public Const SQL_CONS_TAB_PRECO_T1 As String = "SELECT Tabela FROM [Tabela de Preços] " & _
  "ORDER BY Tabela;"

Public Const SQL_CONS_TAB_PRECO_T2 As String = "SELECT DISTINCT Tabela FROM Preços " & _
  "ORDER BY Tabela;"

'---------------------------------------------------------------------------------
'07/05/2002 - mpdea
'
'Nova SQL para exibição das tabelas de preços
'>>-------------------------------------------------------------------------------
Public Const SQL_CONS_TAB_PRECO_SHOW As String = "SELECT DISTINCT " & _
  "[Tabela de Preços].Tabela FROM [Tabela de Preços] " & _
  "INNER JOIN Preços ON [Tabela de Preços].Tabela = Preços.Tabela WHERE " & _
  "[Tabela de Preços].Tabela <> 'CUSTO' ORDER BY [Tabela de Preços].Tabela"
'-------------------------------------------------------------------------------<<


