Attribute VB_Name = "modODBC"
Option Explicit

'03/11/2005 - mpdea
'Flag indicando o uso de conexão ODBC com o sistema
Public g_bln_odbc As Boolean
'DSN para conexão ODBC com a base de dados Quick Store
Public g_str_dsn_quickstore As String
