--
-- DUMP FILE
--
-- Database is ported from MS Access
--------------------------------------------------------------------
-- Created using "MS Access to MSSQL" form http://www.bullzip.com
-- Program Version 5.5.281
--
-- OPTIONS:
--   sourcefilename=C:/Projetos/QuickStore/0_PastaZero_Legado/QuickStore2003/Temp.mdb
--   sourceusername=
--   sourcepassword=
--   sourcesystemdatabase=
--   destinationserver=CAGEPAR-000261\SQLEXPRESS
--   destinationauthentication=SQL
--   destinationdatabase=Temp
--   dropdatabase=1
--   createtables=1
--   unicode=1
--   autocommit=1
--   transferdefaultvalues=1
--   transferindexes=1
--   transferautonumbers=1
--   transferrecords=0
--   columnlist=0
--   tableprefix=
--   negativeboolean=0
--   ignorelargeblobs=0
--   memotype=VARCHAR(MAX)
--   datetimetype=DATETIME2
--

USE [master]
IF EXISTS (SELECT * FROM master.dbo.sysdatabases WHERE name = N'Temp') ALTER DATABASE [Temp] SET SINGLE_USER With ROLLBACK IMMEDIATE
IF EXISTS (SELECT * FROM master.dbo.sysdatabases WHERE name = N'Temp') DROP DATABASE [Temp]
IF NOT EXISTS (SELECT * FROM master.dbo.sysdatabases WHERE name = N'Temp') CREATE DATABASE [Temp]
USE [Temp]

--
-- Table structure for table 'Acerto'
--

IF object_id(N'Acerto', 'U') IS NOT NULL DROP TABLE [Acerto]

CREATE TABLE [Acerto] (
  [Filial] SMALLINT DEFAULT 0, 
  [DataAcerto] DATETIME2, 
  [CodigoProduto] NVARCHAR(20), 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [PrecoCusto] FLOAT DEFAULT 0, 
  [Fornecedor] INT DEFAULT 0, 
  [PrecoVenda] FLOAT DEFAULT 0
)

--
-- Table structure for table 'AcertoConsignacaoEntrada'
--

IF object_id(N'AcertoConsignacaoEntrada', 'U') IS NOT NULL DROP TABLE [AcertoConsignacaoEntrada]

CREATE TABLE [AcertoConsignacaoEntrada] (
  [Filial] SMALLINT DEFAULT 0, 
  [Sequencia] INT DEFAULT 0, 
  [DataAcerto] DATETIME2, 
  [LinhaProduto] SMALLINT DEFAULT 0, 
  [CodigoProduto] NVARCHAR(20), 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [FilialVenda] SMALLINT DEFAULT 0, 
  [SequenciaVenda] INT DEFAULT 0, 
  [PrecoCusto] FLOAT DEFAULT 0, 
  [Fornecedor] INT DEFAULT 0, 
  [PrecoVenda] FLOAT DEFAULT 0
)

--
-- Table structure for table 'Acompa Estoque'
--

IF object_id(N'Acompa Estoque', 'U') IS NOT NULL DROP TABLE [Acompa Estoque]

CREATE TABLE [Acompa Estoque] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Nome] NVARCHAR(100), 
  [Ordenação] NVARCHAR(20), 
  [Fracionado] BIT, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Nome Tamanho] NVARCHAR(30), 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Nome Cor] NVARCHAR(30), 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Nome Edição] NVARCHAR(30), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(30), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(30), 
  [Saldo Anterior] REAL DEFAULT 0, 
  [Transf Entrada] REAL DEFAULT 0, 
  [Compras] REAL DEFAULT 0, 
  [Transf Saída] REAL DEFAULT 0, 
  [Vendas] REAL DEFAULT 0, 
  [Saldo] REAL DEFAULT 0, 
  [Giro] REAL DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código], [Tamanho], [Cor], [Edição])
)

--
-- Table structure for table 'ActiveBar'
--

IF object_id(N'ActiveBar', 'U') IS NOT NULL DROP TABLE [ActiveBar]

CREATE TABLE [ActiveBar] (
  [Numero] INT NOT NULL DEFAULT 0, 
  [ToolID] INT DEFAULT 0, 
  PRIMARY KEY ([Numero])
)

--
-- Table structure for table 'Analítico'
--

IF object_id(N'Analítico', 'U') IS NOT NULL DROP TABLE [Analítico]

CREATE TABLE [Analítico] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Nome] NVARCHAR(100), 
  [Código Ordenação] NVARCHAR(20), 
  [Unidade Venda] NVARCHAR(5), 
  [Estoque] REAL DEFAULT 0, 
  [Vendas Penúltimo] REAL DEFAULT 0, 
  [Valor Vendas Penúltimo] FLOAT DEFAULT 0, 
  [Vendas Último] REAL DEFAULT 0, 
  [Valor Vendas Último] FLOAT DEFAULT 0, 
  [Vendas Atual] REAL DEFAULT 0, 
  [Valor Vendas Atual] FLOAT DEFAULT 0, 
  [Tendência] REAL DEFAULT 0, 
  [Vendas Último AA] REAL DEFAULT 0, 
  [Valor Vendas Último AA] FLOAT DEFAULT 0, 
  [Vendas AA] REAL DEFAULT 0, 
  [Valor Vendas AA] FLOAT DEFAULT 0, 
  [Vendas Próximo AA] REAL DEFAULT 0, 
  [Valor Vendas Próximo AA] FLOAT DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Produto])
)

--
-- Table structure for table 'Cep'
--

IF object_id(N'Cep', 'U') IS NOT NULL DROP TABLE [Cep]

CREATE TABLE [Cep] (
  [Codigo] INT, 
  [Cidade_estado] NVARCHAR(255), 
  [Bairro] NVARCHAR(255), 
  [Logradouro] NVARCHAR(255)
)

--
-- Table structure for table 'CEP_GERAL_ImportErrors'
--

IF object_id(N'CEP_GERAL_ImportErrors', 'U') IS NOT NULL DROP TABLE [CEP_GERAL_ImportErrors]

CREATE TABLE [CEP_GERAL_ImportErrors] (
  [Error] NVARCHAR(255), 
  [Field] NVARCHAR(255), 
  [Row] INT
)

--
-- Table structure for table 'Comissao'
--

IF object_id(N'Comissao', 'U') IS NOT NULL DROP TABLE [Comissao]

CREATE TABLE [Comissao] (
  [Data] DATETIME2, 
  [Vendedor] INT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Comissao] FLOAT DEFAULT 0, 
  [Cliente] INT DEFAULT 0, 
  [Filial] SMALLINT DEFAULT 0
)

--
-- Table structure for table 'ComissoesRetencao'
--

IF object_id(N'ComissoesRetencao', 'U') IS NOT NULL DROP TABLE [ComissoesRetencao]

CREATE TABLE [ComissoesRetencao] (
  [Vendedor] INT DEFAULT 0, 
  [Sequencia] INT DEFAULT 0, 
  [QtdeItens] REAL DEFAULT 0, 
  [VlPagoSemCartao] FLOAT DEFAULT 0, 
  [VlPagoComCartao] FLOAT DEFAULT 0, 
  [TaxaRetencao] REAL DEFAULT 0, 
  [VlPagoComCartaoRetendo] FLOAT DEFAULT 0, 
  [DescontoSubTotal] FLOAT DEFAULT 0, 
  [ValorParaCalculoComi] FLOAT DEFAULT 0, 
  [Comissao] FLOAT DEFAULT 0, 
  [QtdeOperacao] INT DEFAULT 0
)

--
-- Table structure for table 'ComissoesRetencaoGroup'
--

IF object_id(N'ComissoesRetencaoGroup', 'U') IS NOT NULL DROP TABLE [ComissoesRetencaoGroup]

CREATE TABLE [ComissoesRetencaoGroup] (
  [Vendedor] INT DEFAULT 0, 
  [QtdeItens] REAL DEFAULT 0, 
  [VlPagoSemCartao] FLOAT DEFAULT 0, 
  [VlPagoComCartao] FLOAT DEFAULT 0, 
  [VlPagoComCartaoRetendo] FLOAT DEFAULT 0, 
  [DescontoSubTotal] FLOAT DEFAULT 0, 
  [ValorParaCalculoComi] FLOAT DEFAULT 0, 
  [Comissao] FLOAT DEFAULT 0, 
  [QtdeOperacao] INT DEFAULT 0
)

--
-- Table structure for table 'Competencia'
--

IF object_id(N'Competencia', 'U') IS NOT NULL DROP TABLE [Competencia]

CREATE TABLE [Competencia] (
  [Centro] INT DEFAULT 0, 
  [Emissao] DATETIME2, 
  [Vencimento] DATETIME2, 
  [Descricao] NVARCHAR(30), 
  [Fornecedor] INT DEFAULT 0, 
  [Nota] NVARCHAR(15), 
  [Sequencia] INT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [Acrescimo] FLOAT DEFAULT 0, 
  [ValorPago] FLOAT DEFAULT 0, 
  [Pagamento] DATETIME2
)

--
-- Table structure for table 'CompetenciaGroup'
--

IF object_id(N'CompetenciaGroup', 'U') IS NOT NULL DROP TABLE [CompetenciaGroup]

CREATE TABLE [CompetenciaGroup] (
  [Centro] INT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [Acrescimo] FLOAT DEFAULT 0, 
  [ValorPago] FLOAT DEFAULT 0
)

--
-- Table structure for table 'Comprar'
--

IF object_id(N'Comprar', 'U') IS NOT NULL DROP TABLE [Comprar]

CREATE TABLE [Comprar] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Código Ordenação] NVARCHAR(20), 
  [Nome] NVARCHAR(100), 
  [Unidade Venda] NVARCHAR(5), 
  [Fracionado] BIT, 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(30), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(30), 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Nome Cor] NVARCHAR(30), 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Nome Tamanho] NVARCHAR(30), 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Nome Edição] NVARCHAR(30), 
  [Estoque] REAL DEFAULT 0, 
  [Ideal] REAL DEFAULT 0, 
  [Último Custo] FLOAT DEFAULT 0, 
  [Fornece1] INT DEFAULT 0, 
  [Nome1] NVARCHAR(100), 
  [Tel1_1] NVARCHAR(15), 
  [Tel1_2] NVARCHAR(15), 
  [Fax1] NVARCHAR(15), 
  [Fornece2] INT DEFAULT 0, 
  [Nome2] NVARCHAR(100), 
  [Tel2_1] NVARCHAR(15), 
  [Tel2_2] NVARCHAR(15), 
  [Fax2] NVARCHAR(15), 
  [Fornece3] INT DEFAULT 0, 
  [Nome3] NVARCHAR(100), 
  [Tel3_1] NVARCHAR(15), 
  [Tel3_2] NVARCHAR(15), 
  [Fax3] NVARCHAR(15), 
  [Fornece4] INT DEFAULT 0, 
  [Nome4] NVARCHAR(100), 
  [Tel4_1] NVARCHAR(15), 
  [Tel4_2] NVARCHAR(15), 
  [Fax4] NVARCHAR(15), 
  [Fornece5] INT DEFAULT 0, 
  [Nome5] NVARCHAR(100), 
  [Tel5_1] NVARCHAR(15), 
  [Tel5_2] NVARCHAR(15), 
  [Fax5] NVARCHAR(15), 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código], [Tamanho], [Cor], [Edição])
)

--
-- Table structure for table 'Contagem'
--

IF object_id(N'Contagem', 'U') IS NOT NULL DROP TABLE [Contagem]

CREATE TABLE [Contagem] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Código Ordenação] NVARCHAR(20), 
  [Nome] NVARCHAR(100), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(25), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(25), 
  [Unidade] NVARCHAR(5), 
  [Qtde Estoque] REAL DEFAULT 0, 
  [Fracionado] BIT, 
  [Digitado] REAL DEFAULT 0, 
  [Diferença] REAL DEFAULT 0, 
  [Consertar] BIT, 
  [Empresa] INT NOT NULL DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código], [Empresa])
)

--
-- Table structure for table 'Contagem Grade'
--

IF object_id(N'Contagem Grade', 'U') IS NOT NULL DROP TABLE [Contagem Grade]

CREATE TABLE [Contagem Grade] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Código Ordenação] NVARCHAR(20), 
  [Nome] NVARCHAR(100), 
  [Nome Tamanho] NVARCHAR(30), 
  [Nome Cor] NVARCHAR(30), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(25), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(25), 
  [Unidade] NVARCHAR(5), 
  [Qtde Estoque] REAL DEFAULT 0, 
  [Digitado] REAL DEFAULT 0, 
  [Diferença] REAL DEFAULT 0, 
  [Consertar] BIT, 
  [Empresa] INT NOT NULL DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código], [Tamanho], [Cor], [Empresa])
)

--
-- Table structure for table 'ContEstoqueBalanco'
--

IF object_id(N'ContEstoqueBalanco', 'U') IS NOT NULL DROP TABLE [ContEstoqueBalanco]

CREATE TABLE [ContEstoqueBalanco] (
  [Produto] NVARCHAR(20), 
  [Descricao] NVARCHAR(60), 
  [Quantidade] REAL DEFAULT 0
)

--
-- Table structure for table 'ControlVendas'
--

IF object_id(N'ControlVendas', 'U') IS NOT NULL DROP TABLE [ControlVendas]

CREATE TABLE [ControlVendas] (
  [Codigo] INT DEFAULT 0, 
  [Nome] NVARCHAR(35), 
  [QtdeOpera] INT DEFAULT 0, 
  [QtdeItens] INT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Comissao] FLOAT DEFAULT 0
)

--
-- Table structure for table 'DataCurva'
--

IF object_id(N'DataCurva', 'U') IS NOT NULL DROP TABLE [DataCurva]

CREATE TABLE [DataCurva] (
  [Ultima] DATETIME2
)

--
-- Table structure for table 'Entradas'
--

IF object_id(N'Entradas', 'U') IS NOT NULL DROP TABLE [Entradas]

CREATE TABLE [Entradas] (
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Contador] INT NOT NULL IDENTITY, 
  [Data] DATETIME2, 
  [Cód Operação] INT DEFAULT 0, 
  [Nome Operação] NVARCHAR(50), 
  [Cód Digitador] INT DEFAULT 0, 
  [Nome Digitador] NVARCHAR(35), 
  [Cód Fornecedor] INT DEFAULT 0, 
  [Nome Fornecedor] NVARCHAR(100), 
  [Observações] NVARCHAR(70), 
  [Nota Fiscal] NVARCHAR(15), 
  [Pedido] NVARCHAR(15), 
  [Data Emissão] DATETIME2, 
  [Efetivada] BIT, 
  [Total Produtos] FLOAT DEFAULT 0, 
  [Total Desconto] FLOAT DEFAULT 0, 
  [Total IPI] FLOAT DEFAULT 0, 
  [Total Frete] FLOAT DEFAULT 0, 
  [Total B Icm] FLOAT DEFAULT 0, 
  [Total Icm] FLOAT DEFAULT 0, 
  [Total B Icm Subs] FLOAT DEFAULT 0, 
  [Total Icm Subs] FLOAT DEFAULT 0, 
  [Total Nota] FLOAT DEFAULT 0, 
  [Nota] INT DEFAULT 0, 
  [Código] NVARCHAR(100), 
  [Qtde] REAL DEFAULT 0, 
  [Nome] NVARCHAR(100), 
  [Preço] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [ICM] FLOAT DEFAULT 0, 
  [IPI] FLOAT DEFAULT 0, 
  [Preço Final] FLOAT DEFAULT 0, 
  [Etiqueta] BIT, 
  [Fracionado] BIT, 
  [Unidade Venda] NVARCHAR(5), 
  [CodUsuarioOwner] INT DEFAULT 0, 
  [CentroCusto] INT DEFAULT 0, 
  [NomeCentroCusto] NVARCHAR(50), 
  PRIMARY KEY ([Sequência], [Contador])
)

--
-- Table structure for table 'EntradasConsignadas'
--

IF object_id(N'EntradasConsignadas', 'U') IS NOT NULL DROP TABLE [EntradasConsignadas]

CREATE TABLE [EntradasConsignadas] (
  [Filial] SMALLINT DEFAULT 0, 
  [CodProduto] NVARCHAR(20), 
  [Sequencia] INT DEFAULT 0, 
  [Linha] SMALLINT DEFAULT 0, 
  [QtdeAtual] FLOAT DEFAULT 0, 
  [Custo] FLOAT DEFAULT 0, 
  [Fornecedor] INT DEFAULT 0, 
  [QtdeOriginal] FLOAT DEFAULT 0, 
  [PrecoFinal] FLOAT DEFAULT 0, 
  [EstoqueConsignado] FLOAT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0
)

--
-- Table structure for table 'Erros de conversão'
--

IF object_id(N'Erros de conversão', 'U') IS NOT NULL DROP TABLE [Erros de conversão]

CREATE TABLE [Erros de conversão] (
  [Tipo de objeto] NVARCHAR(255), 
  [Nome do objeto] NVARCHAR(255), 
  [Descrição de erro] NVARCHAR(MAX)
)

--
-- Table structure for table 'estColNaoImportados'
--

IF object_id(N'estColNaoImportados', 'U') IS NOT NULL DROP TABLE [estColNaoImportados]

CREATE TABLE [estColNaoImportados] (
  [proID] NVARCHAR(50), 
  [proCor] INT DEFAULT 0, 
  [proTamanho] INT DEFAULT 0, 
  [proEdicao] INT DEFAULT 0, 
  [proQtde] FLOAT DEFAULT 0
)

--
-- Table structure for table 'EstoqueTemp'
--

IF object_id(N'EstoqueTemp', 'U') IS NOT NULL DROP TABLE [EstoqueTemp]

CREATE TABLE [EstoqueTemp] (
  [Produto] NVARCHAR(20), 
  [Final] REAL DEFAULT 0
)

--
-- Table structure for table 'Extrato'
--

IF object_id(N'Extrato', 'U') IS NOT NULL DROP TABLE [Extrato]

CREATE TABLE [Extrato] (
  [Sequencia] INT DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [NomeProduto] NVARCHAR(100), 
  [Tam] NVARCHAR(3), 
  [Cor] NVARCHAR(3), 
  [Data] DATETIME2, 
  [ValorUnitario] FLOAT DEFAULT 0, 
  [Saldo] FLOAT DEFAULT 0
)

--
-- Table structure for table 'ExtratoGroup'
--

IF object_id(N'ExtratoGroup', 'U') IS NOT NULL DROP TABLE [ExtratoGroup]

CREATE TABLE [ExtratoGroup] (
  [CodigoMovi] INT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [NomeProduto] NVARCHAR(100), 
  [Tam] NVARCHAR(3), 
  [Cor] NVARCHAR(3), 
  PRIMARY KEY ([CodigoMovi])
)

--
-- Table structure for table 'ExtratoSeq'
--

IF object_id(N'ExtratoSeq', 'U') IS NOT NULL DROP TABLE [ExtratoSeq]

CREATE TABLE [ExtratoSeq] (
  [CodigoMovi] INT NOT NULL DEFAULT 0, 
  [Sequencia] INT NOT NULL DEFAULT 0, 
  [Saldo] FLOAT DEFAULT 0, 
  [ValorUnitario] FLOAT DEFAULT 0, 
  PRIMARY KEY ([CodigoMovi], [Sequencia])
)

--
-- Table structure for table 'Fluxo'
--

IF object_id(N'Fluxo', 'U') IS NOT NULL DROP TABLE [Fluxo]

CREATE TABLE [Fluxo] (
  [Data] DATETIME2 NOT NULL, 
  [Ordem] INT NOT NULL DEFAULT 0, 
  [Saldo Anterior] MONEY DEFAULT 0, 
  [Cód Entrada] INT DEFAULT 0, 
  [Desc Entrada] NVARCHAR(50), 
  [Valor Entrada] MONEY DEFAULT 0, 
  [Cód Saída] INT DEFAULT 0, 
  [Desc Saída] NVARCHAR(50), 
  [Valor Saída] MONEY DEFAULT 0, 
  [Saldo Final] MONEY DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Data], [Ordem])
)

--
-- Table structure for table 'FolhaPagamento'
--

IF object_id(N'FolhaPagamento', 'U') IS NOT NULL DROP TABLE [FolhaPagamento]

CREATE TABLE [FolhaPagamento] (
  [Codigo] INT DEFAULT 0, 
  [Parcela] INT DEFAULT 0, 
  [Total] FLOAT DEFAULT 0
)

--
-- Table structure for table 'Gráfico3'
--

IF object_id(N'Gráfico3', 'U') IS NOT NULL DROP TABLE [Gráfico3]

CREATE TABLE [Gráfico3] (
  [Ano] INT NOT NULL DEFAULT 0, 
  [Mês] SMALLINT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(20), 
  [Unidades Vendidas] INT DEFAULT 0, 
  [Valor Vendas] FLOAT DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Ano], [Mês])
)

--
-- Table structure for table 'Inventário'
--

IF object_id(N'Inventário', 'U') IS NOT NULL DROP TABLE [Inventário]

CREATE TABLE [Inventário] (
  [Empresa] SMALLINT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Ordenação] NVARCHAR(20), 
  [Nome] NVARCHAR(100), 
  [Unidade Venda] NVARCHAR(5), 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Nome Tamanho] NVARCHAR(30), 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Nome Cor] NVARCHAR(30), 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Nome Edição] NVARCHAR(30), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(30), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(30), 
  [Estoque Final] REAL DEFAULT 0, 
  [Data] DATETIME2, 
  [Preço] REAL DEFAULT 0, 
  [Fracionado] BIT, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  [Classificação Fiscal] NVARCHAR(15), 
  PRIMARY KEY ([Empresa], [Produto], [Tamanho], [Cor], [Edição])
)

--
-- Table structure for table 'Lucratividade'
--

IF object_id(N'Lucratividade', 'U') IS NOT NULL DROP TABLE [Lucratividade]

CREATE TABLE [Lucratividade] (
  [Produto] NVARCHAR(20), 
  [Grupo] NVARCHAR(50), 
  [Código Ordenação] NVARCHAR(20), 
  [Nome] NVARCHAR(100), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(30), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(30), 
  [Vendedor] INT DEFAULT 0, 
  [Nome Vendedor] NVARCHAR(35), 
  [Qtde] REAL DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Custo] FLOAT DEFAULT 0, 
  [Lucro] FLOAT DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0
)

--
-- Table structure for table 'MalaExportacao'
--

IF object_id(N'MalaExportacao', 'U') IS NOT NULL DROP TABLE [MalaExportacao]

CREATE TABLE [MalaExportacao] (
  [Codigo] INT DEFAULT 0, 
  [Nome] NVARCHAR(70), 
  [Endereco] NVARCHAR(100), 
  [Complemento] NVARCHAR(20), 
  [Bairro] NVARCHAR(30), 
  [CEP] NVARCHAR(9), 
  [Cidade] NVARCHAR(40), 
  [Estado] NVARCHAR(2), 
  [Nascimento] DATETIME2, 
  [DataIncorreta] BIT, 
  [Grupo] SMALLINT DEFAULT 0
)

--
-- Table structure for table 'ParametrosTMP'
--

IF object_id(N'ParametrosTMP', 'U') IS NOT NULL DROP TABLE [ParametrosTMP]

CREATE TABLE [ParametrosTMP] (
  [Filial] SMALLINT NOT NULL, 
  [InRelZebrados] BIT, 
  PRIMARY KEY ([Filial])
)

--
-- Table structure for table 'Preço Custo'
--

IF object_id(N'Preço Custo', 'U') IS NOT NULL DROP TABLE [Preço Custo]

CREATE TABLE [Preço Custo] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Nome] NVARCHAR(100), 
  [Preço Custo Anterior] REAL DEFAULT 0, 
  [Preço Custo Atual] REAL DEFAULT 0, 
  [Preço Custo Calc Ant] REAL DEFAULT 0, 
  [Preço Custo Calc Atu] REAL DEFAULT 0, 
  [Preço Venda Anterior] REAL DEFAULT 0, 
  [Lucro Anterior] REAL DEFAULT 0, 
  [Lucro Anterior Perc] REAL DEFAULT 0, 
  [Preço Venda Atual] REAL DEFAULT 0, 
  [Lucro Atual] REAL DEFAULT 0, 
  [Lucro Atual Perc] REAL DEFAULT 0, 
  [Alterar] BIT, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'PrestacaoContas'
--

IF object_id(N'PrestacaoContas', 'U') IS NOT NULL DROP TABLE [PrestacaoContas]

CREATE TABLE [PrestacaoContas] (
  [Filial] SMALLINT DEFAULT 0, 
  [Fornecedor] INT DEFAULT 0, 
  [Sequencia] INT DEFAULT 0, 
  [Linha] SMALLINT DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [Custo] FLOAT DEFAULT 0, 
  [QtdeOriginal] INT DEFAULT 0, 
  [QtdeDevolvida] FLOAT DEFAULT 0, 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [QtdeComprada] FLOAT DEFAULT 0, 
  [DatadaGeracao] DATETIME2, 
  [Finalizado] BIT, 
  [DatadaFinalizacao] DATETIME2, 
  [ImpressoNF] BIT, 
  [Resultado] SMALLINT DEFAULT 0, 
  [PrestacaoFechada] BIT, 
  [CompraFechada] BIT, 
  [PeriodoVenda] DATETIME2, 
  [NotaFiscal] INT DEFAULT 0, 
  [QtdeAcertada] FLOAT DEFAULT 0
)

--
-- Table structure for table 'PrestacaoContasTemp'
--

IF object_id(N'PrestacaoContasTemp', 'U') IS NOT NULL DROP TABLE [PrestacaoContasTemp]

CREATE TABLE [PrestacaoContasTemp] (
  [Filial] SMALLINT DEFAULT 0, 
  [Fornecedor] INT DEFAULT 0, 
  [Sequencia] INT DEFAULT 0, 
  [Linha] SMALLINT DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [Custo] FLOAT DEFAULT 0, 
  [QtdeOriginal] FLOAT DEFAULT 0, 
  [QtdeDevolvida] FLOAT DEFAULT 0, 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [QtdeComprada] FLOAT DEFAULT 0, 
  [DatadaGeracao] DATETIME2, 
  [Finalizado] BIT, 
  [DatadaFinalizacao] DATETIME2, 
  [ImpressoNF] BIT, 
  [Resultado] SMALLINT DEFAULT 0, 
  [PrestacaoFechada] BIT, 
  [CompraFechada] BIT, 
  [PeriodoVenda] DATETIME2, 
  [NotaFiscal] INT DEFAULT 0, 
  [QtdeAcertada] FLOAT DEFAULT 0
)

--
-- Table structure for table 'Rel Grade'
--

IF object_id(N'Rel Grade', 'U') IS NOT NULL DROP TABLE [Rel Grade]

CREATE TABLE [Rel Grade] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Nome] NVARCHAR(100), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(30), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(30), 
  [Cód Cor] INT NOT NULL DEFAULT 0, 
  [Nome Cor] NVARCHAR(50), 
  [Tamanho1] INT DEFAULT 0, 
  [Tamanho2] INT DEFAULT 0, 
  [Tamanho3] INT DEFAULT 0, 
  [Tamanho4] INT DEFAULT 0, 
  [Tamanho5] INT DEFAULT 0, 
  [Tamanho6] INT DEFAULT 0, 
  [Tamanho7] INT DEFAULT 0, 
  [Tamanho8] INT DEFAULT 0, 
  [Tamanho9] INT DEFAULT 0, 
  [Tamanho10] INT DEFAULT 0, 
  [Tamanho11] INT DEFAULT 0, 
  [Tamanho12] INT DEFAULT 0, 
  [Tamanho13] INT DEFAULT 0, 
  [Tamanho14] INT DEFAULT 0, 
  [Tamanho15] INT DEFAULT 0, 
  [Outros] INT DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código], [Cód Cor])
)

--
-- Table structure for table 'Rel Grade2'
--

IF object_id(N'Rel Grade2', 'U') IS NOT NULL DROP TABLE [Rel Grade2]

CREATE TABLE [Rel Grade2] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Nome] NVARCHAR(100), 
  [Classe] INT DEFAULT 0, 
  [Nome Classe] NVARCHAR(30), 
  [Sub Classe] INT DEFAULT 0, 
  [Nome Sub] NVARCHAR(30), 
  [Cód Cor] INT NOT NULL DEFAULT 0, 
  [Nome Cor] NVARCHAR(50), 
  [Tamanho1] INT DEFAULT 0, 
  [Tamanho2] INT DEFAULT 0, 
  [Tamanho3] INT DEFAULT 0, 
  [Tamanho4] INT DEFAULT 0, 
  [Tamanho5] INT DEFAULT 0, 
  [Tamanho6] INT DEFAULT 0, 
  [Tamanho7] INT DEFAULT 0, 
  [Tamanho8] INT DEFAULT 0, 
  [Tamanho9] INT DEFAULT 0, 
  [Tamanho10] INT DEFAULT 0, 
  [Tamanho11] INT DEFAULT 0, 
  [Tamanho12] INT DEFAULT 0, 
  [Tamanho13] INT DEFAULT 0, 
  [Tamanho14] INT DEFAULT 0, 
  [Tamanho15] INT DEFAULT 0, 
  [Outros] INT DEFAULT 0, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  PRIMARY KEY ([Código], [Cód Cor])
)

--
-- Table structure for table 'RelProdutosComprar'
--

IF object_id(N'RelProdutosComprar', 'U') IS NOT NULL DROP TABLE [RelProdutosComprar]

CREATE TABLE [RelProdutosComprar] (
  [CodigoProduto] NVARCHAR(20), 
  [Descricao] NVARCHAR(80), 
  [CodigoClasse] INT DEFAULT 0, 
  [Classe] NVARCHAR(25), 
  [CodigoSubClasse] INT DEFAULT 0, 
  [SubClasse] NVARCHAR(25), 
  [CodigoFornecedor] INT DEFAULT 0, 
  [Fornecedor] NVARCHAR(60), 
  [EstoqueFisicoAtual] FLOAT DEFAULT 0, 
  [EstoqueFisicoTotal] FLOAT DEFAULT 0, 
  [EspacoFisicoDisponivel] FLOAT DEFAULT 0, 
  [MediaVendas] FLOAT DEFAULT 0, 
  [QuantidadeComprar] FLOAT DEFAULT 0
)

--
-- Table structure for table 'RelVendaComissao'
--

IF object_id(N'RelVendaComissao', 'U') IS NOT NULL DROP TABLE [RelVendaComissao]

CREATE TABLE [RelVendaComissao] (
  [CodigoVendedor] INT DEFAULT 0, 
  [NomeVendedor] NVARCHAR(35), 
  [Comissao] REAL DEFAULT 0, 
  [CodigoCliente] INT DEFAULT 0, 
  [NomeCliente] NVARCHAR(60), 
  [Data] DATETIME2, 
  [Sequencia] INT DEFAULT 0, 
  [NotaFiscal] INT DEFAULT 0, 
  [Custo] REAL DEFAULT 0, 
  [ValorFinal] REAL DEFAULT 0, 
  [Lucro] REAL DEFAULT 0, 
  [Indice] REAL DEFAULT 0, 
  [ComissaoValor] REAL DEFAULT 0, 
  [Tipo] NVARCHAR(1), 
  [NomeOperacao] NVARCHAR(50)
)

--
-- Table structure for table 'RelVendasFornecedor'
--

IF object_id(N'RelVendasFornecedor', 'U') IS NOT NULL DROP TABLE [RelVendasFornecedor]

CREATE TABLE [RelVendasFornecedor] (
  [FornecedorCodigo] INT DEFAULT 0, 
  [FornecedorNome] NVARCHAR(100), 
  [ClienteCodigo] INT DEFAULT 0, 
  [ClienteNome] NVARCHAR(100), 
  [ClienteCNPJCPF] NVARCHAR(50), 
  [ClienteCidade] NVARCHAR(50), 
  [ClienteEstado] NVARCHAR(50), 
  [ProdutoCodigo] NVARCHAR(20), 
  [ProdutoNome] NVARCHAR(80), 
  [ProdutoQuantidade] FLOAT DEFAULT 0, 
  [ProdutoValor] FLOAT DEFAULT 0, 
  [ProdutoClasseCodigo] INT DEFAULT 0, 
  [ProdutoClasseNome] NVARCHAR(25), 
  [ProdutoSubClasseCodigo] INT DEFAULT 0, 
  [ProdutoSubClasseNome] NVARCHAR(25)
)

--
-- Table structure for table 'Reports'
--

IF object_id(N'Reports', 'U') IS NOT NULL DROP TABLE [Reports]

CREATE TABLE [Reports] (
  [InRelZebrados] BIT, 
  [nColorRed] SMALLINT, 
  [nColorGreen] SMALLINT, 
  [nColorBlue] SMALLINT
)

--
-- Table structure for table 'Saídas'
--

IF object_id(N'Saídas', 'U') IS NOT NULL DROP TABLE [Saídas]

CREATE TABLE [Saídas] (
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Contador] INT NOT NULL IDENTITY, 
  [Data] DATETIME2, 
  [Cód Operação] INT DEFAULT 0, 
  [Nome Operação] NVARCHAR(50), 
  [Tabela] NVARCHAR(15), 
  [Cód Digitador] INT DEFAULT 0, 
  [Nome Digitador] NVARCHAR(35), 
  [Cód Cliente] INT DEFAULT 0, 
  [Nome Cliente] NVARCHAR(100), 
  [Observações] NVARCHAR(70), 
  [Ref Interna] NVARCHAR(10), 
  [Efetivada] BIT, 
  [Total Produtos] FLOAT DEFAULT 0, 
  [Total Desconto] FLOAT DEFAULT 0, 
  [Total IPI] FLOAT DEFAULT 0, 
  [Total Frete] FLOAT DEFAULT 0, 
  [Total B Icm] FLOAT DEFAULT 0, 
  [Total Icm] FLOAT DEFAULT 0, 
  [Total B Icm Subs] FLOAT DEFAULT 0, 
  [Total Icm Subs] FLOAT DEFAULT 0, 
  [Total Nota] FLOAT DEFAULT 0, 
  [Total Serviços] FLOAT DEFAULT 0, 
  [Total ISS] FLOAT DEFAULT 0, 
  [Nota] INT DEFAULT 0, 
  [Conta] BIT, 
  [Dinheiro] FLOAT DEFAULT 0, 
  [Cartão] FLOAT DEFAULT 0, 
  [Vale] FLOAT DEFAULT 0, 
  [Tipo Prod] NVARCHAR(1), 
  [Código] NVARCHAR(20), 
  [Qtde] REAL DEFAULT 0, 
  [Nome] NVARCHAR(100), 
  [Preço] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [ICM] FLOAT DEFAULT 0, 
  [IPI] FLOAT DEFAULT 0, 
  [Preço Final] FLOAT DEFAULT 0, 
  [Etiqueta] BIT, 
  [Fracionado] BIT, 
  [CodUsuarioOwner] INT DEFAULT 0, 
  [Nota Cancelada] BIT, 
  [DescontoSubTotal] MONEY DEFAULT 0, 
  PRIMARY KEY ([Sequência], [Contador])
)

--
-- Table structure for table 'tblRelCertificados'
--

IF object_id(N'tblRelCertificados', 'U') IS NOT NULL DROP TABLE [tblRelCertificados]

CREATE TABLE [tblRelCertificados] (
  [Nome do Produto] NVARCHAR(80), 
  [Obs do Produto] NVARCHAR(MAX), 
  [Nome do Cliente] NVARCHAR(100), 
  [Data da Saída] DATETIME2, 
  [Código do Produto] NVARCHAR(20), 
  [Nota Fiscal] INT DEFAULT 0, 
  [Contador] INT DEFAULT 0
)

--
-- Table structure for table 'tblRelCompras'
--

IF object_id(N'tblRelCompras', 'U') IS NOT NULL DROP TABLE [tblRelCompras]

CREATE TABLE [tblRelCompras] (
  [filID] INT DEFAULT 0, 
  [proID] NVARCHAR(50), 
  [proTipo] NVARCHAR(1), 
  [tamID] INT DEFAULT 0, 
  [corID] INT DEFAULT 0, 
  [ediID] INT DEFAULT 0, 
  [comData] DATETIME2, 
  [comQtde] FLOAT DEFAULT 0, 
  [comValor] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelEstoquePrecos'
--

IF object_id(N'tblRelEstoquePrecos', 'U') IS NOT NULL DROP TABLE [tblRelEstoquePrecos]

CREATE TABLE [tblRelEstoquePrecos] (
  [codigo] NVARCHAR(20) NOT NULL, 
  [codigo_ordenacao] NVARCHAR(20) NOT NULL, 
  [nome] NVARCHAR(80) NOT NULL, 
  [codigo_fornecedor] NVARCHAR(15), 
  [qtde_filial_1] REAL DEFAULT 0, 
  [qtde_filial_2] REAL DEFAULT 0, 
  [qtde_filial_3] REAL DEFAULT 0, 
  [preco_1] REAL DEFAULT 0, 
  [preco_2] REAL DEFAULT 0, 
  [classe] INT DEFAULT 0, 
  [subclasse] INT DEFAULT 0, 
  PRIMARY KEY ([codigo])
)

--
-- Table structure for table 'tblRelFornecedor'
--

IF object_id(N'tblRelFornecedor', 'U') IS NOT NULL DROP TABLE [tblRelFornecedor]

CREATE TABLE [tblRelFornecedor] (
  [Fornecedor] INT DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [QtdeEntrada] FLOAT DEFAULT 0, 
  [PrecoCusto] FLOAT DEFAULT 0, 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [EstoqueAtual] FLOAT DEFAULT 0, 
  [PrestacaoContas] FLOAT DEFAULT 0, 
  [Venda] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelFornecedorTemp'
--

IF object_id(N'tblRelFornecedorTemp', 'U') IS NOT NULL DROP TABLE [tblRelFornecedorTemp]

CREATE TABLE [tblRelFornecedorTemp] (
  [Fornecedor] INT DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [QtdeEntrada] FLOAT DEFAULT 0, 
  [PrecoCusto] FLOAT DEFAULT 0, 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [EstoqueAtual] FLOAT DEFAULT 0, 
  [PrestacaoContas] FLOAT DEFAULT 0, 
  [Venda] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelMalaAutorizacoes'
--

IF object_id(N'tblRelMalaAutorizacoes', 'U') IS NOT NULL DROP TABLE [tblRelMalaAutorizacoes]

CREATE TABLE [tblRelMalaAutorizacoes] (
  [Nome] NVARCHAR(70), 
  [Endereco] NVARCHAR(100), 
  [Bairro] NVARCHAR(30), 
  [Complemento] NVARCHAR(20), 
  [Cidade] NVARCHAR(40), 
  [Estado] NVARCHAR(2), 
  [Cep] NVARCHAR(9), 
  [Codigo] INT DEFAULT 0
)

--
-- Table structure for table 'tblRelMalaAutorizacoesFatur'
--

IF object_id(N'tblRelMalaAutorizacoesFatur', 'U') IS NOT NULL DROP TABLE [tblRelMalaAutorizacoesFatur]

CREATE TABLE [tblRelMalaAutorizacoesFatur] (
  [Nome] NVARCHAR(70), 
  [Endereco] NVARCHAR(100), 
  [Bairro] NVARCHAR(30), 
  [Complemento] NVARCHAR(20), 
  [Cidade] NVARCHAR(40), 
  [Estado] NVARCHAR(2), 
  [Cep] NVARCHAR(9), 
  [Codigo] INT DEFAULT 0
)

--
-- Table structure for table 'tblRelPosFin1'
--

IF object_id(N'tblRelPosFin1', 'U') IS NOT NULL DROP TABLE [tblRelPosFin1]

CREATE TABLE [tblRelPosFin1] (
  [CodProduto] NVARCHAR(20), 
  [NomeProduto] NVARCHAR(80), 
  [NF] INT DEFAULT 0, 
  [Seq] INT DEFAULT 0
)

--
-- Table structure for table 'tblRelPosFin2'
--

IF object_id(N'tblRelPosFin2', 'U') IS NOT NULL DROP TABLE [tblRelPosFin2]

CREATE TABLE [tblRelPosFin2] (
  [NF] INT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Vencimento] DATETIME2, 
  [Seq] INT DEFAULT 0, 
  [Fatura] NVARCHAR(10)
)

--
-- Table structure for table 'tblRelRecebFormaPgto'
--

IF object_id(N'tblRelRecebFormaPgto', 'U') IS NOT NULL DROP TABLE [tblRelRecebFormaPgto]

CREATE TABLE [tblRelRecebFormaPgto] (
  [Owner] INT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [ContaCliente] FLOAT DEFAULT 0, 
  [Dinheiro] FLOAT DEFAULT 0, 
  [Cartao] FLOAT DEFAULT 0, 
  [ValesOutros] FLOAT DEFAULT 0, 
  [Cheque] REAL DEFAULT 0, 
  [ChequePre] FLOAT DEFAULT 0, 
  [Parcelamento] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Owner], [Data])
)

--
-- Table structure for table 'tblRelVendas'
--

IF object_id(N'tblRelVendas', 'U') IS NOT NULL DROP TABLE [tblRelVendas]

CREATE TABLE [tblRelVendas] (
  [filID] INT DEFAULT 0, 
  [proID] NVARCHAR(50), 
  [proTipo] NVARCHAR(1), 
  [tamID] INT DEFAULT 0, 
  [corID] INT DEFAULT 0, 
  [ediID] INT DEFAULT 0, 
  [venData] DATETIME2, 
  [venQtde] FLOAT DEFAULT 0, 
  [venValor] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelVendasCliente'
--

IF object_id(N'tblRelVendasCliente', 'U') IS NOT NULL DROP TABLE [tblRelVendasCliente]

CREATE TABLE [tblRelVendasCliente] (
  [Cliente] INT DEFAULT 0, 
  [Filial] SMALLINT DEFAULT 0, 
  [Data] DATETIME2, 
  [Produto] NVARCHAR(50), 
  [Tamanho] INT DEFAULT 0, 
  [Cor] INT DEFAULT 0, 
  [Edicao] INT DEFAULT 0, 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [ValorVendido] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelVendasCliente2'
--

IF object_id(N'tblRelVendasCliente2', 'U') IS NOT NULL DROP TABLE [tblRelVendasCliente2]

CREATE TABLE [tblRelVendasCliente2] (
  [Cliente] INT DEFAULT 0, 
  [Filial] SMALLINT DEFAULT 0, 
  [Data] DATETIME2, 
  [Produto] NVARCHAR(50), 
  [Tamanho] INT DEFAULT 0, 
  [Cor] INT DEFAULT 0, 
  [Edicao] INT DEFAULT 0, 
  [QtdeVendida] FLOAT DEFAULT 0, 
  [ValorVendido] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelVendasDescontoSubTotal'
--

IF object_id(N'tblRelVendasDescontoSubTotal', 'U') IS NOT NULL DROP TABLE [tblRelVendasDescontoSubTotal]

CREATE TABLE [tblRelVendasDescontoSubTotal] (
  [filID] INT NOT NULL DEFAULT 0, 
  [movSequencia] INT NOT NULL DEFAULT 0, 
  [movValorDesconto] FLOAT DEFAULT 0, 
  PRIMARY KEY ([filID], [movSequencia])
)

--
-- Table structure for table 'tblRelVendasGroup'
--

IF object_id(N'tblRelVendasGroup', 'U') IS NOT NULL DROP TABLE [tblRelVendasGroup]

CREATE TABLE [tblRelVendasGroup] (
  [Filial] SMALLINT DEFAULT 0, 
  [Produto] NVARCHAR(20), 
  [Qtde] FLOAT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelVendasII'
--

IF object_id(N'tblRelVendasII', 'U') IS NOT NULL DROP TABLE [tblRelVendasII]

CREATE TABLE [tblRelVendasII] (
  [Codigo] NVARCHAR(20), 
  [Descricao] NVARCHAR(80), 
  [ICMS] REAL, 
  [PrecoCusto] FLOAT, 
  [ValorVenda] FLOAT, 
  [Quantidade] FLOAT, 
  [ValorICMSVenda] FLOAT DEFAULT 0
)

--
-- Table structure for table 'tblRelVendasPorVendedor'
--

IF object_id(N'tblRelVendasPorVendedor', 'U') IS NOT NULL DROP TABLE [tblRelVendasPorVendedor]

CREATE TABLE [tblRelVendasPorVendedor] (
  [Filial] INT, 
  [Vendedor] INT, 
  [DataIni1] DATETIME2, 
  [DataFim1] DATETIME2, 
  [DataIni2] DATETIME2, 
  [DataFim2] DATETIME2, 
  [DataIni3] DATETIME2, 
  [DataFim3] DATETIME2, 
  [Operacao] INT, 
  [Cliente] INT, 
  [SumMes1] FLOAT, 
  [SumMes2] FLOAT, 
  [SumMes3] FLOAT, 
  [SumMeses] FLOAT
)

--
-- Table structure for table 'tblRelVendasTamanho'
--

IF object_id(N'tblRelVendasTamanho', 'U') IS NOT NULL DROP TABLE [tblRelVendasTamanho]

CREATE TABLE [tblRelVendasTamanho] (
  [Filial] SMALLINT DEFAULT 0, 
  [Data] DATETIME2, 
  [Produto] NVARCHAR(20), 
  [Tamanho] INT DEFAULT 0, 
  [Edicao] INT DEFAULT 0, 
  [Qtde Vendida] FLOAT DEFAULT 0, 
  [Valor Vendido] FLOAT DEFAULT 0, 
  [Tipo] NVARCHAR(1), 
  [Cor] INT DEFAULT 0
)

--
-- Table structure for table 'tbRelEstoqueFiliais_Campos'
--

IF object_id(N'tbRelEstoqueFiliais_Campos', 'U') IS NOT NULL DROP TABLE [tbRelEstoqueFiliais_Campos]

CREATE TABLE [tbRelEstoqueFiliais_Campos] (
  [proCodigo] NVARCHAR(50), 
  [proCor] INT DEFAULT 0, 
  [proTamanho] INT DEFAULT 0, 
  [proEdicao] INT DEFAULT 0, 
  [est_1] FLOAT DEFAULT 0, 
  [est_2] FLOAT DEFAULT 0, 
  [est_3] FLOAT DEFAULT 0, 
  [est_4] FLOAT DEFAULT 0, 
  [est_5] FLOAT DEFAULT 0, 
  [est_6] FLOAT DEFAULT 0, 
  [est_7] FLOAT DEFAULT 0, 
  [est_8] FLOAT DEFAULT 0, 
  [est_9] FLOAT DEFAULT 0, 
  [est_10] FLOAT DEFAULT 0, 
  [est_Total] FLOAT DEFAULT 0
)

--
-- Table structure for table 'TotalCartoes'
--

IF object_id(N'TotalCartoes', 'U') IS NOT NULL DROP TABLE [TotalCartoes]

CREATE TABLE [TotalCartoes] (
  [Sequencia] INT DEFAULT 0, 
  [Administradora] SMALLINT DEFAULT 0, 
  [Vl_Bruto] FLOAT DEFAULT 0, 
  [Vl_Liquido] FLOAT DEFAULT 0, 
  [Filial] INT DEFAULT 0
)

--
-- Table structure for table 'TotalCartoesGroup'
--

IF object_id(N'TotalCartoesGroup', 'U') IS NOT NULL DROP TABLE [TotalCartoesGroup]

CREATE TABLE [TotalCartoesGroup] (
  [Administradora] SMALLINT DEFAULT 0, 
  [Nome] NVARCHAR(25), 
  [Vl_Bruto] FLOAT DEFAULT 0, 
  [Vl_Liquido] FLOAT DEFAULT 0, 
  [Filial] INT DEFAULT 0
)

