--
-- DUMP FILE
--
-- Database is ported from MS Access
--------------------------------------------------------------------
-- Created using "MS Access to MSSQL" form http://www.bullzip.com
-- Program Version 5.5.281
--
-- OPTIONS:
--   sourcefilename=C:/Projetos/QuickStore/0_PastaZero_Legado/QuickStore2003/BancoDados/QuickStore.mdb
--   sourceusername=
--   sourcepassword=
--   sourcesystemdatabase=
--   destinationserver=CAGEPAR-000261\SQLEXPRESS
--   destinationauthentication=SQL
--   destinationdatabase=QuickStore
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
IF EXISTS (SELECT * FROM master.dbo.sysdatabases WHERE name = N'QuickStore') ALTER DATABASE [QuickStore] SET SINGLE_USER With ROLLBACK IMMEDIATE
IF EXISTS (SELECT * FROM master.dbo.sysdatabases WHERE name = N'QuickStore') DROP DATABASE [QuickStore]
IF NOT EXISTS (SELECT * FROM master.dbo.sysdatabases WHERE name = N'QuickStore') CREATE DATABASE [QuickStore]
USE [QuickStore]

--
-- Table structure for table 'AcertoConsignacaoEntrada'
--

IF object_id(N'AcertoConsignacaoEntrada', 'U') IS NOT NULL DROP TABLE [AcertoConsignacaoEntrada]

CREATE TABLE [AcertoConsignacaoEntrada] (
  [Filial] SMALLINT, 
  [Sequencia] INT, 
  [DataAcerto] DATETIME2, 
  [LinhaProduto] INT, 
  [CodigoProduto] NVARCHAR(100), 
  [QtdeVendida] FLOAT, 
  [FilialVenda] SMALLINT, 
  [SequenciaVenda] INT, 
  [PrecoCusto] FLOAT DEFAULT 0, 
  [PrecoVenda] FLOAT DEFAULT 0
)

--
-- Table structure for table 'Acessos'
--

IF object_id(N'Acessos', 'U') IS NOT NULL DROP TABLE [Acessos]

CREATE TABLE [Acessos] (
  [Programa] NVARCHAR(40), 
  [Usuário] INT DEFAULT 0, 
  [Gravar] BIT, 
  [Apagar] BIT, 
  [Numero] INT
)

--
-- Table structure for table 'AcessoTabelasDePrecosProdutos'
--

IF object_id(N'AcessoTabelasDePrecosProdutos', 'U') IS NOT NULL DROP TABLE [AcessoTabelasDePrecosProdutos]

CREATE TABLE [AcessoTabelasDePrecosProdutos] (
  [Usuario] INT, 
  [Tabela] NVARCHAR(15)
)

--
-- Table structure for table 'Agenda'
--

IF object_id(N'Agenda', 'U') IS NOT NULL DROP TABLE [Agenda]

CREATE TABLE [Agenda] (
  [Funcionário] INT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL IDENTITY, 
  [Data Lembrar] DATETIME2, 
  [Data Digitação] DATETIME2, 
  [Lembrete] NVARCHAR(50), 
  [Lido] BIT, 
  PRIMARY KEY ([Funcionário], [Sequência])
)

--
-- Table structure for table 'AliquotasNCM'
--

IF object_id(N'AliquotasNCM', 'U') IS NOT NULL DROP TABLE [AliquotasNCM]

CREATE TABLE [AliquotasNCM] (
  [Codigo] NVARCHAR(13), 
  [EX] NVARCHAR(2), 
  [Tabela] INT, 
  [AliqNacional] FLOAT, 
  [AliqImportacao] FLOAT, 
  [Nome] NVARCHAR(100), 
  [CEST] NVARCHAR(10), 
  [TemFCP] BIT DEFAULT 0
)

--
-- Table structure for table 'Bancos'
--

IF object_id(N'Bancos', 'U') IS NOT NULL DROP TABLE [Bancos]

CREATE TABLE [Bancos] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(25), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'BooksVendidos'
--

IF object_id(N'BooksVendidos', 'U') IS NOT NULL DROP TABLE [BooksVendidos]

CREATE TABLE [BooksVendidos] (
  [Filial] SMALLINT, 
  [Sequencia] INT, 
  [Linha] SMALLINT
)

--
-- Table structure for table 'Caixa'
--

IF object_id(N'Caixa', 'U') IS NOT NULL DROP TABLE [Caixa]

CREATE TABLE [Caixa] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Caixa] SMALLINT NOT NULL DEFAULT 0, 
  [Ordem] INT NOT NULL DEFAULT 0, 
  [Funcionário] INT DEFAULT 0, 
  [Hora] NVARCHAR(8), 
  [Saldo Anterior] FLOAT DEFAULT 0, 
  [Dinheiro] FLOAT DEFAULT 0, 
  [Cheques] FLOAT DEFAULT 0, 
  [Cheques Pré] FLOAT DEFAULT 0, 
  [Cartões] FLOAT DEFAULT 0, 
  [Vales] FLOAT DEFAULT 0, 
  [Parcelamento] FLOAT DEFAULT 0, 
  [Total Dinheiro] FLOAT DEFAULT 0, 
  [Total Cheques] FLOAT DEFAULT 0, 
  [Total Cheques Pré] FLOAT DEFAULT 0, 
  [Total Cartões] FLOAT DEFAULT 0, 
  [Total Vales] FLOAT DEFAULT 0, 
  [Total Parcelamento] FLOAT DEFAULT 0, 
  [Final] FLOAT DEFAULT 0, 
  [Descrição] NVARCHAR(120), 
  PRIMARY KEY ([Filial], [Data], [Caixa], [Ordem])
)

--
-- Table structure for table 'Caixas em Uso'
--

IF object_id(N'Caixas em Uso', 'U') IS NOT NULL DROP TABLE [Caixas em Uso]

CREATE TABLE [Caixas em Uso] (
  [Caixa] SMALLINT NOT NULL DEFAULT 0, 
  [Descrição] NVARCHAR(50), 
  [Operador] INT DEFAULT 0, 
  PRIMARY KEY ([Caixa])
)

--
-- Table structure for table 'Cartões'
--

IF object_id(N'Cartões', 'U') IS NOT NULL DROP TABLE [Cartões]

CREATE TABLE [Cartões] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(25), 
  [Dias Pagar] SMALLINT DEFAULT 0, 
  [Taxa] REAL DEFAULT 0, 
  [TEF] BIT, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Centros de Custo'
--

IF object_id(N'Centros de Custo', 'U') IS NOT NULL DROP TABLE [Centros de Custo]

CREATE TABLE [Centros de Custo] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(40), 
  [Data Alteração] NVARCHAR(10), 
  [Ativo] BIT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Classes'
--

IF object_id(N'Classes', 'U') IS NOT NULL DROP TABLE [Classes]

CREATE TABLE [Classes] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(25), 
  [Data Alteração] NVARCHAR(10), 
  [LucroMinimoPermitido] FLOAT DEFAULT 0, 
  [IFX] INT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Classificação Fiscal'
--

IF object_id(N'Classificação Fiscal', 'U') IS NOT NULL DROP TABLE [Classificação Fiscal]

CREATE TABLE [Classificação Fiscal] (
  [Código] INT NOT NULL, 
  [Nome] NVARCHAR(15), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Cli_For'
--

IF object_id(N'Cli_For', 'U') IS NOT NULL DROP TABLE [Cli_For]

CREATE TABLE [Cli_For] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Tipo] NVARCHAR(1), 
  [Física_Jurídica] NVARCHAR(1), 
  [Faturado] BIT, 
  [Fantasia] NVARCHAR(60), 
  [Fax] NVARCHAR(15), 
  [Última Compra] DATETIME2, 
  [Bloqueado] BIT, 
  [Tem Conta] BIT, 
  [Desconto] REAL DEFAULT 0, 
  [Comentários] NVARCHAR(MAX), 
  [Transportadora] INT DEFAULT 0, 
  [Inativo] BIT, 
  [Endereço Cob] NVARCHAR(50), 
  [Complemento Cob] NVARCHAR(15), 
  [Bairro Cob] NVARCHAR(20), 
  [Cidade Cob] NVARCHAR(30), 
  [Estado Cob] NVARCHAR(2), 
  [CEP Cob] NVARCHAR(9), 
  [Contabilidade] INT DEFAULT 0, 
  [Informações Crédito] BIT, 
  [Sem Mala Direta] BIT, 
  [Vendedor] INT DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [Mensagem Boleto] NVARCHAR(50), 
  [Home Page] NVARCHAR(100), 
  [Conta Cobrança] SMALLINT DEFAULT 0, 
  [Limite Crédito] FLOAT DEFAULT 0, 
  [Prazo 1] INT DEFAULT 0, 
  [Prazo 2] INT DEFAULT 0, 
  [Prazo 3] INT DEFAULT 0, 
  [Alterar Prazo] BIT, 
  [CodTipoFrete] NVARCHAR(1), 
  [WebShopperID] NVARCHAR(32), 
  [WebDataCadastro] DATETIME2, 
  [WebCountry] NVARCHAR(50), 
  [DataNascimento] DATETIME2, 
  [WebBonus] INT, 
  [Nome] NVARCHAR(100), 
  [Estado] NVARCHAR(40), 
  [Sexo] NVARCHAR(1), 
  [email] NVARCHAR(100), 
  [CEP] NVARCHAR(15), 
  [Cidade] NVARCHAR(50), 
  [CGC] NVARCHAR(20), 
  [datAberturaCadastro] DATETIME2, 
  [DiaBaseConsignacao] SMALLINT, 
  [DataProxAcertoConsignacao] DATETIME2, 
  [UltimaConsignacao] INT, 
  [ConsignacaoFechada] BIT, 
  [TabelaPrecoPadrao] NVARCHAR(15), 
  [IsentoIPI] BIT, 
  [ObsIsentoIPI] NVARCHAR(100), 
  [Web] BIT DEFAULT 0, 
  [Endereço Número] NVARCHAR(10), 
  [DDD_Fone1] NVARCHAR(7), 
  [DDD_Fone2] NVARCHAR(7), 
  [RG_UF] NVARCHAR(2), 
  [WebEMailMerco] BIT DEFAULT 0, 
  [WebEMailLoja] BIT DEFAULT 0, 
  [WebOrigem] NVARCHAR(50), 
  [Complemento] NVARCHAR(50), 
  [Bairro] NVARCHAR(50), 
  [Fone 2] NVARCHAR(43), 
  [Inscrição] NVARCHAR(23), 
  [Endereço] NVARCHAR(211), 
  [Fone 1] NVARCHAR(43), 
  [CodGrupo] SMALLINT DEFAULT 0, 
  [TotDinheiroBoletos] FLOAT DEFAULT 0, 
  [TotCheques] FLOAT DEFAULT 0, 
  [TotCartoes] FLOAT DEFAULT 0, 
  [TotRecebido] FLOAT DEFAULT 0, 
  [AgenciaPublicidade] BIT, 
  [NomeSacadorAvalista] NVARCHAR(40), 
  [CPFSacadorAvalista] NVARCHAR(20), 
  [CPFCNPJSacadorAvalista] NVARCHAR(4), 
  [SadigWeb_Tipo] NVARCHAR(40), 
  [InscricaoMunicipal] NVARCHAR(20), 
  [CNAE] NVARCHAR(10), 
  [Pais] NVARCHAR(60), 
  [InscricaoSuframa] NVARCHAR(9), 
  [IndicadorIE] INT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Cli_For - Crédito'
--

IF object_id(N'Cli_For - Crédito', 'U') IS NOT NULL DROP TABLE [Cli_For - Crédito]

CREATE TABLE [Cli_For - Crédito] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Empresa] NVARCHAR(50), 
  [Salário] NVARCHAR(20), 
  [Contratação] NVARCHAR(15), 
  [Telefone] NVARCHAR(20), 
  [Conjugê] NVARCHAR(50), 
  [C_CPF] NVARCHAR(30), 
  [C_Empresa] NVARCHAR(50), 
  [C_Salário] NVARCHAR(20), 
  [C_Cargo] NVARCHAR(20), 
  [C_Contratação] NVARCHAR(15), 
  [C_Telefone] NVARCHAR(20), 
  [OBS] NVARCHAR(MAX), 
  [Cargo] NVARCHAR(100), 
  [PercentualLimite] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'CliForCaract'
--

IF object_id(N'CliForCaract', 'U') IS NOT NULL DROP TABLE [CliForCaract]

CREATE TABLE [CliForCaract] (
  [CodCliCaract] INT NOT NULL, 
  [TipoCliCaract] NVARCHAR(1), 
  [CodCaract] INT NOT NULL, 
  [ValCaract] NVARCHAR(255), 
  PRIMARY KEY ([CodCliCaract], [CodCaract])
)

--
-- Table structure for table 'CliForNumeravel'
--

IF object_id(N'CliForNumeravel', 'U') IS NOT NULL DROP TABLE [CliForNumeravel]

CREATE TABLE [CliForNumeravel] (
  [CodCliNumer] INT NOT NULL, 
  [TipoCliNumer] NVARCHAR(1) NOT NULL, 
  [CodProdNumer] NVARCHAR(20), 
  [Data1Numer] DATETIME2, 
  [Data2Numer] DATETIME2, 
  [CodRefDocNumer] NVARCHAR(20), 
  [CodNumer] NVARCHAR(25) NOT NULL, 
  PRIMARY KEY ([CodCliNumer], [TipoCliNumer], [CodNumer])
)

--
-- Table structure for table 'CNABCarteira'
--

IF object_id(N'CNABCarteira', 'U') IS NOT NULL DROP TABLE [CNABCarteira]

CREATE TABLE [CNABCarteira] (
  [NumeroCarteira] NVARCHAR(3) NOT NULL, 
  [Banco] NVARCHAR(25), 
  PRIMARY KEY ([NumeroCarteira])
)

--
-- Table structure for table 'CodigoBeneficio'
--

IF object_id(N'CodigoBeneficio', 'U') IS NOT NULL DROP TABLE [CodigoBeneficio]

CREATE TABLE [CodigoBeneficio] (
  [Estado] NVARCHAR(2), 
  [CodigoBenef] NVARCHAR(10)
)

--
-- Table structure for table 'CodigoNBM'
--

IF object_id(N'CodigoNBM', 'U') IS NOT NULL DROP TABLE [CodigoNBM]

CREATE TABLE [CodigoNBM] (
  [Código] NVARCHAR(8) NOT NULL, 
  [Nome] NVARCHAR(100), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Códigos da Grade'
--

IF object_id(N'Códigos da Grade', 'U') IS NOT NULL DROP TABLE [Códigos da Grade]

CREATE TABLE [Códigos da Grade] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Código Original] NVARCHAR(20), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Comissão'
--

IF object_id(N'Comissão', 'U') IS NOT NULL DROP TABLE [Comissão]

CREATE TABLE [Comissão] (
  [Data] DATETIME2 NOT NULL, 
  [Vendedor] INT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Qtde] REAL DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Comissão] FLOAT DEFAULT 0, 
  [Sequência] INT DEFAULT 0, 
  [Cliente] INT DEFAULT 0, 
  [Tabela] NVARCHAR(15), 
  [Contador] INT NOT NULL IDENTITY, 
  [Filial] INT DEFAULT 0, 
  [VlPagoEmCartao] FLOAT DEFAULT 0, 
  [VlPagoEmCartaoComRetencao] FLOAT DEFAULT 0, 
  [TaxaRetencao] REAL DEFAULT 0, 
  PRIMARY KEY ([Data], [Vendedor], [Produto], [Tamanho], [Cor], [Edição], [Contador])
)

--
-- Table structure for table 'Comissão Serviços'
--

IF object_id(N'Comissão Serviços', 'U') IS NOT NULL DROP TABLE [Comissão Serviços]

CREATE TABLE [Comissão Serviços] (
  [Data] DATETIME2 NOT NULL, 
  [Vendedor] INT NOT NULL DEFAULT 0, 
  [Serviço] INT NOT NULL DEFAULT 0, 
  [Descrição] NVARCHAR(70), 
  [Tempo] NVARCHAR(10), 
  [Valor] FLOAT DEFAULT 0, 
  [Comissão] FLOAT DEFAULT 0, 
  [Valor Comissão] FLOAT DEFAULT 0, 
  [Sequência] INT DEFAULT 0, 
  [Cliente] INT DEFAULT 0, 
  [Contador] INT NOT NULL IDENTITY, 
  [Filial] INT DEFAULT 0, 
  PRIMARY KEY ([Data], [Vendedor], [Serviço], [Contador])
)

--
-- Table structure for table 'Configurações'
--

IF object_id(N'Configurações', 'U') IS NOT NULL DROP TABLE [Configurações]

CREATE TABLE [Configurações] (
  [Nome] NVARCHAR(40) NOT NULL, 
  [Configuração 1] NVARCHAR(30), 
  [Configuração 2] BIT, 
  PRIMARY KEY ([Nome])
)

--
-- Table structure for table 'Consignação Entrada'
--

IF object_id(N'Consignação Entrada', 'U') IS NOT NULL DROP TABLE [Consignação Entrada]

CREATE TABLE [Consignação Entrada] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Fornecedor] INT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Ordem] INT NOT NULL DEFAULT 0, 
  [Data Operação] DATETIME2, 
  [Saldo Anterior] INT DEFAULT 0, 
  [Vendas] INT DEFAULT 0, 
  [Devolução] INT DEFAULT 0, 
  [Empréstimo Recebido] INT DEFAULT 0, 
  [Saldo Atual] INT DEFAULT 0, 
  [Preço Unitário] REAL DEFAULT 0, 
  [Data Cobrança] DATETIME2, 
  [Concluído] BIT, 
  [Gerado Compra] BIT, 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Filial], [Sequência], [Fornecedor], [Produto], [Tamanho], [Cor], [Edição], [Ordem])
)

--
-- Table structure for table 'Consignação Saída'
--

IF object_id(N'Consignação Saída', 'U') IS NOT NULL DROP TABLE [Consignação Saída]

CREATE TABLE [Consignação Saída] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Cliente] INT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Ordem] INT NOT NULL DEFAULT 0, 
  [Data Operação] DATETIME2, 
  [Saldo Anterior] INT DEFAULT 0, 
  [Vendas Cliente] INT DEFAULT 0, 
  [Devolução] INT DEFAULT 0, 
  [Novo Empréstimo] INT DEFAULT 0, 
  [Saldo Atual] INT DEFAULT 0, 
  [Preço Unitário] REAL DEFAULT 0, 
  [Data Cobrança] DATETIME2, 
  [Concluído] BIT, 
  [Gerado Venda] BIT, 
  [Data Alteração] NVARCHAR(10), 
  [QtdeVendidaAcumulada] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Sequência], [Cliente], [Produto], [Tamanho], [Cor], [Edição], [Ordem])
)

--
-- Table structure for table 'Conta Cliente'
--

IF object_id(N'Conta Cliente', 'U') IS NOT NULL DROP TABLE [Conta Cliente]

CREATE TABLE [Conta Cliente] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Cliente] INT DEFAULT 0, 
  [Data] DATETIME2, 
  [Produto] NVARCHAR(20), 
  [Descrição] NVARCHAR(70), 
  [Contador] INT NOT NULL IDENTITY, 
  [Qtde] REAL DEFAULT 0, 
  [Valor] REAL DEFAULT 0, 
  [Valor Pago] REAL DEFAULT 0, 
  [Data Pagamento] DATETIME2, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [TabPrecos] NVARCHAR(15), 
  PRIMARY KEY ([Filial], [Sequência], [Contador])
)

--
-- Table structure for table 'ContaClienteRecebimento'
--

IF object_id(N'ContaClienteRecebimento', 'U') IS NOT NULL DROP TABLE [ContaClienteRecebimento]

CREATE TABLE [ContaClienteRecebimento] (
  [Filial] SMALLINT, 
  [Contador] INT, 
  [Sequencia] INT, 
  [Recebe - Dinheiro] FLOAT, 
  [Recebe - Emp Cartão] INT, 
  [Recebe - Num Cartão] NVARCHAR(20), 
  [Recebe - Cartão] FLOAT, 
  [Recebe - Vale] FLOAT, 
  [Total Prazo] FLOAT, 
  [Tipo Parcela] NVARCHAR(1), 
  [Conta] SMALLINT, 
  [Parcela Cartão] NVARCHAR(1), 
  [Qtde Parcelas] SMALLINT, 
  [Valor Parcela] FLOAT, 
  [Valor Recebido] FLOAT, 
  [Troco] FLOAT
)

--
-- Table structure for table 'Contas a Pagar'
--

IF object_id(N'Contas a Pagar', 'U') IS NOT NULL DROP TABLE [Contas a Pagar]

CREATE TABLE [Contas a Pagar] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Fornecedor] INT NOT NULL DEFAULT 0, 
  [Contador] INT NOT NULL IDENTITY, 
  [Data Emissão] DATETIME2, 
  [Descrição] NVARCHAR(30), 
  [Vencimento] DATETIME2 NOT NULL, 
  [Valor] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [Acréscimo] FLOAT DEFAULT 0, 
  [Valor Pago] FLOAT DEFAULT 0, 
  [Pagamento] DATETIME2, 
  [Sequência] INT DEFAULT 0, 
  [Nota] NVARCHAR(15), 
  [Centro de Custo] INT DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [OrigemDinheiro] NVARCHAR(255), 
  PRIMARY KEY ([Filial], [Vencimento], [Fornecedor], [Contador])
)

--
-- Table structure for table 'Contas a Receber'
--

IF object_id(N'Contas a Receber', 'U') IS NOT NULL DROP TABLE [Contas a Receber]

CREATE TABLE [Contas a Receber] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Cliente] INT DEFAULT 0, 
  [Sequência] INT DEFAULT 0, 
  [Contador] INT NOT NULL IDENTITY, 
  [Tipo] NVARCHAR(1) NOT NULL, 
  [Tipo Parcelamento] NVARCHAR(1), 
  [Conta Boleto] SMALLINT DEFAULT 0, 
  [Data Emissão] DATETIME2, 
  [Vencimento] DATETIME2 NOT NULL, 
  [Valor Cartão] FLOAT DEFAULT 0, 
  [Valor] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [Acréscimo] FLOAT DEFAULT 0, 
  [Valor Recebido] FLOAT DEFAULT 0, 
  [Data Recebimento] DATETIME2, 
  [Vendedor] INT DEFAULT 0, 
  [Banco] INT DEFAULT 0, 
  [Cheque] NVARCHAR(10), 
  [Administradora] SMALLINT DEFAULT 0, 
  [Cartão] NVARCHAR(20), 
  [Descrição] NVARCHAR(40), 
  [Nota] INT DEFAULT 0, 
  [Fatura] NVARCHAR(10), 
  [Parcela] INT DEFAULT 0, 
  [Devolvido] BIT, 
  [Impresso] BIT, 
  [Carnet Impresso] BIT, 
  [Processado] BIT, 
  [Data Alteração] NVARCHAR(10), 
  [CNAB_CodMovRet] SMALLINT, 
  [CNAB_DescrMovRet] NVARCHAR(255), 
  [CNAB_CodIdComplementar] SMALLINT, 
  [CNAB_NossoNumero] NVARCHAR(40), 
  [CNAB_DigitoVerificador] NVARCHAR(3), 
  [CNAB_Carteira] NVARCHAR(3), 
  [CNAB_Bordero] INT, 
  [CarneCodigoBarras] NVARCHAR(40), 
  [FornecedorCreditado] INT, 
  [SequenciaEntrada] INT, 
  [Pendencia] BIT DEFAULT 0, 
  PRIMARY KEY ([Tipo], [Filial], [Vencimento], [Contador])
)

--
-- Table structure for table 'Contas Bancárias'
--

IF object_id(N'Contas Bancárias', 'U') IS NOT NULL DROP TABLE [Contas Bancárias]

CREATE TABLE [Contas Bancárias] (
  [Código] SMALLINT NOT NULL DEFAULT 0, 
  [Agência] NVARCHAR(15), 
  [Conta] NVARCHAR(15), 
  [Banco] INT DEFAULT 0, 
  [Descrição] NVARCHAR(30), 
  [Gerente] NVARCHAR(30), 
  [Telefone] NVARCHAR(30), 
  [Contabilidade] INT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Contatos'
--

IF object_id(N'Contatos', 'U') IS NOT NULL DROP TABLE [Contatos]

CREATE TABLE [Contatos] (
  [Cliente] INT NOT NULL DEFAULT 0, 
  [Seqüência] INT NOT NULL IDENTITY, 
  [Contato] NVARCHAR(30), 
  [Cargo] NVARCHAR(20), 
  [Dia Aniversário] SMALLINT DEFAULT 0, 
  [Mês Aniversário] NVARCHAR(3), 
  [Ramal] NVARCHAR(10), 
  [email] NVARCHAR(40), 
  PRIMARY KEY ([Cliente], [Seqüência])
)

--
-- Table structure for table 'Contatos Efetuados'
--

IF object_id(N'Contatos Efetuados', 'U') IS NOT NULL DROP TABLE [Contatos Efetuados]

CREATE TABLE [Contatos Efetuados] (
  [Cliente] INT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Seqüência] INT NOT NULL IDENTITY, 
  [Descrição] NVARCHAR(MAX), 
  [Pendência] BIT, 
  [Data Aviso] DATETIME2, 
  PRIMARY KEY ([Cliente], [Data], [Seqüência])
)

--
-- Table structure for table 'Contrato'
--

IF object_id(N'Contrato', 'U') IS NOT NULL DROP TABLE [Contrato]

CREATE TABLE [Contrato] (
  [Num Autorizacao] INT NOT NULL, 
  [Cod Cliente] INT NOT NULL, 
  [Radio] NVARCHAR(20), 
  [Cod Fornecedor] INT, 
  [Programacao] NVARCHAR(15), 
  [Dia 01] INT, 
  [Dia 02] INT, 
  [Dia 03] INT, 
  [Dia 04] INT, 
  [Dia 05] INT, 
  [Dia 06] INT, 
  [Dia 07] INT, 
  [Dia 08] INT, 
  [Dia 09] INT, 
  [Dia 10] INT, 
  [Dia 11] INT, 
  [Dia 12] INT, 
  [Dia 13] INT, 
  [Dia 14] INT, 
  [Dia 15] INT, 
  [Dia 16] INT, 
  [Dia 17] INT, 
  [Dia 18] INT, 
  [Dia 19] INT, 
  [Dia 20] INT, 
  [Dia 21] INT, 
  [Dia 22] INT, 
  [Dia 23] INT, 
  [Dia 24] INT, 
  [Dia 25] INT, 
  [Dia 26] INT, 
  [Dia 27] INT, 
  [Dia 28] INT, 
  [Dia 29] INT, 
  [Dia 30] INT, 
  [Dia 31] INT, 
  [Total de Insercoes] FLOAT, 
  [Valor Unitario] FLOAT, 
  [Valor Total] FLOAT, 
  [Patrocinio] NVARCHAR(MAX), 
  [Observacoes] NVARCHAR(MAX), 
  [Tipo Comercial] NVARCHAR(15), 
  [Periodo Ini] DATETIME2, 
  [Periodo Fin] DATETIME2, 
  [Faixa Ini] NVARCHAR(7), 
  [Faixa Fin] NVARCHAR(7), 
  [Frequencia] NVARCHAR(3), 
  [Duracao] FLOAT, 
  [Mes] NVARCHAR(3), 
  [Condicoes Pagamento] NVARCHAR(15), 
  [Data Assinatura] DATETIME2, 
  [VlTotContrato] FLOAT DEFAULT 0, 
  [Cod Radio] INT, 
  [Cod TipoComercial] INT, 
  [Cod Vendedor] INT, 
  [Comissao] FLOAT, 
  PRIMARY KEY ([Num Autorizacao], [Cod Cliente])
)

--
-- Table structure for table 'Cores'
--

IF object_id(N'Cores', 'U') IS NOT NULL DROP TABLE [Cores]

CREATE TABLE [Cores] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(30), 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Cotações'
--

IF object_id(N'Cotações', 'U') IS NOT NULL DROP TABLE [Cotações]

CREATE TABLE [Cotações] (
  [Moeda] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Cotação] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Moeda], [Data])
)

--
-- Table structure for table 'Cupom_temp'
--

IF object_id(N'Cupom_temp', 'U') IS NOT NULL DROP TABLE [Cupom_temp]

CREATE TABLE [Cupom_temp] (
  [N_NF] INT, 
  [Serie] INT, 
  [CNPJ] NVARCHAR(255), 
  [Codigo] NVARCHAR(255), 
  [Descricao] NVARCHAR(255), 
  [Qtd] NVARCHAR(255), 
  [Un] NVARCHAR(255), 
  [vl_unit] NVARCHAR(255), 
  [vl_total] NVARCHAR(255)
)

--
-- Table structure for table 'Diferimento'
--

IF object_id(N'Diferimento', 'U') IS NOT NULL DROP TABLE [Diferimento]

CREATE TABLE [Diferimento] (
  [Filial] SMALLINT NOT NULL, 
  [Total] FLOAT, 
  [Base] FLOAT, 
  [EstadoCorrente] NVARCHAR(2), 
  [ObsDiferimento] NVARCHAR(70), 
  PRIMARY KEY ([Filial])
)

--
-- Table structure for table 'DRE_anexos'
--

IF object_id(N'DRE_anexos', 'U') IS NOT NULL DROP TABLE [DRE_anexos]

CREATE TABLE [DRE_anexos] (
  [CodigoAnexo] INT, 
  [Obs] NVARCHAR(150), 
  [ValorDe] MONEY, 
  [ValorAte] MONEY, 
  [Aliquota] MONEY, 
  [ValorRedutor] MONEY
)

--
-- Table structure for table 'DRE_quick'
--

IF object_id(N'DRE_quick', 'U') IS NOT NULL DROP TABLE [DRE_quick]

CREATE TABLE [DRE_quick] (
  [CodigoDRE] INT IDENTITY, 
  [Filial] INT, 
  [Usuario] INT, 
  [DataANO] INT, 
  [DataMES] INT, 
  [DataCriacao] DATETIME2, 
  [Obs] NVARCHAR(255), 
  [ReceitaBruta] MONEY, 
  [Devolucoes] MONEY, 
  [ImpostoSobreVendas] MONEY, 
  [ReceitaOperacionalLiquida] MONEY, 
  [CMV] MONEY, 
  [LucroBruto] MONEY, 
  [DespesasAdministrativas] MONEY, 
  [DespesasComerciais] MONEY, 
  [DespesasDepreciacao] MONEY, 
  [DespesasFinanceiras] MONEY, 
  [ReceitasFinanceiras] MONEY, 
  [LucroPrejuizoOperacional] MONEY, 
  [DespesasNaoOperacionais] MONEY, 
  [ReceitasNaoOperacionais] MONEY, 
  [LAIR] MONEY, 
  [ProvisaoIR] MONEY, 
  [ProvisaoCSLL] MONEY, 
  [LucroLiquido] MONEY
)

--
-- Table structure for table 'Edições'
--

IF object_id(N'Edições', 'U') IS NOT NULL DROP TABLE [Edições]

CREATE TABLE [Edições] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(30), 
  PRIMARY KEY ([Produto], [Código])
)

--
-- Table structure for table 'Email'
--

IF object_id(N'Email', 'U') IS NOT NULL DROP TABLE [Email]

CREATE TABLE [Email] (
  [Filial] SMALLINT, 
  [ServidorSmtp] NVARCHAR(255), 
  [ServidorPop3] NVARCHAR(255), 
  [Autenticacao] BIT, 
  [AutenticacaoPop3] BIT, 
  [Usuario] NVARCHAR(255), 
  [Senha] NVARCHAR(255), 
  [NomeExibicaoRemetente] NVARCHAR(255), 
  [EmailRemetente] NVARCHAR(255)
)

--
-- Table structure for table 'Encomendas'
--

IF object_id(N'Encomendas', 'U') IS NOT NULL DROP TABLE [Encomendas]

CREATE TABLE [Encomendas] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Cliente] INT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Obervações] NVARCHAR(50), 
  [Qtde] REAL DEFAULT 0, 
  [Qtde Disponível] REAL DEFAULT 0, 
  PRIMARY KEY ([Filial], [Data], [Cliente], [Produto])
)

--
-- Table structure for table 'Entradas'
--

IF object_id(N'Entradas', 'U') IS NOT NULL DROP TABLE [Entradas]

CREATE TABLE [Entradas] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Operação] INT DEFAULT 0, 
  [Digitador] INT DEFAULT 0, 
  [Fornecedor] INT DEFAULT 0, 
  [Observações] NVARCHAR(70), 
  [Nota Fiscal] NVARCHAR(15), 
  [Data Emissão] DATETIME2, 
  [Pedido] NVARCHAR(15), 
  [Forma Pagto] SMALLINT DEFAULT 0, 
  [Produtos] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [IPI] FLOAT DEFAULT 0, 
  [Frete] FLOAT DEFAULT 0, 
  [Base ICM] FLOAT DEFAULT 0, 
  [Valor ICM] FLOAT DEFAULT 0, 
  [Base ICM Subs] FLOAT DEFAULT 0, 
  [Valor ICM Subs] FLOAT DEFAULT 0, 
  [Total] FLOAT DEFAULT 0, 
  [Dinheiro Caixa] FLOAT DEFAULT 0, 
  [Cheque Caixa] FLOAT DEFAULT 0, 
  [Caixa] SMALLINT DEFAULT 0, 
  [Conta] SMALLINT DEFAULT 0, 
  [Num Cheque] NVARCHAR(10), 
  [Bom para] DATETIME2, 
  [Valor Cheque] FLOAT DEFAULT 0, 
  [Descrição] NVARCHAR(40), 
  [Efetivada] BIT, 
  [Nota Impressa] INT DEFAULT 0, 
  [Nota Cancelada] BIT DEFAULT 0, 
  [Data Acerto Empréstimo] DATETIME2, 
  [WebOrderFormID] INT, 
  [CentroCusto] INT, 
  [ConsignacaoMestre] INT, 
  [obs_Obs1] NVARCHAR(30), 
  [obs_Obs2] NVARCHAR(30), 
  [obs_Obs3] NVARCHAR(30), 
  [obs_Obs4] NVARCHAR(30), 
  [obs_Obs5] NVARCHAR(30), 
  [obs_Obs6] NVARCHAR(30), 
  [obs_Obs7] NVARCHAR(30), 
  [obs_Obs8] NVARCHAR(30), 
  [obs_Transportadora] NVARCHAR(50), 
  [obs_Placa] NVARCHAR(8), 
  [obs_Uf] NVARCHAR(2), 
  [obs_Qtde] NVARCHAR(10), 
  [obs_Especie] NVARCHAR(10), 
  [obs_Marca] NVARCHAR(10), 
  [obs_PesoLiquido] FLOAT, 
  [obs_PesoBruto] FLOAT, 
  [obs_FretePago] SMALLINT, 
  [ConsignacaoFechada] BIT, 
  [SerieNF] NVARCHAR(3), 
  [InfoICMSporUF] BIT DEFAULT 0, 
  [NSU] FLOAT DEFAULT 0, 
  [NSU_Data] DATETIME2, 
  [NSU_Hora] DATETIME2, 
  [ModeloDocumentoFiscal] NVARCHAR(2), 
  [Troco] FLOAT DEFAULT 0, 
  [NumeroDI] NVARCHAR(10), 
  [CodigoExportador] NVARCHAR(60), 
  [DataDeRegistroDI] DATETIME2, 
  [UFDesembaracoDI] NVARCHAR(2), 
  [LocalDesembaracoDI] NVARCHAR(60), 
  [DataDesembaracoDI] DATETIME2, 
  [NumeroAdicaoDI] INT DEFAULT 0, 
  [NumeroSeqItemAdicaoDI] INT DEFAULT 0, 
  [CodigoFabricanteAdicaoDI] NVARCHAR(60), 
  [DescontoAdicaoDI] FLOAT DEFAULT 0, 
  [TotalNCM] FLOAT DEFAULT 0, 
  [Consumidor_Final] INT DEFAULT 0, 
  [Presenca_Comprador] INT DEFAULT 0, 
  [TotalDesoneracaoICMS] FLOAT DEFAULT 0, 
  [FinalidadeNFe] INT DEFAULT 0, 
  [ChaveReferenciada] NVARCHAR(100), 
  [obs_infCpl1] NVARCHAR(255), 
  [obs_infCpl2] NVARCHAR(255), 
  PRIMARY KEY ([Filial], [Sequência])
)

--
-- Table structure for table 'Entradas - Produtos'
--

IF object_id(N'Entradas - Produtos', 'U') IS NOT NULL DROP TABLE [Entradas - Produtos]

CREATE TABLE [Entradas - Produtos] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Linha] SMALLINT NOT NULL DEFAULT 0, 
  [Código] NVARCHAR(20), 
  [Qtde] REAL DEFAULT 0, 
  [Preço] REAL DEFAULT 0, 
  [Desconto] REAL DEFAULT 0, 
  [ICM] REAL DEFAULT 0, 
  [IPI] REAL DEFAULT 0, 
  [Preço Final] REAL DEFAULT 0, 
  [Etiqueta] BIT, 
  [Código sem Grade] NVARCHAR(20), 
  [InGeradoViaConsig] BIT, 
  [ConsignacaoFechada] BIT, 
  [IndiceFinanceiro] FLOAT DEFAULT 0, 
  [QtdeAtual] REAL DEFAULT 0, 
  [Selecionado] BIT, 
  [Acertado] BIT, 
  [EntradaConsignada] BIT, 
  [ValorIcmsRetido] FLOAT DEFAULT 0, 
  [ValorICMSDesonerado] FLOAT DEFAULT 0, 
  [MotivoDesoneracaoICMS] INT DEFAULT 0, 
  [Valor_Aprox_Impostos] FLOAT DEFAULT 0, 
  [Percentual_Diferimento] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Sequência], [Linha])
)

--
-- Table structure for table 'Estados'
--

IF object_id(N'Estados', 'U') IS NOT NULL DROP TABLE [Estados]

CREATE TABLE [Estados] (
  [Estado] NVARCHAR(50) NOT NULL, 
  [ICM] INT DEFAULT 0, 
  PRIMARY KEY ([Estado])
)

--
-- Table structure for table 'Estoque'
--

IF object_id(N'Estoque', 'U') IS NOT NULL DROP TABLE [Estoque]

CREATE TABLE [Estoque] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Classe] INT DEFAULT 0, 
  [Sub Classe] INT DEFAULT 0, 
  [Estoque Anterior] REAL DEFAULT 0, 
  [Vendas] REAL DEFAULT 0, 
  [Valor Vendas] FLOAT DEFAULT 0, 
  [Compras] REAL DEFAULT 0, 
  [Valor Compras] FLOAT DEFAULT 0, 
  [Transf Saída] REAL DEFAULT 0, 
  [Valor T Saída] FLOAT DEFAULT 0, 
  [Transf Entra] REAL DEFAULT 0, 
  [Valor T Entra] FLOAT DEFAULT 0, 
  [Ajuste Saída] REAL DEFAULT 0, 
  [Valor Ajuste Saída] FLOAT DEFAULT 0, 
  [Ajuste Entra] REAL DEFAULT 0, 
  [Valor Ajuste Entra] FLOAT DEFAULT 0, 
  [Grátis Saída] REAL DEFAULT 0, 
  [Valor Grátis Saída] FLOAT DEFAULT 0, 
  [Grátis Entra] REAL DEFAULT 0, 
  [Valor Grátis Entra] FLOAT DEFAULT 0, 
  [Quebras] REAL DEFAULT 0, 
  [Valor Quebras] FLOAT DEFAULT 0, 
  [Empre Saída] REAL DEFAULT 0, 
  [Valor Empre Saída] FLOAT DEFAULT 0, 
  [Empre Entra] REAL DEFAULT 0, 
  [Valor Empre Entra] FLOAT DEFAULT 0, 
  [Devolução] REAL DEFAULT 0, 
  [Valor Devolução] FLOAT DEFAULT 0, 
  [Estoque Final] REAL DEFAULT 0, 
  PRIMARY KEY ([Filial], [Data], [Produto], [Tamanho], [Cor], [Edição])
)

--
-- Table structure for table 'Estoque - Tempo'
--

IF object_id(N'Estoque - Tempo', 'U') IS NOT NULL DROP TABLE [Estoque - Tempo]

CREATE TABLE [Estoque - Tempo] (
  [Produto] FLOAT NOT NULL DEFAULT 0, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Classe] INT DEFAULT 0, 
  [Estoque Final] REAL DEFAULT 0, 
  [Data] DATETIME2, 
  [Preço] REAL, 
  PRIMARY KEY ([Produto], [Tamanho], [Cor])
)

--
-- Table structure for table 'Estoque Final'
--

IF object_id(N'Estoque Final', 'U') IS NOT NULL DROP TABLE [Estoque Final]

CREATE TABLE [Estoque Final] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Estoque Atual] REAL DEFAULT 0, 
  [Classe] INT DEFAULT 0, 
  [Sub Classe] INT DEFAULT 0, 
  [Última Data] NVARCHAR(10), 
  PRIMARY KEY ([Filial], [Produto], [Tamanho], [Cor], [Edição])
)

--
-- Table structure for table 'Etiquetas'
--

IF object_id(N'Etiquetas', 'U') IS NOT NULL DROP TABLE [Etiquetas]

CREATE TABLE [Etiquetas] (
  [Funcionário] INT NOT NULL DEFAULT 0, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Qtde] INT DEFAULT 0, 
  [Sequência] INT DEFAULT 0, 
  [Preço] REAL DEFAULT 0, 
  [Preco2] REAL, 
  PRIMARY KEY ([Funcionário], [Produto], [Tamanho], [Cor])
)

--
-- Table structure for table 'Etiquetas - Tempo'
--

IF object_id(N'Etiquetas - Tempo', 'U') IS NOT NULL DROP TABLE [Etiquetas - Tempo]

CREATE TABLE [Etiquetas - Tempo] (
  [Código] NVARCHAR(20), 
  [Código Barra] NVARCHAR(22), 
  [Código Produto] NVARCHAR(20), 
  [Descrição] NVARCHAR(70), 
  [Sem Preço] NVARCHAR(50), 
  [Tamanho] NVARCHAR(30), 
  [Cor] NVARCHAR(30), 
  [Seq] INT NOT NULL IDENTITY, 
  [Texto Grande] NVARCHAR(MAX), 
  [Preco] FLOAT, 
  [Descricao2] NVARCHAR(70), 
  [ImprimirUmaEtiq] BIT DEFAULT 0, 
  [ImprimirPrecoEtiq] BIT DEFAULT 0, 
  [Funcionario] INT DEFAULT 0, 
  [DividirPrecoEtiqueta] INT DEFAULT 0, 
  [Lote] NVARCHAR(15), 
  [DataValidade] DATETIME2, 
  [PrecoPrazo] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Seq])
)

--
-- Table structure for table 'FCP_AdesaoPercentual_Quick'
--

IF object_id(N'FCP_AdesaoPercentual_Quick', 'U') IS NOT NULL DROP TABLE [FCP_AdesaoPercentual_Quick]

CREATE TABLE [FCP_AdesaoPercentual_Quick] (
  [ESTADO_ORIGEM] NVARCHAR(50), 
  [ALIQUOTA1] FLOAT DEFAULT 0, 
  [ALIQUOTA2] FLOAT DEFAULT 0, 
  [ALIQUOTA3] FLOAT DEFAULT 0, 
  [DESCRICAO] NVARCHAR(100)
)

--
-- Table structure for table 'FISParametros'
--

IF object_id(N'FISParametros', 'U') IS NOT NULL DROP TABLE [FISParametros]

CREATE TABLE [FISParametros] (
  [Código] SMALLINT NOT NULL, 
  [Versão] NVARCHAR(15), 
  [Isenção] NVARCHAR(3), 
  [Substituição] NVARCHAR(3), 
  [Não Incidência] NVARCHAR(3), 
  [TabelaPreco] NVARCHAR(15), 
  [OperacoesAcharVenda] NVARCHAR(255), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'FISReg60Analitico'
--

IF object_id(N'FISReg60Analitico', 'U') IS NOT NULL DROP TABLE [FISReg60Analitico]

CREATE TABLE [FISReg60Analitico] (
  [Codigo] INT NOT NULL IDENTITY, 
  [Filial] INT, 
  [Data] DATETIME2, 
  [NrECF] INT, 
  [ST_Aliquota] NVARCHAR(4), 
  [VlrAcumulado] MONEY, 
  [NrSerie] NVARCHAR(20), 
  PRIMARY KEY ([Codigo])
)

--
-- Table structure for table 'FISReg60Mestre'
--

IF object_id(N'FISReg60Mestre', 'U') IS NOT NULL DROP TABLE [FISReg60Mestre]

CREATE TABLE [FISReg60Mestre] (
  [Codigo] INT IDENTITY, 
  [Status] INT, 
  [Filial] INT NOT NULL, 
  [Data] DATETIME2 NOT NULL, 
  [NrECF] INT, 
  [NrCOOInicioDia] INT, 
  [NrCOOFimDia] INT, 
  [NrContReducaoZ] INT, 
  [GTInicioDia] MONEY, 
  [GTFimDia] MONEY, 
  [NrSerie] NVARCHAR(20) NOT NULL, 
  [NrCRO] INT DEFAULT 0, 
  [VendaBruta] MONEY DEFAULT 0, 
  [VendaLiquidaSEF] MONEY DEFAULT 0, 
  PRIMARY KEY ([Filial], [Data], [NrSerie])
)

--
-- Table structure for table 'FISSenhas'
--

IF object_id(N'FISSenhas', 'U') IS NOT NULL DROP TABLE [FISSenhas]

CREATE TABLE [FISSenhas] (
  [Código] SMALLINT NOT NULL, 
  [Abrir Dia] NVARCHAR(8), 
  [Fechar Dia] NVARCHAR(8), 
  [Leitura X] NVARCHAR(8), 
  [Redução Z] NVARCHAR(8), 
  [Leitura Fiscal] NVARCHAR(8), 
  [Tela Senhas] NVARCHAR(8), 
  [Alíquotas] NVARCHAR(8), 
  [Parâmetros] NVARCHAR(8), 
  [Movimentação] NVARCHAR(8), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Forn_Prod'
--

IF object_id(N'Forn_Prod', 'U') IS NOT NULL DROP TABLE [Forn_Prod]

CREATE TABLE [Forn_Prod] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Fornecedor] INT NOT NULL, 
  PRIMARY KEY ([Produto], [Fornecedor])
)

--
-- Table structure for table 'Funcionários'
--

IF object_id(N'Funcionários', 'U') IS NOT NULL DROP TABLE [Funcionários]

CREATE TABLE [Funcionários] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(35), 
  [Apelido] NVARCHAR(10), 
  [Sexo] NVARCHAR(1), 
  [Admissão] DATETIME2, 
  [Nascimento] NVARCHAR(15), 
  [Endereço] NVARCHAR(50), 
  [Cidade] NVARCHAR(25), 
  [Bairro] NVARCHAR(30), 
  [Estado] NVARCHAR(2), 
  [CEP] NVARCHAR(9), 
  [Telefone] NVARCHAR(30), 
  [Cargo] NVARCHAR(20), 
  [Identidade] NVARCHAR(15), 
  [CPF] NVARCHAR(15), 
  [Carteira Trabalho] NVARCHAR(15), 
  [Senha] NVARCHAR(8), 
  [Superusuário] BIT, 
  [Liberado] BIT, 
  [Filial Acesso] SMALLINT DEFAULT 0, 
  [Comissão] REAL DEFAULT 0, 
  [Comissão Serviço] REAL DEFAULT 0, 
  [Recebimento] BIT NOT NULL, 
  [Clientes] BIT NOT NULL, 
  [Mostrar Ajuda] BIT, 
  [Movimentar Caixa] BIT, 
  [Pasta Compras] BIT, 
  [Pasta Pagar] BIT, 
  [Pasta Receber] BIT, 
  [Pasta Cheques] BIT, 
  [Pasta Outras] BIT, 
  [Pasta Conta] BIT, 
  [Pasta Serviços] BIT, 
  [Custo Produtos] BIT, 
  [Observação] NVARCHAR(MAX), 
  [Recebimento Saídas] BIT, 
  [ValorP] NVARCHAR(30), 
  [bPermiteDesconto] BIT, 
  [nPercDesconto] REAL, 
  [VRVisualizarEstoque] BIT, 
  [VRVisualizarPreco] BIT, 
  [MargemLimiteCredito] REAL, 
  [PermiteAcharVenda] BIT, 
  [VR_PermiteVisualizarLimiteCredito] BIT, 
  [SenhaConfirmarCRDiff] BIT, 
  [ImprimirTicket] BIT, 
  [SenhaClear] BIT, 
  [Marketing] BIT, 
  [Supervisor] INT DEFAULT 0, 
  [AllowDescProd] BIT DEFAULT 0, 
  [Ativo] BIT DEFAULT 0, 
  [SadigWeb_CDRC] NVARCHAR(20), 
  [ContatosEfetuadosLembrarEm] BIT DEFAULT 0, 
  [LucroMinimoPermitido] BIT DEFAULT 0, 
  [bMostrarTelaPesquisaProdutoTipoFoto] BIT DEFAULT 0, 
  [bUsuarioAcessoApenasTelaVendaRapida] BIT DEFAULT 0, 
  [isPrestServ] BIT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Grade - Tempo'
--

IF object_id(N'Grade - Tempo', 'U') IS NOT NULL DROP TABLE [Grade - Tempo]

CREATE TABLE [Grade - Tempo] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  PRIMARY KEY ([Produto], [Tamanho], [Cor])
)

--
-- Table structure for table 'GrupoFiscal'
--

IF object_id(N'GrupoFiscal', 'U') IS NOT NULL DROP TABLE [GrupoFiscal]

CREATE TABLE [GrupoFiscal] (
  [Código] INT NOT NULL, 
  [Nome] NVARCHAR(50), 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Grupos Interesse'
--

IF object_id(N'Grupos Interesse', 'U') IS NOT NULL DROP TABLE [Grupos Interesse]

CREATE TABLE [Grupos Interesse] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(25), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'GruposDeClientes'
--

IF object_id(N'GruposDeClientes', 'U') IS NOT NULL DROP TABLE [GruposDeClientes]

CREATE TABLE [GruposDeClientes] (
  [Filial] SMALLINT NOT NULL, 
  [NomeG1] NVARCHAR(40), 
  [NomeG2] NVARCHAR(40), 
  [NomeG3] NVARCHAR(40), 
  [NomeG4] NVARCHAR(40), 
  [LimiteIniG1] FLOAT, 
  [LimiteIniG2] FLOAT, 
  [LimiteIniG3] FLOAT, 
  [LimiteIniG4] FLOAT, 
  [LimiteFinG1] FLOAT, 
  [LimiteFinG2] FLOAT, 
  [LimiteFinG3] FLOAT, 
  [CodigoG1] SMALLINT, 
  [CodigoG2] SMALLINT, 
  [CodigoG3] SMALLINT, 
  [CodigoG4] SMALLINT, 
  PRIMARY KEY ([Filial])
)

--
-- Table structure for table 'ICMS_PERCENTUAL_ESTADOS'
--

IF object_id(N'ICMS_PERCENTUAL_ESTADOS', 'U') IS NOT NULL DROP TABLE [ICMS_PERCENTUAL_ESTADOS]

CREATE TABLE [ICMS_PERCENTUAL_ESTADOS] (
  [ESTADO_ORIGEM] NVARCHAR(50), 
  [ESTADO_DESTINO] NVARCHAR(50), 
  [ALIQUOTA] FLOAT DEFAULT 0, 
  [ADESAO] NVARCHAR(50)
)

--
-- Table structure for table 'Lançamentos Bancários'
--

IF object_id(N'Lançamentos Bancários', 'U') IS NOT NULL DROP TABLE [Lançamentos Bancários]

CREATE TABLE [Lançamentos Bancários] (
  [Conta] INT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Ordem] INT NOT NULL IDENTITY, 
  [Descrição] NVARCHAR(40), 
  [Cheque] NVARCHAR(10), 
  [Saldo Anterior] FLOAT DEFAULT 0, 
  [Débito] FLOAT DEFAULT 0, 
  [Crédito] FLOAT DEFAULT 0, 
  [Saldo Atual] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Conta], [Data], [Ordem])
)

--
-- Table structure for table 'Livro Ponto'
--

IF object_id(N'Livro Ponto', 'U') IS NOT NULL DROP TABLE [Livro Ponto]

CREATE TABLE [Livro Ponto] (
  [Funcionário] INT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Entrada Manhã] NVARCHAR(6), 
  [Saída Manhã] NVARCHAR(6), 
  [Entrada Tarde] NVARCHAR(6), 
  [Saída Tarde] NVARCHAR(6), 
  [Entrada Noite] NVARCHAR(6), 
  [Saída Noite] NVARCHAR(6), 
  [Entrada Extra] NVARCHAR(6), 
  [Saída Extra] NVARCHAR(6), 
  [Horas] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Funcionário], [Data])
)

--
-- Table structure for table 'Mala Direta'
--

IF object_id(N'Mala Direta', 'U') IS NOT NULL DROP TABLE [Mala Direta]

CREATE TABLE [Mala Direta] (
  [Cliente] INT NOT NULL DEFAULT 0, 
  [Grupo] INT NOT NULL DEFAULT 0, 
  PRIMARY KEY ([Cliente], [Grupo])
)

--
-- Table structure for table 'Mala Direta - Tempo'
--

IF object_id(N'Mala Direta - Tempo', 'U') IS NOT NULL DROP TABLE [Mala Direta - Tempo]

CREATE TABLE [Mala Direta - Tempo] (
  [Cliente] INT NOT NULL DEFAULT 0, 
  [Ordem] INT NOT NULL IDENTITY, 
  [Nome] NVARCHAR(30), 
  PRIMARY KEY ([Cliente], [Ordem])
)

--
-- Table structure for table 'MensagensNotaFiscal'
--

IF object_id(N'MensagensNotaFiscal', 'U') IS NOT NULL DROP TABLE [MensagensNotaFiscal]

CREATE TABLE [MensagensNotaFiscal] (
  [Codigo] INT IDENTITY, 
  [Ordem] INT NOT NULL, 
  [TipoFiltroProduto] SMALLINT NOT NULL, 
  [TipoFiltroOpSaida] SMALLINT NOT NULL, 
  [TipoFiltroUF] SMALLINT NOT NULL, 
  [FiltroProduto] NVARCHAR(20) NOT NULL, 
  [FiltroOpSaida] NVARCHAR(20) NOT NULL, 
  [FiltroUF] NVARCHAR(20) NOT NULL, 
  [Mensagem] NVARCHAR(80) NOT NULL
)

--
-- Table structure for table 'Moedas'
--

IF object_id(N'Moedas', 'U') IS NOT NULL DROP TABLE [Moedas]

CREATE TABLE [Moedas] (
  [Código] SMALLINT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(20), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Movimento - Cartoes'
--

IF object_id(N'Movimento - Cartoes', 'U') IS NOT NULL DROP TABLE [Movimento - Cartoes]

CREATE TABLE [Movimento - Cartoes] (
  [Filial] INT NOT NULL, 
  [Sequência] INT NOT NULL, 
  [Ordem] INT NOT NULL, 
  [Administradora] NVARCHAR(25), 
  [Valor] FLOAT, 
  [Parcelas] INT, 
  [ValorParcelas] FLOAT, 
  [NumeroCartao] NVARCHAR(25), 
  [Credito] BIT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Sequência], [Ordem])
)

--
-- Table structure for table 'Movimento - Cheques'
--

IF object_id(N'Movimento - Cheques', 'U') IS NOT NULL DROP TABLE [Movimento - Cheques]

CREATE TABLE [Movimento - Cheques] (
  [Filial] INT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Ordem] INT NOT NULL DEFAULT 0, 
  [Banco] INT DEFAULT 0, 
  [Cheque] NVARCHAR(10), 
  [Bom] DATETIME2, 
  [Valor] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Sequência], [Ordem])
)

--
-- Table structure for table 'Movimento - Parcelas'
--

IF object_id(N'Movimento - Parcelas', 'U') IS NOT NULL DROP TABLE [Movimento - Parcelas]

CREATE TABLE [Movimento - Parcelas] (
  [Filial] INT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Ordem] INT NOT NULL DEFAULT 0, 
  [Bom] DATETIME2, 
  [Valor] FLOAT DEFAULT 0, 
  [Parcelas] INT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Sequência], [Ordem])
)

--
-- Table structure for table 'NCM96'
--

IF object_id(N'NCM96', 'U') IS NOT NULL DROP TABLE [NCM96]

CREATE TABLE [NCM96] (
  [NCM] INT, 
  [DESCRIÇÃO  NCM  96] NVARCHAR(255)
)

--
-- Table structure for table 'NFCE_ENVI'
--

IF object_id(N'NFCE_ENVI', 'U') IS NOT NULL DROP TABLE [NFCE_ENVI]

CREATE TABLE [NFCE_ENVI] (
  [CNPJ] NVARCHAR(255), 
  [ID] INT, 
  [Serie] INT, 
  [N_NF] INT, 
  [C_NF] INT, 
  [Chave] NVARCHAR(100), 
  [Detalhe_Autorizacao] NVARCHAR(255), 
  [Detalhe_Cancelamento] NVARCHAR(255), 
  [Dh_Autorizacao] NVARCHAR(20), 
  [Em_Contingencia] NVARCHAR(1), 
  [Ex_Message] NVARCHAR(255), 
  [Numero] INT, 
  [Numero_Protocolo_Autorizacao] NVARCHAR(100), 
  [O_Id] INT, 
  [Status_Autorizacao] NVARCHAR(100), 
  [Status_Cancelamento] NVARCHAR(100), 
  [Protocolo_Xml] NVARCHAR(255), 
  [URL_QRCode] NVARCHAR(MAX)
)

--
-- Table structure for table 'NFCE_job'
--

IF object_id(N'NFCE_job', 'U') IS NOT NULL DROP TABLE [NFCE_job]

CREATE TABLE [NFCE_job] (
  [CNPJ] NVARCHAR(255), 
  [Xml] NVARCHAR(MAX), 
  [Tipo] NVARCHAR(255), 
  [Serie] INT, 
  [N_NF] INT, 
  [Chave] NVARCHAR(100), 
  [Cupom] NVARCHAR(MAX), 
  [Justificativa] NVARCHAR(100), 
  [Processado] NVARCHAR(255) DEFAULT N'N', 
  [CPF] NVARCHAR(255), 
  [Nome_Consumidor] NVARCHAR(255), 
  [Data_Emissao] NVARCHAR(255), 
  [Total_Tributos] NVARCHAR(255), 
  [Nome_Emitente] NVARCHAR(255), 
  [Endereco_Emitente] NVARCHAR(255), 
  [IE_Emitente] NVARCHAR(255)
)

--
-- Table structure for table 'NFe'
--

IF object_id(N'NFe', 'U') IS NOT NULL DROP TABLE [NFe]

CREATE TABLE [NFe] (
  [Filial] INT NOT NULL, 
  [Sequencia] INT NOT NULL, 
  [TipoMovimento] SMALLINT NOT NULL, 
  [DataHoraEnvio] DATETIME2 NOT NULL, 
  [Status] INT NOT NULL, 
  [Ambiente] SMALLINT NOT NULL, 
  [FormaEmissao] SMALLINT NOT NULL, 
  [Numero] INT NOT NULL, 
  [Serie] NVARCHAR(3) NOT NULL, 
  [Modelo] NVARCHAR(2) NOT NULL, 
  [ChaveAcesso] NVARCHAR(44) NOT NULL, 
  [ProtocoloAutorizacao] NVARCHAR(15) NOT NULL, 
  [DataHoraAutorizacao] DATETIME2, 
  [ProtocoloCancelamento] NVARCHAR(15) NOT NULL, 
  [DataHoraCancelamento] DATETIME2, 
  [nrEvento] INT DEFAULT 0, 
  [oid_xmlLoteBenefix] INT DEFAULT 0, 
  [nomeDanfe] NVARCHAR(120), 
  PRIMARY KEY ([Filial], [Sequencia], [TipoMovimento])
)

--
-- Table structure for table 'NFeCartaCorrecao'
--

IF object_id(N'NFeCartaCorrecao', 'U') IS NOT NULL DROP TABLE [NFeCartaCorrecao]

CREATE TABLE [NFeCartaCorrecao] (
  [Filial] INT DEFAULT 0, 
  [CNPJ] NVARCHAR(15), 
  [DataHora] DATETIME2, 
  [Serie] NVARCHAR(3), 
  [Numero] INT DEFAULT 0, 
  [Descricao] NVARCHAR(255), 
  [arquivoDanfeCC] NVARCHAR(150)
)

--
-- Table structure for table 'NFeInutilizadas'
--

IF object_id(N'NFeInutilizadas', 'U') IS NOT NULL DROP TABLE [NFeInutilizadas]

CREATE TABLE [NFeInutilizadas] (
  [Filial] INT DEFAULT 0, 
  [CNPJ] NVARCHAR(15), 
  [Ano] INT DEFAULT 0, 
  [Serie] NVARCHAR(3), 
  [NumeroInicial] INT DEFAULT 0, 
  [NumeroFinal] INT DEFAULT 0, 
  [Justificativa] NVARCHAR(255), 
  [DataHora] DATETIME2, 
  [Modelo] INT DEFAULT 0
)

--
-- Table structure for table 'NFeRetorno'
--

IF object_id(N'NFeRetorno', 'U') IS NOT NULL DROP TABLE [NFeRetorno]

CREATE TABLE [NFeRetorno] (
  [Filial] INT NOT NULL, 
  [Sequencia] INT NOT NULL, 
  [TipoMovimento] SMALLINT NOT NULL, 
  [DataHora] DATETIME2 NOT NULL, 
  [Protocolo] NVARCHAR(15) NOT NULL, 
  [StatusDescricao] NVARCHAR(255) NOT NULL, 
  [StatusDescricao2] NVARCHAR(255), 
  [DigestValue] NVARCHAR(255), 
  [Status] INT, 
  PRIMARY KEY ([Filial], [Sequencia], [TipoMovimento])
)

--
-- Table structure for table 'NSU'
--

IF object_id(N'NSU', 'U') IS NOT NULL DROP TABLE [NSU]

CREATE TABLE [NSU] (
  [Filial] SMALLINT, 
  [NSU] NVARCHAR(10), 
  [Movimento] NVARCHAR(10), 
  [Motivo] NVARCHAR(20), 
  [Sequencia] INT, 
  [NotaFiscal] INT, 
  [Data_Hora] DATETIME2, 
  [Total] FLOAT
)

--
-- Table structure for table 'Operações Entrada'
--

IF object_id(N'Operações Entrada', 'U') IS NOT NULL DROP TABLE [Operações Entrada]

CREATE TABLE [Operações Entrada] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(50), 
  [Tipo] NVARCHAR(1), 
  [Estoque] BIT, 
  [Dinheiro] BIT, 
  [Comissão] BIT, 
  [Etiquetas] BIT, 
  [Nota] BIT, 
  [ICM] BIT, 
  [IPI] BIT, 
  [Gravar Custo] BIT, 
  [Senha] BIT, 
  [Somar Frete ao Total] BIT, 
  [Somar Frete ao Custo] BIT, 
  [Código Fiscal] NVARCHAR(14), 
  [Base ICM com IPI] BIT, 
  [IPI TOT] BIT, 
  [Locked] BIT, 
  [Estorno] BIT, 
  [Tabela] NVARCHAR(15), 
  [PermitirAlterPreco] BIT, 
  [InformanteProprio] BIT DEFAULT 0, 
  [EmitirNFManualmente] BIT DEFAULT 0, 
  [GravaCustoPrecoListaSemIPI] BIT DEFAULT 0, 
  [SomarFreteCustoProduto] BIT DEFAULT 0, 
  [BaseICMSFrete] BIT DEFAULT 0, 
  [ICMSSobreIPI] BIT DEFAULT 0, 
  [PrecoCustoCalculado] BIT DEFAULT 0, 
  [ModeloDocumentoFiscal] NVARCHAR(2), 
  [CSO] NVARCHAR(3), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Operações Saída'
--

IF object_id(N'Operações Saída', 'U') IS NOT NULL DROP TABLE [Operações Saída]

CREATE TABLE [Operações Saída] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(50), 
  [Tipo] NVARCHAR(1), 
  [Estoque] BIT, 
  [Dinheiro] BIT, 
  [Comissão] BIT, 
  [Etiquetas] BIT, 
  [Nota] BIT, 
  [ICM] BIT, 
  [IPI] BIT, 
  [Calcula ISS] BIT, 
  [Base ICM com IPI] BIT, 
  [Senha] BIT, 
  [Ticket Imprimir] NVARCHAR(30), 
  [IR sobre ISS] REAL DEFAULT 0, 
  [InTelaObsTransp] BIT, 
  [Código Fiscal] NVARCHAR(14), 
  [IPI TOT] BIT, 
  [Perc Icms Frete] INT, 
  [Calcula Icm Frete] BIT, 
  [Soma Frete] BIT, 
  [Locked] BIT, 
  [ControleEntregas] BIT, 
  [OpEntrega] INT, 
  [ExigeAprovacaoOrcamento] BIT, 
  [Validade] BIT, 
  [ComissaoServicos] BIT, 
  [AcertaEmprestimoEntrada] BIT, 
  [InformanteProprio] BIT DEFAULT 0, 
  [SomarSeguro] BIT DEFAULT 0, 
  [EmitirNFManualmente] BIT DEFAULT 0, 
  [AlteraStatusPedidoWeb] BIT DEFAULT 0, 
  [GrupoFiscal] INT DEFAULT 0, 
  [SadigWeb_Tipo] NVARCHAR(15), 
  [SomarProdutosTotalNota] BIT DEFAULT 0, 
  [ExibirTelaNumeroDocumento] BIT DEFAULT 0, 
  [SomaIcmsRetidoTotalNota] BIT DEFAULT 0, 
  [ModeloDocumentoFiscal] NVARCHAR(2), 
  [CSO] NVARCHAR(3), 
  [PermiteMostrarCliente] BIT DEFAULT 0, 
  [ObterTributosProduto_EntradaOuSaida] INT DEFAULT 0, 
  [SomaIpiTotalNota] BIT DEFAULT 0, 
  [SomaIpiTotalBC] BIT DEFAULT 0, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'ParamDevoMat'
--

IF object_id(N'ParamDevoMat', 'U') IS NOT NULL DROP TABLE [ParamDevoMat]

CREATE TABLE [ParamDevoMat] (
  [Filial] SMALLINT NOT NULL, 
  [Operacao] INT, 
  [Caixa] SMALLINT, 
  [Tabela] NVARCHAR(15), 
  PRIMARY KEY ([Filial])
)

--
-- Table structure for table 'Parâmetros Filial'
--

IF object_id(N'Parâmetros Filial', 'U') IS NOT NULL DROP TABLE [Parâmetros Filial]

CREATE TABLE [Parâmetros Filial] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(35), 
  [Razão Social] NVARCHAR(50), 
  [Endereço] NVARCHAR(50), 
  [Bairro] NVARCHAR(30), 
  [Fone] NVARCHAR(30), 
  [Cidade] NVARCHAR(30), 
  [Estado] NVARCHAR(2), 
  [CGC] NVARCHAR(20), 
  [Inscrição] NVARCHAR(20), 
  [Senha Sempre] BIT, 
  [Linhas Digitação] SMALLINT DEFAULT 0, 
  [Linhas Serviço] SMALLINT DEFAULT 0, 
  [Última Movimentação] INT DEFAULT 0, 
  [Última Nota] INT DEFAULT 0, 
  [Venda Sem Estoque] BIT, 
  [Qtde Dígitos Preço] SMALLINT DEFAULT 0, 
  [Senha Gerente] NVARCHAR(8), 
  [Senha Especial] NVARCHAR(8), 
  [Qtde Balança] SMALLINT DEFAULT 0, 
  [Juros] REAL DEFAULT 0, 
  [Cód Comp 1] NVARCHAR(3), 
  [Cód Comp 2] NVARCHAR(3), 
  [Cód Comp 3] NVARCHAR(3), 
  [Cód Oitavo 1] NVARCHAR(3), 
  [Cód Oitavo 2] NVARCHAR(3), 
  [Cód Oitavo 3] NVARCHAR(3), 
  [Nota Saída] NVARCHAR(8), 
  [Nota Entrada] NVARCHAR(8), 
  [VR Linhas Digitação] SMALLINT DEFAULT 0, 
  [VR Código Operação] INT DEFAULT 0, 
  [VR Tab Preço] NVARCHAR(15), 
  [VR Altera Tabela] BIT, 
  [VR Altera Preço] BIT, 
  [VR Cliente] INT DEFAULT 0, 
  [VR Altera Cliente] BIT, 
  [VR Cadastra Cliente] BIT, 
  [VR Desconto] REAL DEFAULT 0, 
  [VR Permite Desconto] BIT, 
  [VR Mostrar Estoque] BIT, 
  [VR Verifica Limite] BIT, 
  [Consulta Tab1] NVARCHAR(15), 
  [Consulta Tab2] NVARCHAR(15), 
  [Consulta Tab3] NVARCHAR(15), 
  [Consulta Tab4] NVARCHAR(15), 
  [Consulta Tab5] NVARCHAR(15), 
  [Consulta Tab6] NVARCHAR(15), 
  [Lista 1] NVARCHAR(80), 
  [Lista 2] NVARCHAR(80), 
  [Lista 3] NVARCHAR(80), 
  [Lista 4] NVARCHAR(80), 
  [Lista 5] NVARCHAR(80), 
  [Cheque Favorecido] NVARCHAR(50), 
  [Cheque Cidade] NVARCHAR(30), 
  [Código Favorecido 1] NVARCHAR(3), 
  [Código Favorecido 2] NVARCHAR(3), 
  [Código Favorecido 3] NVARCHAR(3), 
  [Código Cidade 1] NVARCHAR(3), 
  [Código Cidade 2] NVARCHAR(3), 
  [Código Cidade 3] NVARCHAR(3), 
  [Código Banco 1] NVARCHAR(3), 
  [Código Banco 2] NVARCHAR(3), 
  [Código Banco 3] NVARCHAR(3), 
  [Código Valor 1] NVARCHAR(3), 
  [Código Valor 2] NVARCHAR(3), 
  [Código Valor 3] NVARCHAR(3), 
  [Código Data 1] NVARCHAR(3), 
  [Código Data 2] NVARCHAR(3), 
  [Código Data 3] NVARCHAR(3), 
  [Início Cheque 1] NVARCHAR(3), 
  [Início Cheque 2] NVARCHAR(3), 
  [Imprime Cheque 1] NVARCHAR(3), 
  [Imprime Cheque 2] NVARCHAR(3), 
  [Mensagem Troca] NVARCHAR(140), 
  [Mensagem Etiq 1] NVARCHAR(20), 
  [Mensagem Etiq 2] NVARCHAR(20), 
  [Verifica Agenda] BIT, 
  [Dias Receber] SMALLINT DEFAULT 0, 
  [Tipo Ajuda] NVARCHAR(1), 
  [Tempo Ajuda] REAL DEFAULT 0, 
  [Três Tabelas] BIT, 
  [Tabela 1] NVARCHAR(15), 
  [Tabela 2] NVARCHAR(15), 
  [Tabela 3] NVARCHAR(15), 
  [Usar Grade] BIT, 
  [Usar Edições] BIT, 
  [Usar Códigos Alfa] BIT, 
  [Último Uso] NVARCHAR(10), 
  [VR Permite Rec Rápido] BIT, 
  [VR Permite Dinheiro] BIT, 
  [VR Permite Vales] BIT, 
  [VR Permite Cartão] BIT, 
  [VR Permite Cheques] BIT, 
  [VR Qtde Cheques] INT, 
  [VR Prazo Cheques] INT, 
  [VR Permite Parcela] BIT, 
  [VR Qtde Parcela] INT, 
  [VR Prazo Parcela] INT, 
  [VR Parcela Padrão] NVARCHAR(1), 
  [VR Altera Parcela] BIT, 
  [VR Recebimento Normal] BIT, 
  [VR Conta Padrão] NVARCHAR(1), 
  [VR Conta Usar] SMALLINT DEFAULT 0, 
  [VR Intervalo Parc] INT DEFAULT 0, 
  [Saída Parcela Padrão] NVARCHAR(1), 
  [Saída Altera Parcela] BIT, 
  [Saída Verifica Limite] BIT, 
  [Saída Intervalo Parc] INT DEFAULT 0, 
  [Usar Serviços] BIT, 
  [Alterar Serviços] BIT, 
  [Usa Vários Caixas] BIT, 
  [Arquivos de Ajuda] NVARCHAR(90), 
  [Nome Pesquisa 1] NVARCHAR(10), 
  [Nome Pesquisa 2] NVARCHAR(10), 
  [Nome Pesquisa 3] NVARCHAR(10), 
  [Gerar Conta Paga] BIT, 
  [Impressora Cheques] NVARCHAR(1), 
  [Código Banco Cheques] INT DEFAULT 0, 
  [Imprimir Centavos] BIT, 
  [Superusuário Libera Telas] BIT, 
  [Cód Comprim 1] NVARCHAR(3), 
  [Cód Comprim 2] NVARCHAR(3), 
  [Cód Comprim 3] NVARCHAR(3), 
  [Saida Descr Adicional] BIT, 
  [Saida Altera Preco] BIT, 
  [LinkSerial] NVARCHAR(11), 
  [WorkWeb] BIT, 
  [NrOrcamento] INT, 
  [PesquisaCodigoENome_VR] BIT, 
  [WorkTrafficLight] BIT, 
  [Venda Sem Estoque Saidas] BIT, 
  [OpSaidaOrcVenda] INT, 
  [DescSubTotalRateado] BIT, 
  [VROrdenacaoCombo] BIT, 
  [UltimaConsignacao] INT, 
  [Consignacao_OpEntrada] INT, 
  [Consignacao_OpSaida] INT, 
  [Consignacao_Caixa] INT, 
  [Consignacao_TabelaPrecos] NVARCHAR(15), 
  [Consignacao_OpFechamento] INT, 
  [CheckInstance] BIT, 
  [VerificaEstoqueSaidas] BIT, 
  [DiasBloqueioVenda] INT, 
  [VR_GravarExigeSenhaVend] BIT, 
  [CSLL] FLOAT DEFAULT 0, 
  [COFINS] FLOAT DEFAULT 0, 
  [PIS] FLOAT DEFAULT 0, 
  [IRRF] FLOAT DEFAULT 0, 
  [VR_RecalcularPreço] BIT, 
  [Zero a Esquerda] BIT, 
  [TaxaDesconto] FLOAT DEFAULT 0, 
  [BoletoPadrao] NVARCHAR(30), 
  [TicketPadrao] NVARCHAR(30), 
  [Permitir5Casas] BIT, 
  [CliWebComprarPrazo] BIT DEFAULT 0, 
  [VerificaLimiteCli] BIT DEFAULT 0, 
  [UtilizarCodFornec] BIT DEFAULT 0, 
  [ExibirFabricante] BIT DEFAULT 0, 
  [AlterVendedorCliFor] BIT DEFAULT 0, 
  [VR_Tela_CheckOut] BIT DEFAULT 0, 
  [ImprimeTicketMovEfetivada] BIT DEFAULT 0, 
  [ExigeSenhaGerReimpTicket] BIT DEFAULT 0, 
  [NumeroUltMapaECF] INT DEFAULT 0, 
  [ConsiderarSaldoAnterior] BIT DEFAULT 0, 
  [ExibeCFOP] BIT DEFAULT 0, 
  [VendedorSenhaGerente] BIT DEFAULT 0, 
  [MantemInformacaoUltimaNotaFiscal] BIT DEFAULT 0, 
  [ExigeSenhaGerVndContaAtraso] BIT DEFAULT 0, 
  [ImprimeNotaMovEfetivada] BIT DEFAULT 0, 
  [NaoPermiteDuplicarCNPJ] BIT DEFAULT 0, 
  [NSU] FLOAT DEFAULT 0, 
  [ValorIsencaoPisCofinsCsll] FLOAT DEFAULT 0, 
  [AliquotaAprovCreditoIcms] FLOAT DEFAULT 0, 
  [InscricaoMunicipal] NVARCHAR(20), 
  [CNAE] NVARCHAR(10), 
  [EnderecoNumero] NVARCHAR(10), 
  [EnderecoComplemento] NVARCHAR(30), 
  [Pais] NVARCHAR(60), 
  [InscricaoSuframa] NVARCHAR(9), 
  [AmbienteNFe] SMALLINT DEFAULT 0, 
  [FormatoImpressaoDanfeNfe] SMALLINT DEFAULT 0, 
  [ModDetBaseCalculoIcms] SMALLINT DEFAULT 0, 
  [ModDetBaseCalculoIcmsSt] SMALLINT DEFAULT 0, 
  [PastaEnvioNfe] NVARCHAR(255), 
  [PastaRetornoNfe] NVARCHAR(255), 
  [HabilitarNotaFiscalEletronica] BIT DEFAULT 0, 
  [UltimaNFe] INT, 
  [CEP] NVARCHAR(10), 
  [VersaoLayoutEnvio] NVARCHAR(6), 
  [CodigoRegimeTributario] SMALLINT DEFAULT 0, 
  [PercentualSimplesNacional] FLOAT DEFAULT 0, 
  [PercentualReducaoBCSimplesNacional] FLOAT DEFAULT 0, 
  [PadraoArquivoIntegracao] NVARCHAR(6), 
  [VRUtilizarTicketModoRelatorio] BIT DEFAULT 0, 
  [FiltrarProdutosInativos] BIT DEFAULT 0, 
  [TrabalharComComanda] BIT DEFAULT 0, 
  [BancoPDV] NVARCHAR(255), 
  [UltimaNFCe] INT DEFAULT 0, 
  [NrUltimaMovimentacaoTemp] INT DEFAULT 0, 
  [participaProgramaFidelidade] INT DEFAULT 0, 
  [TipoSituacaoTributariaPIS] INT DEFAULT 0, 
  [Quick_viaRDP] INT DEFAULT 0, 
  [NumCasasDecimais] INT DEFAULT 0, 
  [CobrarMultaAposVencimentoParcela] BIT DEFAULT 0, 
  [TaxaMultaParcelaVencida] REAL DEFAULT 0, 
  [MultaDiasAposParcelaVencida] REAL DEFAULT 0, 
  [Transf_OpEntrada] INT DEFAULT 0, 
  [Transf_OpSaida] INT DEFAULT 0, 
  [Transf_TabelaPrecos] NVARCHAR(15), 
  [Quick_viaRDP_ticket] INT DEFAULT 0, 
  [VR_OcultaOrc] BIT DEFAULT 0, 
  [comPrestServ] BIT DEFAULT 0, 
  PRIMARY KEY ([Filial])
)

--
-- Table structure for table 'ParametroVLC'
--

IF object_id(N'ParametroVLC', 'U') IS NOT NULL DROP TABLE [ParametroVLC]

CREATE TABLE [ParametroVLC] (
  [dadoValida] NVARCHAR(50), 
  [dadoConteudo] NVARCHAR(50), 
  [dadoV2] NVARCHAR(50), 
  [managerNFEXML] NVARCHAR(50), 
  [estrategicoREL] NVARCHAR(50)
)

--
-- Table structure for table 'ParametroVLCM'
--

IF object_id(N'ParametroVLCM', 'U') IS NOT NULL DROP TABLE [ParametroVLCM]

CREATE TABLE [ParametroVLCM] (
  [dadoChaveDia] NVARCHAR(100)
)

--
-- Table structure for table 'ParamFaturameAuto'
--

IF object_id(N'ParamFaturameAuto', 'U') IS NOT NULL DROP TABLE [ParamFaturameAuto]

CREATE TABLE [ParamFaturameAuto] (
  [Filial] SMALLINT NOT NULL, 
  [Operacao] INT, 
  [Servico] INT, 
  [Caixa] SMALLINT, 
  [Tabela] NVARCHAR(15), 
  [ISS] INT, 
  PRIMARY KEY ([Filial])
)

--
-- Table structure for table 'Pesquisa 1'
--

IF object_id(N'Pesquisa 1', 'U') IS NOT NULL DROP TABLE [Pesquisa 1]

CREATE TABLE [Pesquisa 1] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [Nome] NVARCHAR(80), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Pesquisa 2'
--

IF object_id(N'Pesquisa 2', 'U') IS NOT NULL DROP TABLE [Pesquisa 2]

CREATE TABLE [Pesquisa 2] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [Nome] NVARCHAR(80), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Pesquisa 3'
--

IF object_id(N'Pesquisa 3', 'U') IS NOT NULL DROP TABLE [Pesquisa 3]

CREATE TABLE [Pesquisa 3] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [Nome] NVARCHAR(80), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Preços'
--

IF object_id(N'Preços', 'U') IS NOT NULL DROP TABLE [Preços]

CREATE TABLE [Preços] (
  [Tabela] NVARCHAR(15) NOT NULL, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Preço] REAL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Tabela], [Produto])
)

--
-- Table structure for table 'Preços - Tempo'
--

IF object_id(N'Preços - Tempo', 'U') IS NOT NULL DROP TABLE [Preços - Tempo]

CREATE TABLE [Preços - Tempo] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Classe] INT DEFAULT 0, 
  [SubClasse] INT DEFAULT 0, 
  [Preço 1] REAL DEFAULT 0, 
  [Preço 2] REAL DEFAULT 0, 
  [Preço 3] REAL DEFAULT 0, 
  [Preço 4] REAL DEFAULT 0, 
  [Preço 5] REAL DEFAULT 0, 
  [Preço 6] REAL DEFAULT 0, 
  [PreçoNacional 1] FLOAT DEFAULT 0, 
  [PreçoNacional 2] FLOAT DEFAULT 0, 
  [PreçoNacional 3] FLOAT DEFAULT 0, 
  [PreçoNacional 4] FLOAT DEFAULT 0, 
  [PreçoNacional 5] FLOAT DEFAULT 0, 
  [PreçoNacional 6] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Produto])
)

--
-- Table structure for table 'PrestacaoContas'
--

IF object_id(N'PrestacaoContas', 'U') IS NOT NULL DROP TABLE [PrestacaoContas]

CREATE TABLE [PrestacaoContas] (
  [Filial] SMALLINT, 
  [Fornecedor] INT, 
  [Sequencia] INT, 
  [Linha] SMALLINT, 
  [Produto] NVARCHAR(20), 
  [Custo] FLOAT, 
  [QtdeOriginal] FLOAT, 
  [QtdeDevolvida] FLOAT, 
  [QtdeVendida] FLOAT, 
  [QtdeComprada] FLOAT, 
  [DatadaGeracao] DATETIME2, 
  [Finalizado] BIT, 
  [DatadaFinalizacao] DATETIME2, 
  [ImpressoNF] BIT, 
  [Resultado] SMALLINT, 
  [PrestacaoFechada] BIT, 
  [CompraFechada] BIT, 
  [PeriodoVenda] DATETIME2, 
  [NotaFiscal] INT, 
  [QtdeAcertada] FLOAT DEFAULT 0
)

--
-- Table structure for table 'ProdutoCesta'
--

IF object_id(N'ProdutoCesta', 'U') IS NOT NULL DROP TABLE [ProdutoCesta]

CREATE TABLE [ProdutoCesta] (
  [CodigoCesta] NVARCHAR(20) NOT NULL, 
  [CodigoItem] NVARCHAR(20) NOT NULL, 
  [QuantidadeItem] INT DEFAULT 0, 
  PRIMARY KEY ([CodigoCesta], [CodigoItem])
)

--
-- Table structure for table 'ProdutoCFOP'
--

IF object_id(N'ProdutoCFOP', 'U') IS NOT NULL DROP TABLE [ProdutoCFOP]

CREATE TABLE [ProdutoCFOP] (
  [CodProduto] NVARCHAR(20) NOT NULL, 
  [CodOperacao] INT NOT NULL, 
  [CFOP] NVARCHAR(14), 
  [CSO] NVARCHAR(3), 
  PRIMARY KEY ([CodProduto], [CodOperacao])
)

--
-- Table structure for table 'ProdutoFavoritos'
--

IF object_id(N'ProdutoFavoritos', 'U') IS NOT NULL DROP TABLE [ProdutoFavoritos]

CREATE TABLE [ProdutoFavoritos] (
  [Filial] INT, 
  [Produto] NVARCHAR(20)
)

--
-- Table structure for table 'ProdutoPareamentoFornecedor'
--

IF object_id(N'ProdutoPareamentoFornecedor', 'U') IS NOT NULL DROP TABLE [ProdutoPareamentoFornecedor]

CREATE TABLE [ProdutoPareamentoFornecedor] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tipo] NVARCHAR(1), 
  [ProdutoForn] NVARCHAR(20) NOT NULL, 
  [Fornecedor] INT NOT NULL, 
  PRIMARY KEY ([Produto], [ProdutoForn], [Fornecedor])
)

--
-- Table structure for table 'Produtos'
--

IF object_id(N'Produtos', 'U') IS NOT NULL DROP TABLE [Produtos]

CREATE TABLE [Produtos] (
  [Código] NVARCHAR(20) NOT NULL, 
  [Código Ordenação] NVARCHAR(20), 
  [Tipo] NVARCHAR(1), 
  [Fracionado] BIT, 
  [Classe] INT DEFAULT 0, 
  [Sub Classe] INT DEFAULT 0, 
  [Moeda] SMALLINT DEFAULT 0, 
  [Estoque] BIT, 
  [Unidade Venda] NVARCHAR(5), 
  [Código do Fornecedor] NVARCHAR(15), 
  [Estoque Ideal] REAL DEFAULT 0, 
  [Estoque Mínimo] REAL DEFAULT 0, 
  [Localização] NVARCHAR(15), 
  [Desconto] REAL DEFAULT 0, 
  [Desconto Máximo] REAL DEFAULT 0, 
  [Obs] NVARCHAR(MAX), 
  [Percentual ICM] REAL DEFAULT 0, 
  [Percentual IPI] REAL DEFAULT 0, 
  [Desativado] BIT, 
  [Fabricante] NVARCHAR(15), 
  [Pesquisa 1] INT DEFAULT 0, 
  [Pesquisa 2] INT DEFAULT 0, 
  [Pesquisa 3] INT DEFAULT 0, 
  [Comissão Sobrepõe] BIT, 
  [Comissão] REAL DEFAULT 0, 
  [Última Compra] NVARCHAR(10), 
  [Último Custo] REAL DEFAULT 0, 
  [Custo Médio] REAL DEFAULT 0, 
  [Último Fornecedor] INT DEFAULT 0, 
  [Tipo ICM] NVARCHAR(1), 
  [Base Cálculo] FLOAT DEFAULT 0, 
  [Redução ICM] FLOAT DEFAULT 0, 
  [Foto] NVARCHAR(120), 
  [Não Incluir na Tabela] BIT, 
  [Custo Preço Valor] FLOAT DEFAULT 0, 
  [Custo Desconto Perc] FLOAT DEFAULT 0, 
  [Custo Desconto Valor] FLOAT DEFAULT 0, 
  [Custo Desconto Fixo] NVARCHAR(1), 
  [Custo Frete Perc] FLOAT DEFAULT 0, 
  [Custo Frete Valor] FLOAT DEFAULT 0, 
  [Custo Frete Fixo] NVARCHAR(1), 
  [Custo ICM Compra Perc] FLOAT DEFAULT 0, 
  [Custo ICM Compra Valor] FLOAT DEFAULT 0, 
  [Custo ICM Compra Fixo] NVARCHAR(1), 
  [Custo IPI Compra Perc] FLOAT DEFAULT 0, 
  [Custo IPI Compra Valor] FLOAT DEFAULT 0, 
  [Custo IPI Compra Fixo] NVARCHAR(1), 
  [Custo Custo Finan Perc] FLOAT DEFAULT 0, 
  [Custo Custo Finan Valor] FLOAT DEFAULT 0, 
  [Custo Custo Finan Fixo] NVARCHAR(1), 
  [Custo Outros Compra Perc] FLOAT DEFAULT 0, 
  [Custo Outros Compra Valor] FLOAT DEFAULT 0, 
  [Custo Outros Compra Fixo] NVARCHAR(1), 
  [Custo Perc Compra Sem] FLOAT DEFAULT 0, 
  [Custo Custo Calculado] FLOAT DEFAULT 0, 
  [Custo Preço Venda] FLOAT DEFAULT 0, 
  [Custo ICM Venda Perc] FLOAT DEFAULT 0, 
  [Custo ICM Venda Valor] FLOAT DEFAULT 0, 
  [Custo ICM Venda Fixo] NVARCHAR(1), 
  [Custo IPI Venda Perc] FLOAT DEFAULT 0, 
  [Custo IPI Venda Valor] FLOAT DEFAULT 0, 
  [Custo IPI Venda Fixo] NVARCHAR(1), 
  [Custo Impostos Perc] FLOAT DEFAULT 0, 
  [Custo Impostos Valor] FLOAT DEFAULT 0, 
  [Custo Impostos Fixo] NVARCHAR(1), 
  [Custo Outros Venda Perc] FLOAT DEFAULT 0, 
  [Custo Outros Venda Valor] FLOAT DEFAULT 0, 
  [Custo Outros Venda Fixo] NVARCHAR(1), 
  [Custo Perc Venda Sem] FLOAT DEFAULT 0, 
  [Custo Lucro Perc] FLOAT DEFAULT 0, 
  [Custo Lucro Valor] FLOAT DEFAULT 0, 
  [Custo Manter] NVARCHAR(1), 
  [Data Alteração] NVARCHAR(10), 
  [QtdeCasasDecimais] INT, 
  [PesoLiquido] REAL, 
  [PesoBruto] REAL, 
  [Percentual Icm Entrada] REAL, 
  [Percentual Icm Saida] REAL, 
  [WebIncluded] BIT, 
  [WebSynchronize] BIT, 
  [WebLastOp] NVARCHAR(1), 
  [WebBonus] INT, 
  [WebOfferDateStart] DATETIME2, 
  [WebOfferDateEnd] DATETIME2, 
  [WebOfferTablePrice] NVARCHAR(15), 
  [WebSaleTablePrice] NVARCHAR(15), 
  [WebAttribFabricante] BIT, 
  [WebAttribPesquisa123] BIT, 
  [Nome] NVARCHAR(80), 
  [Situação Tributária] NVARCHAR(4), 
  [Nome Nota] NVARCHAR(80), 
  [DontAllowDesc] BIT, 
  [UsaDescrAdic] BIT, 
  [IndiceFinanceiro] FLOAT DEFAULT 0, 
  [Volumagem] INT DEFAULT 0, 
  [Cubagem] FLOAT DEFAULT 0, 
  [Lote] NVARCHAR(15), 
  [DataValidade] DATETIME2, 
  [ImprimirUmaEtiq] BIT DEFAULT 0, 
  [ImprimirPrecoEtiq] BIT DEFAULT 0, 
  [CodigoNBM] NVARCHAR(8), 
  [ConsumoDeTecido] FLOAT DEFAULT 0, 
  [PrecoDoMetroLinear] FLOAT DEFAULT 0, 
  [CustoDoTecido] FLOAT DEFAULT 0, 
  [VlMaoDeObraFaccao] FLOAT DEFAULT 0, 
  [VlLavanderia] FLOAT DEFAULT 0, 
  [VlBordado] FLOAT DEFAULT 0, 
  [VlEstamparia] FLOAT DEFAULT 0, 
  [VlAviamentos] FLOAT DEFAULT 0, 
  [OutrosCustos] FLOAT DEFAULT 0, 
  [IndicePrecoEntrada] FLOAT DEFAULT 0, 
  [GrupoFiscal] INT DEFAULT 0, 
  [IndiceIcmsRetido] FLOAT DEFAULT 0, 
  [DividirPrecoEtiqueta] INT DEFAULT 0, 
  [EspacoFisicoTotal] FLOAT DEFAULT 0, 
  [CSO] NVARCHAR(3), 
  [IPI_Reduzido] BIT DEFAULT 0, 
  [Classificação Fiscal] INT, 
  [AliqNCM] FLOAT DEFAULT 0, 
  [MotivoDesoneracaoICMS] INT DEFAULT 0, 
  [TipoSituacaoTributariaPIS] INT DEFAULT 0, 
  [CodigoBeneficio] NVARCHAR(10), 
  [Percentual_IPI_Entrada] REAL DEFAULT 0, 
  [BaseCalculoICMSST_Saida] FLOAT DEFAULT 0, 
  [BaseCalculoICMSST_Entrada] FLOAT DEFAULT 0, 
  [Percentual_ICMSST_Entrada] REAL DEFAULT 0, 
  [Percentual_ICMSST_Saida] REAL DEFAULT 0, 
  [SituacaoTributariaEntrada] NVARCHAR(4), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Programacao'
--

IF object_id(N'Programacao', 'U') IS NOT NULL DROP TABLE [Programacao]

CREATE TABLE [Programacao] (
  [Num Autorizacao] INT NOT NULL, 
  [MesX] INT NOT NULL, 
  [Programacao] NVARCHAR(25), 
  [Dia 01] INT, 
  [Dia 02] INT, 
  [Dia 03] INT, 
  [Dia 04] INT, 
  [Dia 05] INT, 
  [Dia 06] INT, 
  [Dia 07] INT, 
  [Dia 08] INT, 
  [Dia 09] INT, 
  [Dia 10] INT, 
  [Dia 11] INT, 
  [Dia 12] INT, 
  [Dia 13] INT, 
  [Dia 14] INT, 
  [Dia 15] INT, 
  [Dia 16] INT, 
  [Dia 17] INT, 
  [Dia 18] INT, 
  [Dia 19] INT, 
  [Dia 20] INT, 
  [Dia 21] INT, 
  [Dia 22] INT, 
  [Dia 23] INT, 
  [Dia 24] INT, 
  [Dia 25] INT, 
  [Dia 26] INT, 
  [Dia 27] INT, 
  [Dia 28] INT, 
  [Dia 29] INT, 
  [Dia 30] INT, 
  [Dia 31] INT, 
  [Total de Insercoes] FLOAT, 
  [Valor Unitario] FLOAT, 
  [Valor Total] FLOAT, 
  [Periodo Ini] DATETIME2, 
  [Periodo Fin] DATETIME2, 
  [Faixa Ini] NVARCHAR(7), 
  [Faixa Fin] NVARCHAR(7), 
  [Frequencia] NVARCHAR(3), 
  [Duracao] NVARCHAR(5), 
  [Mes] NVARCHAR(3), 
  [Condicoes Pagamento] NVARCHAR(30), 
  [Gerar Etiqueta] BIT, 
  [Cancela Contrato] BIT, 
  [Faturado] BIT, 
  [Valor1] FLOAT, 
  [Valor2] FLOAT, 
  [Valor3] FLOAT, 
  [Valor4] FLOAT, 
  [Vencimento1] DATETIME2, 
  [Vencimento2] DATETIME2, 
  [Vencimento3] DATETIME2, 
  [Vencimento4] DATETIME2, 
  [Status1] BIT, 
  [Status2] BIT, 
  [Status3] BIT, 
  [Status4] BIT, 
  [Cancel1] BIT, 
  [Cancel2] BIT, 
  [Cancel3] BIT, 
  [Cancel4] BIT, 
  [ImpressoNF] BIT, 
  [SomaCancelamento] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Num Autorizacao], [MesX])
)

--
-- Table structure for table 'Radio'
--

IF object_id(N'Radio', 'U') IS NOT NULL DROP TABLE [Radio]

CREATE TABLE [Radio] (
  [Código] INT NOT NULL, 
  [Nome] NVARCHAR(50), 
  [Endereco] NVARCHAR(50), 
  [Cidade] NVARCHAR(30), 
  [Estado] NVARCHAR(2), 
  [CNPJ] NVARCHAR(20), 
  [Inscricao] NVARCHAR(20), 
  [Telefone] NVARCHAR(20), 
  [Contatos] NVARCHAR(40), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Ref_CEST_NCM'
--

IF object_id(N'Ref_CEST_NCM', 'U') IS NOT NULL DROP TABLE [Ref_CEST_NCM]

CREATE TABLE [Ref_CEST_NCM] (
  [cest] NVARCHAR(8), 
  [ncm] NVARCHAR(7)
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
-- Table structure for table 'Resumo Clientes'
--

IF object_id(N'Resumo Clientes', 'U') IS NOT NULL DROP TABLE [Resumo Clientes]

CREATE TABLE [Resumo Clientes] (
  [Dia] DATETIME2 NOT NULL, 
  [Cliente] INT NOT NULL, 
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Qtde] FLOAT, 
  [Valor Total] FLOAT, 
  [Filial] SMALLINT DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Tipo] NVARCHAR(1), 
  [Descricao Adicional] NVARCHAR(MAX), 
  PRIMARY KEY ([Cliente], [Dia], [Produto], [Tamanho], [Cor], [Edição], [Sequência])
)

--
-- Table structure for table 'Resumo Diário'
--

IF object_id(N'Resumo Diário', 'U') IS NOT NULL DROP TABLE [Resumo Diário]

CREATE TABLE [Resumo Diário] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Valor Vendas] FLOAT DEFAULT 0, 
  [Valor Serviços] FLOAT DEFAULT 0, 
  [Valor Compras] FLOAT DEFAULT 0, 
  [Valor T Saída] FLOAT DEFAULT 0, 
  [Valor T Entrada] FLOAT DEFAULT 0, 
  [Valor A Saída] FLOAT DEFAULT 0, 
  [Valor A Entrada] FLOAT DEFAULT 0, 
  [Valor G Saída] FLOAT DEFAULT 0, 
  [Valor G Entrada] FLOAT DEFAULT 0, 
  [Valor Quebras] FLOAT DEFAULT 0, 
  [Valor E Saída] FLOAT DEFAULT 0, 
  [Valor E Entrada] FLOAT DEFAULT 0, 
  [Valor Devolução] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Data])
)

--
-- Table structure for table 'Resumo Diário Financeiro'
--

IF object_id(N'Resumo Diário Financeiro', 'U') IS NOT NULL DROP TABLE [Resumo Diário Financeiro]

CREATE TABLE [Resumo Diário Financeiro] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2 NOT NULL, 
  [Valor Vendas] FLOAT DEFAULT 0, 
  [Valor Serviços] FLOAT DEFAULT 0, 
  [Valor Compras] FLOAT DEFAULT 0, 
  [Valor T Saída] FLOAT DEFAULT 0, 
  [Valor T Entrada] FLOAT DEFAULT 0, 
  [Valor A Saída] FLOAT DEFAULT 0, 
  [Valor A Entrada] FLOAT DEFAULT 0, 
  [Valor G Saída] FLOAT DEFAULT 0, 
  [Valor G Entrada] FLOAT DEFAULT 0, 
  [Valor Quebras] FLOAT DEFAULT 0, 
  [Valor E Saída] FLOAT DEFAULT 0, 
  [Valor E Entrada] FLOAT DEFAULT 0, 
  [Valor Devolução] FLOAT DEFAULT 0, 
  PRIMARY KEY ([Filial], [Data])
)

--
-- Table structure for table 'Retencao'
--

IF object_id(N'Retencao', 'U') IS NOT NULL DROP TABLE [Retencao]

CREATE TABLE [Retencao] (
  [Código] INT NOT NULL, 
  [Nome] NVARCHAR(50), 
  [NomeDaFinanceira] NVARCHAR(10), 
  [ValorRetencao] FLOAT, 
  [Tipo] NVARCHAR(16), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Saídas'
--

IF object_id(N'Saídas', 'U') IS NOT NULL DROP TABLE [Saídas]

CREATE TABLE [Saídas] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Data] DATETIME2, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Operação] INT DEFAULT 0, 
  [Caixa] SMALLINT DEFAULT 0, 
  [Tabela] NVARCHAR(15), 
  [Digitador] INT DEFAULT 0, 
  [Operador] INT DEFAULT 0, 
  [Cliente] INT DEFAULT 0, 
  [Observações] NVARCHAR(70), 
  [Produtos] FLOAT DEFAULT 0, 
  [Desconto] FLOAT DEFAULT 0, 
  [Serviços] FLOAT, 
  [Base ISS] FLOAT DEFAULT 0, 
  [Valor ISS] FLOAT DEFAULT 0, 
  [Perc IR Sobre ISS] REAL DEFAULT 0, 
  [Valor IR Sobre ISS] FLOAT DEFAULT 0, 
  [IPI] FLOAT DEFAULT 0, 
  [Frete] FLOAT DEFAULT 0, 
  [Base ICM] FLOAT DEFAULT 0, 
  [Valor ICM] FLOAT DEFAULT 0, 
  [Base ICM Subs] FLOAT DEFAULT 0, 
  [Valor ICM Subs] FLOAT DEFAULT 0, 
  [Total] FLOAT DEFAULT 0, 
  [Efetivada] BIT, 
  [Recebimento] BIT, 
  [Nota Impressa] INT DEFAULT 0, 
  [Recebe - Conta] BIT, 
  [Recebe - Dinheiro] FLOAT DEFAULT 0, 
  [Recebe - Emp Cartão] INT DEFAULT 0, 
  [Recebe - Num Cartão] NVARCHAR(20), 
  [Recebe - Cartão] FLOAT DEFAULT 0, 
  [Recebe - Vale] FLOAT DEFAULT 0, 
  [Referência] NVARCHAR(20), 
  [Nota Cancelada] BIT DEFAULT 0, 
  [Movimentação Desfeita] BIT, 
  [Total Vista] FLOAT, 
  [Total Prazo] FLOAT, 
  [Tipo Parcela] NVARCHAR(1), 
  [Conta] SMALLINT DEFAULT 0, 
  [Prometido Para] NVARCHAR(50), 
  [Orçamento Aprovado] NVARCHAR(50), 
  [Data Acerto Empréstimo] DATETIME2, 
  [Técnico] INT DEFAULT 0, 
  [Cupom Fiscal Impresso] BIT, 
  [Parcela Cartão] NVARCHAR(1), 
  [Qtde Parcelas] SMALLINT DEFAULT 0, 
  [Valor Parcela] FLOAT DEFAULT 0, 
  [WebOrderFormID] INT, 
  [InfoNrOrcamento] NVARCHAR(255), 
  [DescontoSubTotal] MONEY, 
  [SequênciaPai] INT, 
  [Locked] BIT, 
  [DataEmissaoNota] DATETIME2, 
  [ConsignacaoMestre] INT, 
  [obs_Transportadora] NVARCHAR(50), 
  [obs_Placa] NVARCHAR(8), 
  [obs_Uf] NVARCHAR(2), 
  [obs_Qtde] NVARCHAR(10), 
  [obs_Especie] NVARCHAR(10), 
  [obs_Marca] NVARCHAR(10), 
  [obs_PesoLiquido] FLOAT, 
  [obs_PesoBruto] FLOAT, 
  [obs_FretePago] SMALLINT, 
  [OrcamentoAprovado] BIT, 
  [ComentariosSobreOrcamento] NVARCHAR(MAX), 
  [Valor Recebido] FLOAT DEFAULT 0, 
  [Troco] FLOAT DEFAULT 0, 
  [Percentual CSLL] REAL DEFAULT 0, 
  [Percentual COFINS] REAL DEFAULT 0, 
  [Percentual PIS] REAL DEFAULT 0, 
  [Percentual IRRF] REAL DEFAULT 0, 
  [Data Validade] DATETIME2, 
  [Num Autorizacao] INT, 
  [MesX] INT DEFAULT 0, 
  [Total CSLL] FLOAT DEFAULT 0, 
  [Total COFINS] FLOAT DEFAULT 0, 
  [Total PIS] FLOAT DEFAULT 0, 
  [Total IRRF] FLOAT DEFAULT 0, 
  [FaturaSourceReserva] BIT, 
  [TotalMenosServ] FLOAT DEFAULT 0, 
  [Codigo Func Comprador] INT DEFAULT 0, 
  [Status Venda Func] BIT, 
  [CodigoRetencao] INT DEFAULT 0, 
  [Seguro] FLOAT DEFAULT 0, 
  [Nota Fiscal] INT, 
  [SerieNF] NVARCHAR(3), 
  [InfoICMSporUF] BIT DEFAULT 0, 
  [DataEmissaoNotaManual] DATETIME2, 
  [Ticket Impresso] BIT DEFAULT 0, 
  [NSU] FLOAT DEFAULT 0, 
  [NSU_Data] DATETIME2, 
  [NSU_Hora] DATETIME2, 
  [NumeroDocumentoCliente] NVARCHAR(20), 
  [ModeloDocumentoFiscal] NVARCHAR(2), 
  [TotalNCM] FLOAT DEFAULT 0, 
  [Total_Desp_Acessorias] FLOAT DEFAULT 0, 
  [Consumidor_Final] INT DEFAULT 0, 
  [Presenca_Comprador] INT DEFAULT 0, 
  [TotalDesoneracaoICMS] FLOAT DEFAULT 0, 
  [FinalidadeNFe] INT DEFAULT 0, 
  [ChaveReferenciada] NVARCHAR(100), 
  [NFCe] INT DEFAULT 0, 
  [TotalCartaoDebito] FLOAT DEFAULT 0, 
  [TotalCartaoCredito] FLOAT DEFAULT 0, 
  [ChaveNFCe] NVARCHAR(255), 
  [CPF_CPNJ_Cliente] NVARCHAR(255), 
  [Emitiu_Dados_Cliente_NFCe] BIT DEFAULT 0, 
  [Emitiu_Somente_CPF_Cliente_NFCe] BIT DEFAULT 0, 
  [Nome_Cliente_NFCe] NVARCHAR(255), 
  [aliquota_origem] NVARCHAR(100), 
  [aliquota_destino] NVARCHAR(100), 
  [FreteSomaOuNaoEstimativa] BIT, 
  [obs_infCpl1] NVARCHAR(255), 
  [obs_infCpl2] NVARCHAR(255), 
  [retNFCe] NVARCHAR(MAX), 
  [NFCe_contingencia_num] INT DEFAULT 0, 
  [NFCe_contingencia_serie] INT DEFAULT 0, 
  [NFCe_contingencia_status] NVARCHAR(30), 
  [retNFCe_contingencia] NVARCHAR(MAX), 
  [NFCe_contingencia_chave] NVARCHAR(50), 
  [PrestadorServico] INT, 
  PRIMARY KEY ([Filial], [Sequência])
)

--
-- Table structure for table 'Saídas - Produtos'
--

IF object_id(N'Saídas - Produtos', 'U') IS NOT NULL DROP TABLE [Saídas - Produtos]

CREATE TABLE [Saídas - Produtos] (
  [Filial] SMALLINT NOT NULL DEFAULT 0, 
  [Sequência] INT NOT NULL DEFAULT 0, 
  [Linha] SMALLINT NOT NULL DEFAULT 0, 
  [Código] NVARCHAR(20), 
  [Qtde] REAL DEFAULT 0, 
  [Preço] REAL DEFAULT 0, 
  [Desconto] REAL DEFAULT 0, 
  [Desconto Valor] REAL DEFAULT 0, 
  [ICM] REAL DEFAULT 0, 
  [IPI] REAL DEFAULT 0, 
  [Preço Final] REAL DEFAULT 0, 
  [Etiqueta] BIT, 
  [Código sem Grade] NVARCHAR(20), 
  [InGeradoViaConsig] BIT, 
  [Unidade Venda] NVARCHAR(5), 
  [Descricao Adicional] NVARCHAR(50), 
  [QtdeEntregue] FLOAT, 
  [NrSerie] NVARCHAR(20), 
  [BaseCalculoICMS] FLOAT, 
  [ValorICMS] FLOAT, 
  [NrCOO] INT, 
  [NrItemFiscal] INT, 
  [SitTribAliqICMS] NVARCHAR(4), 
  [CFOP] NVARCHAR(14), 
  [PrecoCusto] REAL DEFAULT 0, 
  [Desp_Acessorias] FLOAT DEFAULT 0, 
  [ValorICMSDesonerado] FLOAT DEFAULT 0, 
  [MotivoDesoneracaoICMS] INT DEFAULT 0, 
  [Percentual_Diferimento] FLOAT DEFAULT 0, 
  [Valor_Aprox_Impostos] FLOAT DEFAULT 0, 
  [Situação Tributária] NVARCHAR(255), 
  PRIMARY KEY ([Filial], [Sequência], [Linha])
)

--
-- Table structure for table 'Saídas - Serviços'
--

IF object_id(N'Saídas - Serviços', 'U') IS NOT NULL DROP TABLE [Saídas - Serviços]

CREATE TABLE [Saídas - Serviços] (
  [Filial] SMALLINT NOT NULL, 
  [Sequência] INT NOT NULL, 
  [Linha] SMALLINT NOT NULL, 
  [Código] INT DEFAULT 0, 
  [Descrição] NVARCHAR(70), 
  [Tempo] NVARCHAR(10), 
  [Completo] BIT, 
  [Preço] FLOAT, 
  [Técnico] INT DEFAULT 0, 
  [NrSerie] NVARCHAR(20), 
  [BaseCalculoICMS] FLOAT, 
  [ValorICMS] FLOAT, 
  [NrCOO] INT, 
  [NrItemFiscal] INT, 
  [SitTribAliqICMS] NVARCHAR(4), 
  [CST] NVARCHAR(1), 
  [CFOP] NVARCHAR(14), 
  PRIMARY KEY ([Filial], [Sequência], [Linha])
)

--
-- Table structure for table 'SaidasChaves'
--

IF object_id(N'SaidasChaves', 'U') IS NOT NULL DROP TABLE [SaidasChaves]

CREATE TABLE [SaidasChaves] (
  [Filial] INT, 
  [Sequencia] INT, 
  [Chave] NVARCHAR(55)
)

--
-- Table structure for table 'SaidasComandas'
--

IF object_id(N'SaidasComandas', 'U') IS NOT NULL DROP TABLE [SaidasComandas]

CREATE TABLE [SaidasComandas] (
  [CodComanda] NVARCHAR(13), 
  [CodSaida] INT, 
  [Filial] SMALLINT
)

--
-- Table structure for table 'ServicoCFOP'
--

IF object_id(N'ServicoCFOP', 'U') IS NOT NULL DROP TABLE [ServicoCFOP]

CREATE TABLE [ServicoCFOP] (
  [CodServico] INT NOT NULL, 
  [CodOperacao] INT NOT NULL, 
  [CFOP] NVARCHAR(14), 
  PRIMARY KEY ([CodServico], [CodOperacao])
)

--
-- Table structure for table 'Serviços'
--

IF object_id(N'Serviços', 'U') IS NOT NULL DROP TABLE [Serviços]

CREATE TABLE [Serviços] (
  [Código] INT NOT NULL, 
  [Descrição] NVARCHAR(60), 
  [Preço] FLOAT, 
  [Comissão Sobrepõe] BIT, 
  [Comissão] REAL DEFAULT 0, 
  [ISS] REAL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [Publicidade] BIT, 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Sub Classes'
--

IF object_id(N'Sub Classes', 'U') IS NOT NULL DROP TABLE [Sub Classes]

CREATE TABLE [Sub Classes] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(25), 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'Supervisores'
--

IF object_id(N'Supervisores', 'U') IS NOT NULL DROP TABLE [Supervisores]

CREATE TABLE [Supervisores] (
  [Código] INT NOT NULL, 
  [Nome] NVARCHAR(50), 
  [Obs] NVARCHAR(MAX), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'TabCaractCliFor'
--

IF object_id(N'TabCaractCliFor', 'U') IS NOT NULL DROP TABLE [TabCaractCliFor]

CREATE TABLE [TabCaractCliFor] (
  [CodCaract] INT NOT NULL, 
  [TipoCliCaract] NVARCHAR(1) NOT NULL, 
  [DescCaract] NVARCHAR(255), 
  PRIMARY KEY ([CodCaract], [TipoCliCaract])
)

--
-- Table structure for table 'Tabela de Preços'
--

IF object_id(N'Tabela de Preços', 'U') IS NOT NULL DROP TABLE [Tabela de Preços]

CREATE TABLE [Tabela de Preços] (
  [Tabela] NVARCHAR(15) NOT NULL, 
  [Aceita Pré] BIT, 
  [Prazo Pré] INT DEFAULT 0, 
  [Aceita Parcelamento] BIT, 
  [Prazo Parcelamento] INT DEFAULT 0, 
  [Aceita Cartão] BIT, 
  [Aceita Vale] BIT, 
  [Multiplicador Comissão] REAL DEFAULT 0, 
  [Data Alteração] NVARCHAR(10), 
  [Dividir] SMALLINT DEFAULT 0, 
  [PercentualComissaoDesconto] REAL, 
  PRIMARY KEY ([Tabela])
)

--
-- Table structure for table 'Tamanhos'
--

IF object_id(N'Tamanhos', 'U') IS NOT NULL DROP TABLE [Tamanhos]

CREATE TABLE [Tamanhos] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(30), 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'TerritorioMunicipio'
--

IF object_id(N'TerritorioMunicipio', 'U') IS NOT NULL DROP TABLE [TerritorioMunicipio]

CREATE TABLE [TerritorioMunicipio] (
  [Uf] NVARCHAR(2) NOT NULL, 
  [Nome] NVARCHAR(64) NOT NULL, 
  [Nome2] NVARCHAR(64) NOT NULL, 
  [CodigoIbge] INT, 
  PRIMARY KEY ([Uf], [Nome])
)

--
-- Table structure for table 'TerritorioPais'
--

IF object_id(N'TerritorioPais', 'U') IS NOT NULL DROP TABLE [TerritorioPais]

CREATE TABLE [TerritorioPais] (
  [Nome] NVARCHAR(64) NOT NULL, 
  [CodigoBacen] INT, 
  PRIMARY KEY ([Nome])
)

--
-- Table structure for table 'TerritorioUf'
--

IF object_id(N'TerritorioUf', 'U') IS NOT NULL DROP TABLE [TerritorioUf]

CREATE TABLE [TerritorioUf] (
  [Nome] NVARCHAR(64) NOT NULL, 
  [Sigla] NVARCHAR(2) NOT NULL, 
  [CodigoIbge] SMALLINT, 
  PRIMARY KEY ([Sigla])
)

--
-- Table structure for table 'TipoComercial'
--

IF object_id(N'TipoComercial', 'U') IS NOT NULL DROP TABLE [TipoComercial]

CREATE TABLE [TipoComercial] (
  [Código] INT NOT NULL, 
  [Descricao] NVARCHAR(50), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'TransferenciaEntreFiliais'
--

IF object_id(N'TransferenciaEntreFiliais', 'U') IS NOT NULL DROP TABLE [TransferenciaEntreFiliais]

CREATE TABLE [TransferenciaEntreFiliais] (
  [CodigoTransf] INT NOT NULL IDENTITY, 
  [FilialLogada] INT, 
  [FilialExportada] INT, 
  [CodigoFornecedor] INT, 
  [CodigoCliente] INT, 
  [CodigoOperSaida] INT, 
  [CodigoOperEntrada] INT, 
  [TabelaPrecos] NVARCHAR(15), 
  [PermitirTransfEstoqueInsuf] INT, 
  [Data] DATETIME2, 
  [Status] INT, 
  [CodigoUsuario] INT, 
  [QuantidadeItens] INT, 
  [NumItens] INT, 
  PRIMARY KEY ([CodigoTransf])
)

--
-- Table structure for table 'TransferenciaProdutos'
--

IF object_id(N'TransferenciaProdutos', 'U') IS NOT NULL DROP TABLE [TransferenciaProdutos]

CREATE TABLE [TransferenciaProdutos] (
  [CodigoTransf] INT, 
  [codigoProduto] NVARCHAR(20), 
  [nomeProduto] NVARCHAR(100), 
  [Quantidade] INT
)

--
-- Table structure for table 'Transportadoras'
--

IF object_id(N'Transportadoras', 'U') IS NOT NULL DROP TABLE [Transportadoras]

CREATE TABLE [Transportadoras] (
  [Código] INT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(50), 
  [Endereço] NVARCHAR(50), 
  [Cidade] NVARCHAR(30), 
  [Estado] NVARCHAR(2), 
  [CGC] NVARCHAR(20), 
  [Inscrição] NVARCHAR(20), 
  [Telefone] NVARCHAR(40), 
  [Contatos] NVARCHAR(40), 
  [Data Alteração] NVARCHAR(10), 
  PRIMARY KEY ([Código])
)

--
-- Table structure for table 'WEB_ClienteOrigem'
--

IF object_id(N'WEB_ClienteOrigem', 'U') IS NOT NULL DROP TABLE [WEB_ClienteOrigem]

CREATE TABLE [WEB_ClienteOrigem] (
  [ID] NVARCHAR(50) NOT NULL, 
  [Origem] NVARCHAR(255), 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_Config'
--

IF object_id(N'WEB_Config', 'U') IS NOT NULL DROP TABLE [WEB_Config]

CREATE TABLE [WEB_Config] (
  [ID] SMALLINT NOT NULL, 
  [xml] INT, 
  [image] IMAGE, 
  [Filial] SMALLINT, 
  [CNX_User] NVARCHAR(255), 
  [CNX_Password] NVARCHAR(255), 
  [CNX_Store] NVARCHAR(255), 
  [Password] NVARCHAR(255), 
  [CodOpReserva] INT, 
  [CodOpVenda] INT, 
  [CodOpCancelamento] INT, 
  [AllowExp_SemDescricao] BIT, 
  [UnitSaleDefault] NVARCHAR(1), 
  [ExportWithClasseAndSubClasse] BIT, 
  [PassiveMode] BIT, 
  [MoedaReal] INT DEFAULT 0, 
  [MoedaDolar] INT DEFAULT 0, 
  [CNX_ServerAddress] NVARCHAR(255), 
  [last_xml_received] INT DEFAULT 0, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_OrderForms'
--

IF object_id(N'WEB_OrderForms', 'U') IS NOT NULL DROP TABLE [WEB_OrderForms]

CREATE TABLE [WEB_OrderForms] (
  [ID] INT NOT NULL IDENTITY, 
  [Filial] SMALLINT NOT NULL, 
  [Sequencia] INT, 
  [OrderID] NVARCHAR(26) NOT NULL, 
  [Origem] NVARCHAR(1) NOT NULL, 
  [Total] MONEY NOT NULL, 
  [Passo] SMALLINT NOT NULL, 
  [StatusShopper] NVARCHAR(255), 
  [StatusAdmin] NVARCHAR(255), 
  [Data] DATETIME2 NOT NULL, 
  [CodPagamento] SMALLINT NOT NULL, 
  [Boleto] INT NOT NULL, 
  [BonusTotal] INT NOT NULL, 
  [BonusUtilizado] INT NOT NULL, 
  [SubTotal] MONEY NOT NULL, 
  [ShippingMethod] SMALLINT NOT NULL, 
  [ShippingTotal] MONEY NOT NULL, 
  [TraceCode] NVARCHAR(20), 
  [ShopperID] NVARCHAR(32) NOT NULL, 
  [ShipName] NVARCHAR(100) NOT NULL, 
  [ShipAddress] NVARCHAR(200) NOT NULL, 
  [ShipCity] NVARCHAR(50) NOT NULL, 
  [ShipState] NVARCHAR(40) NOT NULL, 
  [ShipZip] NVARCHAR(15) NOT NULL, 
  [ShipCountry] NVARCHAR(50) NOT NULL, 
  [ShipPhone] NVARCHAR(35) NOT NULL, 
  [BillName] NVARCHAR(100) NOT NULL, 
  [BillAddress] NVARCHAR(200) NOT NULL, 
  [BillCity] NVARCHAR(50) NOT NULL, 
  [BillState] NVARCHAR(40) NOT NULL, 
  [BillZip] NVARCHAR(15) NOT NULL, 
  [BillCountry] NVARCHAR(50) NOT NULL, 
  [BillPhone] NVARCHAR(35) NOT NULL, 
  [Comentario] NVARCHAR(255), 
  [NumParcelas] INT, 
  [CCName] NVARCHAR(255), 
  [CCType] NVARCHAR(255), 
  [BancoNum] NVARCHAR(4), 
  [BancoNome] NVARCHAR(255), 
  [Agencia] NVARCHAR(10), 
  [ContaCorrente] NVARCHAR(20), 
  [CPF_CNPJ] NVARCHAR(20), 
  [Titular] NVARCHAR(100), 
  [ShipStreetNumber] NVARCHAR(10), 
  [ShipStreetCompl] NVARCHAR(50), 
  [ShipDistrict] NVARCHAR(50), 
  [ShipDDDPhone] NVARCHAR(7), 
  [BillStreetNumber] NVARCHAR(10), 
  [BillStreetCompl] NVARCHAR(50), 
  [BillDistrict] NVARCHAR(50), 
  [BillDDDPhone] NVARCHAR(7), 
  [Seguro] FLOAT DEFAULT 0, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_OrderItens'
--

IF object_id(N'WEB_OrderItens', 'U') IS NOT NULL DROP TABLE [WEB_OrderItens]

CREATE TABLE [WEB_OrderItens] (
  [ID] INT NOT NULL IDENTITY, 
  [OrderFormID] INT NOT NULL, 
  [sku] NVARCHAR(100) NOT NULL, 
  [Quantity] INT NOT NULL, 
  [ListPrice] MONEY NOT NULL, 
  [Moeda] INT NOT NULL, 
  [Discount] MONEY NOT NULL, 
  [Total] MONEY NOT NULL, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_OrderStatus'
--

IF object_id(N'WEB_OrderStatus', 'U') IS NOT NULL DROP TABLE [WEB_OrderStatus]

CREATE TABLE [WEB_OrderStatus] (
  [ID] SMALLINT NOT NULL, 
  [Name] NVARCHAR(255) NOT NULL, 
  [StatusShopper] NVARCHAR(255) NOT NULL, 
  [StatusAdmin] NVARCHAR(255) NOT NULL, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_OrderStatusHistoric'
--

IF object_id(N'WEB_OrderStatusHistoric', 'U') IS NOT NULL DROP TABLE [WEB_OrderStatusHistoric]

CREATE TABLE [WEB_OrderStatusHistoric] (
  [ID] INT NOT NULL IDENTITY, 
  [OrderFormID] INT NOT NULL, 
  [Passo] SMALLINT NOT NULL, 
  [StatusShopper] NVARCHAR(255), 
  [StatusAdmin] NVARCHAR(255), 
  [Data] DATETIME2 NOT NULL, 
  [WebSynchronize] BIT DEFAULT 1, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_PaymentMethods'
--

IF object_id(N'WEB_PaymentMethods', 'U') IS NOT NULL DROP TABLE [WEB_PaymentMethods]

CREATE TABLE [WEB_PaymentMethods] (
  [ID] SMALLINT NOT NULL, 
  [Name] NVARCHAR(255) NOT NULL, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'WEB_ProdutosExcluir'
--

IF object_id(N'WEB_ProdutosExcluir', 'U') IS NOT NULL DROP TABLE [WEB_ProdutosExcluir]

CREATE TABLE [WEB_ProdutosExcluir] (
  [Codigo] NVARCHAR(20) NOT NULL, 
  PRIMARY KEY ([Codigo])
)

--
-- Table structure for table 'WEB_ShippingMethods'
--

IF object_id(N'WEB_ShippingMethods', 'U') IS NOT NULL DROP TABLE [WEB_ShippingMethods]

CREATE TABLE [WEB_ShippingMethods] (
  [ID] SMALLINT NOT NULL, 
  [Name] NVARCHAR(255) NOT NULL, 
  PRIMARY KEY ([ID])
)

--
-- Table structure for table 'ZZZ'
--

IF object_id(N'ZZZ', 'U') IS NOT NULL DROP TABLE [ZZZ]

CREATE TABLE [ZZZ] (
  [Ordem] SMALLINT NOT NULL DEFAULT 0, 
  [Nome] NVARCHAR(30), 
  [Senha] FLOAT DEFAULT 0, 
  [Registro] FLOAT DEFAULT 0, 
  [Liberação] FLOAT DEFAULT 0, 
  [Senha Master] NVARCHAR(8), 
  [DBVersion] NVARCHAR(10), 
  [CGCCPF] NVARCHAR(30), 
  PRIMARY KEY ([Ordem])
)

--
-- Table structure for table 'ZZZGeral'
--

IF object_id(N'ZZZGeral', 'U') IS NOT NULL DROP TABLE [ZZZGeral]

CREATE TABLE [ZZZGeral] (
  [Linha] INT NOT NULL IDENTITY, 
  [Texto] NVARCHAR(70), 
  [Valor 1] FLOAT DEFAULT 0, 
  [Valor 2] FLOAT DEFAULT 0, 
  [Ordem] INT DEFAULT 0, 
  PRIMARY KEY ([Linha])
)

--
-- Table structure for table 'ZZZGráfico2'
--

IF object_id(N'ZZZGráfico2', 'U') IS NOT NULL DROP TABLE [ZZZGráfico2]

CREATE TABLE [ZZZGráfico2] (
  [Nome] NVARCHAR(20), 
  [Unidades] REAL DEFAULT 0, 
  [Valor Vendas] FLOAT DEFAULT 0
)

--
-- Table structure for table 'ZZZLog'
--

IF object_id(N'ZZZLog', 'U') IS NOT NULL DROP TABLE [ZZZLog]

CREATE TABLE [ZZZLog] (
  [Contador] INT NOT NULL IDENTITY, 
  [Data] DATETIME2, 
  [Texto] NVARCHAR(80), 
  [Tipo] NVARCHAR(20), 
  PRIMARY KEY ([Contador])
)

--
-- Table structure for table 'ZZZProgramas'
--

IF object_id(N'ZZZProgramas', 'U') IS NOT NULL DROP TABLE [ZZZProgramas]

CREATE TABLE [ZZZProgramas] (
  [Nome Programa] NVARCHAR(40) NOT NULL, 
  [Descrição] NVARCHAR(50), 
  [Senha Master] BIT, 
  [Especial] BIT, 
  [Número] INT DEFAULT 0, 
  [ToolID] INT, 
  PRIMARY KEY ([Nome Programa])
)

--
-- Table structure for table 'ZZZVendas'
--

IF object_id(N'ZZZVendas', 'U') IS NOT NULL DROP TABLE [ZZZVendas]

CREATE TABLE [ZZZVendas] (
  [Produto] NVARCHAR(20) NOT NULL, 
  [Tamanho] INT NOT NULL DEFAULT 0, 
  [Cor] INT NOT NULL DEFAULT 0, 
  [Edição] INT NOT NULL DEFAULT 0, 
  [Vendas] FLOAT DEFAULT 0, 
  [Valor Vendas] FLOAT DEFAULT 0, 
  [Classe] INT DEFAULT 0, 
  [Sub Classe] INT DEFAULT 0, 
  [Última Data] NVARCHAR(10), 
  PRIMARY KEY ([Produto], [Tamanho], [Cor], [Edição])
)

