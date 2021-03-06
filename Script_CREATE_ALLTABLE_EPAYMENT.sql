
CREATE TABLE [dbo].[FMSLogin](
	[UserName] [nvarchar](50) NULL,
	[Password] [nvarchar](50) NULL
) 


CREATE TABLE [dbo].[FMSEPAYTEMP](
	[ERRCODE] [nvarchar](20) NULL,
	[BATCH] [nvarchar](100) NULL,
	[ENTRYNO] [nvarchar](100) NULL,
	[BeneficiaryName] [nvarchar](100) NULL,
	[TOTALAMOUNT] [nvarchar](100) NULL,
	[Taxbase] [nvarchar](100) NULL,
	[Acct] [nvarchar](100) NULL,
	[TaxId] [nvarchar](100) NULL,
	[RemittoAddress] [nvarchar](100) NULL,
	[OptService] [nvarchar](100) NULL,
	[DATEINV] [nvarchar](20) NULL,
	[BankCode] [nvarchar](20) NULL
)

CREATE TABLE [dbo].[FMSEPayEx](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[RUNNO] [nvarchar](30) NULL,
	[CNTBTCH] [nvarchar](30) NULL,
	[CNTENTY] [nvarchar](30) NULL,
	[EXPORTDATE] [nvarchar](30) NULL
)


CREATE TABLE FMSEPAYBANKMASTER 
(BANKCODE VARCHAR(3),
BANKNAME VARCHAR(140))INSERT INTO FMSEPAYBANKMASTER
(BANKCODE,BANKNAME)
VALUES
('001','BANK OF THAILAND'),
('002','BANGKOK BANK PUBLIC COMPANY LTD.'),
('004','KASIKORNBANK PUBLIC COMPANY LTD.'),
('005','THE ROYAL BANK OF SCOTLAND PLC'),
('006','KRUNG THAI BANK PUBLIC COMPANY LTD.'),
('008','JPMORGAN CHASE BANK, NATIONAL ASSOCIATION'),
('009','OVER SEA-CHINESE BANKING CORPORATION LIMITED'),
('011','TMB BANK PUBLIC COMPANY LIMITED'),
('014','SIAM COMMERCIAL BANK PUBLIC COMPANY LTD.'),
('017','CITIBANK, N.A.'),
('018','SUMITOMO MITSUI BANKING CORPORATION'),
('020','STANDARD CHARTERED BANK (THAI) PUBLIC COMPANY LIMITED'),
('022','CIMB THAI BANK Public Company Limited'),
('023','RHB BANK BERHAD'),
('024','UNITED OVERSEAS BANK (THAI) PUBLIC COMPANY LIMITED'),
('025','BANK OF AYUDHYA PUBLIC COMPANY LTD.'),
('026','MEGA INTERNATIONAL COMMERCIAL BANK PUBLIC COMPANY LIMITED'),
('027','BANK OF AMERICA, NATIONAL ASSOCIATION'),
('029','INDIAN OVERSEA BANK'),
('030','THE GOVERNMENT SAVINGS BANK'),
('031','THE HONGKONG AND SHANGHAI BANKING CORPORATION LTD.'),
('032','DEUTSCHE BANK AG.'),
('033','THE GOVERNMENT HOUSING BANK'),
('034','BANK FOR AGRICULTURE AND AGRICULTURAL COOPERATIVES'),
('035','EXPORT-IMPORT BANK OF THAILAND'),
('039','Mizuho Bank, Ltd. Bangkok Branch'),
('045','BNP PARIBAS'),
('052','BANK OF CHINA (THAI) PUBLIC COMPANY LIMITED'),
('065','THANACHART BANK PUBLIC COMPANY LTD.'),
('066','ISLAMIC BANK OF THAILAND'),
('067','TISCO BANK PUBLIC COMPANY LIMITED'),
('069','KIATNAKIN BANK PUBLIC COMPANY LIMITED'),
('070','INDUSTRIAL AND COMMERCIAL BANK OF CHINA (THAI) PUBLIC COMPANY LIMITED'),
('071','THE THAI CREDIT RETAIL BANK PUBLIC COMPANY LIMITED'),
('073','LAND AND HOUSES BANK PUBLIC COMPANY LIMITED'),
('079','ANZ BANK (THAI) PUBLIC COMPANY LIMITED'),
('080','SUMITOMO MITSUI TRUST BANK (THAI) PUBLIC COMPANY LIMITED'),
('098','SMALL AND MEDIUM ENTERPRISE DEVELOPMENT BANK OF THAILAND')

