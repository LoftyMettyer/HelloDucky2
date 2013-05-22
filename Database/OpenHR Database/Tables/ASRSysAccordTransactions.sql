CREATE TABLE [dbo].[ASRSysAccordTransactions](
	[TransactionID] [int] NOT NULL,
	[TransferType] [smallint] NOT NULL,
	[TransactionType] [smallint] NOT NULL,
	[CreatedDateTime] [datetime] NOT NULL,
	[TransferedDateTime] [datetime] NULL,
	[Status] [smallint] NOT NULL,
	[ErrorText] [varchar](max) NULL,
	[CompanyCode] [varchar](255) NULL,
	[EmployeeCode] [varchar](255) NULL,
	[CreatedUser] [varchar](100) NOT NULL,
	[HRProRecordID] [int] NULL,
	[Archived] [bit] NULL,
	[EmployeeName] [varchar](255) NULL,
	[DepartmentCode] [varchar](255) NULL,
	[DepartmentName] [varchar](255) NULL,
	[PayrollCode] [varchar](255) NULL,
	[BatchID] [int] NULL,
 CONSTRAINT [PK_ASRSysAccordTransactions] PRIMARY KEY CLUSTERED 
(
	[TransactionID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]