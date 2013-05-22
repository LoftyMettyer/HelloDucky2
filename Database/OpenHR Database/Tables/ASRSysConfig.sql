CREATE TABLE [dbo].[ASRSysConfig](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[SupportTelNo] [varchar](20) NULL,
	[SupportEMail] [varchar](50) NULL,
	[SupportFax] [varchar](20) NULL,
	[URL] [varchar](50) NULL,
	[AuthCode] [varchar](50) NULL,
	[CustNo] [int] NULL,
	[SystemManagerAppName] [varchar](50) NULL,
	[SecurityManagerAppName] [varchar](50) NULL,
	[CustName] [varchar](50) NULL,
	[ModuleCode] [varchar](15) NULL,
	[UserModuleAppName] [varchar](50) NULL,
	[databaseVersion] [int] NULL,
	[refreshStoredProcedures] [bit] NULL,
	[SystemManagerVersion] [varchar](50) NULL,
	[SecurityManagerVersion] [varchar](50) NULL,
	[DataManagerVersion] [varchar](50) NULL,
	[UpdatingDateDependentColumns] [bit] NOT NULL,
	[MinimumPasswordLength] [int] NOT NULL,
	[ChangePasswordFrequency] [int] NOT NULL,
	[ChangePasswordPeriod] [varchar](1) NULL,
	[EmailDateFormat] [varchar](3) NULL,
	[IntranetModuleAppName] [varchar](50) NULL,
	[IntranetVersion] [varchar](50) NULL,
	[IntranetPicturePath] [varchar](2000) NULL,
	[EmailAttachmentsPath] [varchar](255) NULL,
 CONSTRAINT [PK_ASRSysConfig] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysConfig] ADD  CONSTRAINT [DF__ASRSysCon__Updat__4B4950E4]  DEFAULT (0) FOR [UpdatingDateDependentColumns]
GO
ALTER TABLE [dbo].[ASRSysConfig] ADD  CONSTRAINT [DF_ASRSysConfig_MinimumPasswordLength]  DEFAULT (0) FOR [MinimumPasswordLength]
GO
ALTER TABLE [dbo].[ASRSysConfig] ADD  CONSTRAINT [DF_ASRSysConfig_ChangePasswordFrequency]  DEFAULT (0) FOR [ChangePasswordFrequency]