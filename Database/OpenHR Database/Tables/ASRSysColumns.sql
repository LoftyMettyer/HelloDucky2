CREATE TABLE [dbo].[ASRSysColumns](
	[columnID] [int] NOT NULL,
	[tableID] [int] NOT NULL,
	[columnType] [smallint] NOT NULL,
	[datatype] [smallint] NOT NULL,
	[defaultValue] [varchar](max) NULL,
	[size] [int] NULL,
	[decimals] [smallint] NULL,
	[lookupTableID] [int] NULL,
	[lookupColumnID] [int] NULL,
	[controltype] [int] NULL,
	[spinnerMinimum] [smallint] NULL,
	[spinnerMaximum] [smallint] NULL,
	[spinnerIncrement] [smallint] NULL,
	[audit] [bit] NOT NULL,
	[duplicate] [bit] NOT NULL,
	[mandatory] [bit] NOT NULL,
	[uniquecheck] [bit] NOT NULL,
	[convertcase] [smallint] NULL,
	[mask] [varchar](max) NULL,
	[alphaonly] [bit] NOT NULL,
	[blankIfZero] [bit] NOT NULL,
	[multiline] [bit] NOT NULL,
	[alignment] [smallint] NULL,
	[calcExprID] [int] NULL,
	[gotFocusExprID] [int] NULL,
	[lostFocusExprID] [int] NULL,
	[calcTrigger] [smallint] NULL,
	[readOnly] [bit] NOT NULL,
	[statusBarMessage] [varchar](200) NULL,
	[errorMessage] [varchar](max) NULL,
	[linkTableID] [int] NULL,
	[Afdenabled] [bit] NOT NULL,
	[Afdindividual] [bit] NOT NULL,
	[Afdforename] [int] NOT NULL,
	[Afdsurname] [int] NOT NULL,
	[Afdinitial] [int] NOT NULL,
	[Afdtelephone] [int] NOT NULL,
	[Afdaddress] [int] NOT NULL,
	[Afdproperty] [int] NOT NULL,
	[Afdstreet] [int] NOT NULL,
	[Afdlocality] [int] NOT NULL,
	[Afdtown] [int] NOT NULL,
	[Afdcounty] [int] NOT NULL,
	[dfltValueExprID] [int] NULL,
	[linkOrderID] [int] NULL,
	[OleOnServer] [bit] NOT NULL,
	[childUniqueCheck] [bit] NULL,
	[LinkViewID] [int] NULL,
	[DefaultDisplayWidth] [int] NULL,
	[ColumnName] [varchar](128) NULL,
	[UniqueCheckType] [int] NULL,
	[Trimming] [int] NULL,
	[Use1000Separator] [bit] NULL,
	[LookupFilterColumnID] [int] NULL,
	[LookupFilterValueID] [int] NULL,
	[QAddressEnabled] [int] NULL,
	[QAIndividual] [bit] NULL,
	[QAAddress] [int] NULL,
	[QAProperty] [int] NULL,
	[QAStreet] [int] NULL,
	[QALocality] [int] NULL,
	[QATown] [int] NULL,
	[QACounty] [int] NULL,
	[LookupFilterOperator] [int] NULL,
	[Embedded] [int] NULL,
	[OLEType] [int] NULL,
	[MaxOLESizeEnabled] [bit] NULL,
	[MaxOLESize] [int] NULL,
	[AutoUpdateLookupValues] [bit] NULL,
	[CalculateIfEmpty] [bit] NOT NULL,
	[Guid] uniqueidentifier,
	[Locked] bit,
 CONSTRAINT [PK_ASRSysColumns] PRIMARY KEY CLUSTERED 
(
	[columnID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdenabled]  DEFAULT (0) FOR [Afdenabled]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdindividual]  DEFAULT (0) FOR [Afdindividual]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdforename]  DEFAULT (0) FOR [Afdforename]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdsurname]  DEFAULT (0) FOR [Afdsurname]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdinitial]  DEFAULT (0) FOR [Afdinitial]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdtelephone]  DEFAULT (0) FOR [Afdtelephone]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdaddress]  DEFAULT (0) FOR [Afdaddress]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdproperty]  DEFAULT (0) FOR [Afdproperty]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdstreet]  DEFAULT (0) FOR [Afdstreet]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdlocality]  DEFAULT (0) FOR [Afdlocality]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdtown]  DEFAULT (0) FOR [Afdtown]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_Afdcounty]  DEFAULT (0) FOR [Afdcounty]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  DEFAULT (1) FOR [OleOnServer]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_DefaultDisplayWidth]  DEFAULT (1) FOR [DefaultDisplayWidth]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  CONSTRAINT [DF_ASRSysColumns_QAddressEnabled]  DEFAULT (0) FOR [QAddressEnabled]
GO
ALTER TABLE [dbo].[tbsys_columns] ADD  DEFAULT (0) FOR [CalculateIfEmpty]