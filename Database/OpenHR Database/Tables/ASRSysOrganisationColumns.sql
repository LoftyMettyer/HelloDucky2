CREATE TABLE [dbo].[ASRSysOrganisationColumns](
		[ID] [int] IDENTITY(1,1) NOT NULL,
		[OrganisationID] [int] NOT NULL,
      [ViewID] [int] NOT NULL,
		[ColumnID] [int] NOT NULL,
		[Prefix] [varchar](50) NULL,
		[Suffix] [varchar](50) NULL,
		[FontSize] int,
		[Decimals] int,
		[Height] int,
		[ConcatenateWithNext] bit)