CREATE TABLE [dbo].[ASRSysOrganisationReportFilters](
		[ID] [int] IDENTITY(1,1) NOT NULL,
		[OrganisationID] int NOT NULL,
		[FieldID] int NOT NULL,
		[Operator] [int] NOT NULL,
		[Value] nvarchar(MAX) NOT NULL)