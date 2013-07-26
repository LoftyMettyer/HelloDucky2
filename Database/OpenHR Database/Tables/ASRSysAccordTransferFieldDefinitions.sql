﻿CREATE TABLE [dbo].[ASRSysAccordTransferFieldDefinitions](
	[TransferFieldID] [int] NOT NULL,
	[TransferTypeID] [int] NOT NULL,
	[Mandatory] [bit] NOT NULL,
	[Description] [char](40) NOT NULL,
	[AlwaysTransfer] [bit] NOT NULL,
	[IsKeyField] [bit] NOT NULL,
	[IsCompanyCode] [bit] NOT NULL,
	[IsEmployeeCode] [bit] NOT NULL,
	[Direction] [int] NOT NULL,
	[ASRMapType] [int] NULL,
	[ASRTableID] [int] NULL,
	[ASRColumnID] [int] NULL,
	[ASRExprID] [int] NULL,
	[ASRValue] [char](40) NULL,
	[ConvertData] [bit] NULL,
	[IsEmployeeName] [bit] NULL,
	[IsDepartmentCode] [bit] NULL,
	[IsDepartmentName] [bit] NULL,
	[IsPayrollCode] [bit] NULL,
	[GroupBy] [int] NULL,
	[PreventModify] [bit] NULL
)


GO

CREATE CLUSTERED INDEX [IDX_TransferTypeID]
    ON [dbo].[ASRSysAccordTransferFieldDefinitions]([TransferTypeID] ASC);
GO