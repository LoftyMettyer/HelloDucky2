CREATE TYPE [dbo].[SelfServiceType] AS TABLE
(
	[Login] nvarchar(255) NULL,
	[Email] nvarchar(255),
	[LeavingDate] datetime,
	[KnownAs] nvarchar(255)
)
