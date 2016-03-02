CREATE TYPE [dbo].[SelfServiceType] AS TABLE
(
	[ID] integer NOT NULL,
	[Login] nvarchar(255) NULL,
	[Email] nvarchar(255),
	[StartDate] datetime,
	[LeavingDate] datetime,
	[KnownAs] nvarchar(255),
	[SecurityGroup] nvarchar(255)
)
