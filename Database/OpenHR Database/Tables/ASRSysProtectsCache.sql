CREATE TABLE [dbo].[ASRSysProtectsCache]
(
	[ID] int NOT NULL,
	[Action] tinyint NOT NULL,
	[Columns] varbinary(8000),
	[ProtectType] int NOT NULL,
	[UID] integer NOT NULL
)

GO

CREATE CLUSTERED INDEX [IDX_ProtectsCache_UID] ON [dbo].[ASRSysProtectsCache] ([UID])
