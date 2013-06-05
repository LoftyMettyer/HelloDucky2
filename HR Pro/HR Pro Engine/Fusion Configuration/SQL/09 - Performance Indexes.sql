
DROP INDEX [IDX_LocalLookup] ON [fusion].[IdTranslation]
GO

CREATE UNIQUE NONCLUSTERED INDEX [IDX_LocalLookup] ON [fusion].[IdTranslation]
(
	[TranslationName] ASC,
	[LocalId] ASC
)
INCLUDE ( [BusRef])

