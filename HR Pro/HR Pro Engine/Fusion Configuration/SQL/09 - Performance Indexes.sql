
DROP INDEX [NonClusteredIndex-20120907-104049] ON [fusion].[IdTranslation]
GO

CREATE UNIQUE NONCLUSTERED INDEX [IDX_LocalLookup] ON [fusion].[IdTranslation]
(
	[TranslationName] ASC,
	[LocalId] ASC
)
INCLUDE ( [BusRef])

