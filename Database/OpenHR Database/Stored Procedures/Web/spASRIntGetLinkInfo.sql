CREATE PROCEDURE [dbo].[spASRIntGetLinkInfo]
(
	@piLinkID 		integer,
	@piScreenID		integer			OUTPUT,
	@piTableID		integer			OUTPUT,
	@psTitle		varchar(MAX)	OUTPUT,
	@piStartMode	integer			OUTPUT, 
	@piTableType	integer			OUTPUT
)
AS
BEGIN
	SELECT 
		@piScreenID = ASRSysSSIntranetLinks.screenID,
		@piTableID = ASRSysScreens.tableID,
		@psTitle = ASRSysSSIntranetLinks.pageTitle,
		@piStartMode = ASRSysSSIntranetLinks.startMode,
		@piTableType = ASRSysTables.TableType
	FROM ASRSysSSIntranetLinks
			INNER JOIN ASRSysScreens 
			ON ASRSysSSIntranetLinks.screenID = ASRSysScreens.screenID
				INNER JOIN ASRSysTables
				ON ASRSysScreens.TableID = ASRSysTables.TableID
	WHERE ID = @piLinkID;
END