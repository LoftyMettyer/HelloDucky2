	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spClearFusionMessageQueue]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spClearFusionMessageQueue];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spClearFusionTranslations]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spClearFusionTranslations];


GO

---------------------------------------------------------------------------------
-- Name:    spClearFusionMessageQueue
--
-- Purpose: Clears the fusion message queue
--
-- Returns: n/a
---------------------------------------------------------------------------------
CREATE PROCEDURE [fusion].[spClearFusionMessageQueue]
	(@MessageType varchar(50))
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @conversationHandle uniqueidentifier;

	SELECT TOP 1 @conversationHandle = conversation_handle FROM sys.conversation_endpoints;
	WHILE @@rowcount = 1
	BEGIN
		END CONVERSATION @conversationHandle WITH CLEANUP;
 		SELECT TOP 1 @conversationHandle = conversation_handle FROM sys.conversation_endpoints;
	END

	-- Queue may well have been poisoned - reenable the queue.
	ALTER QUEUE fusion.qFusion WITH STATUS = ON

END

GO

---------------------------------------------------------------------------------
-- Name:    spClearFusionTranslations
--
-- Purpose: Clears the fusion translations
--
-- Returns: n/a
---------------------------------------------------------------------------------
CREATE PROCEDURE [fusion].[spClearFusionTranslations]
	(@MessageType varchar(50))
AS
BEGIN
	SET NOCOUNT ON;

	DELETE FROM fusion.idTranslation;
	DELETE FROM fusion.messagelog;
	DELETE FROM fusion.messagetracking;

END

GO


--go
--exec fusion.[spClearFusionMessageQueue] ''
--go
--exec fusion.[spClearFusionTranslations] ''