
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spClearFusionMessageQueue]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [fusion].[spClearFusionMessageQueue];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSendFusionMessageCheckContext]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pSendFusionMessageCheckContext];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSendFusionMessage]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pSendFusionMessage];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pIdTranslateSetBusRef]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pIdTranslateSetBusRef];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pIdTranslateGetLocalId]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pIdTranslateGetLocalId];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pIdTranslateGetBusRef]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pIdTranslateGetBusRef];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageTrackingSetLastGeneratedDate]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate];
	
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageTrackingSetLastProcessedDate]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageTrackingGetLastMessageDates]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates];
	
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageTrackingGetLastGeneratedXml]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml];
	
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageTrackingSetLastGeneratedXml]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageLogAdd]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageLogAdd];
	
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pMessageLogCheck]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pMessageLogCheck];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSetFusionContext]') AND xtype = 'P')
	DROP PROCEDURE [fusion].[pSetFusionContext];
	
	


-- Stored procs that access the above views should be left well alone
EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[spClearFusionMessageQueue]
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

END'


	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[MessageTracking]') AND type in (N'U'))
	BEGIN
		CREATE TABLE [fusion].[MessageTracking](
			[MessageType] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL,
			[BusRef] [uniqueidentifier] NOT NULL,
			[LastGeneratedDate] [datetime] NULL,
			[LastProcessedDate] [datetime] NULL,
			[LastGeneratedXml] [varchar] (max) COLLATE Latin1_General_CI_AS NULL)
	END


	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[MessageLog]') AND type in (N'U'))
	BEGIN
		CREATE TABLE [fusion].[MessageLog](
			[MessageType] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL,
			[MessageRef] [uniqueidentifier] NOT NULL,
			[ReceivedDate] [datetime] NOT NULL,
			[Originator] [varchar] (50) COLLATE Latin1_General_CI_AS NULL)
	END






-- Stop reading this. You do not need to be playing around in this section of the script. Altering the following code is dangerous
-- and should be left to the developers or serious risk takers!


EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pIdTranslateSetBusRef]
	(
		@TranslationName varchar(50),
		@LocalId varchar(25),
		@BusRef uniqueidentifier
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	BEGIN TRAN;
	
	DELETE fusion.IdTranslation WITH (ROWLOCK) 
		WHERE TranslationName = @TranslationName and LocalId = @LocalId;
		
	INSERT fusion.IdTranslation WITH (ROWLOCK) (TranslationName, LocalId, BusRef) 
		VALUES (@TranslationName, @LocalId, @BusRef);

	COMMIT TRAN;
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pSendFusionMessage]
	(
		@MessageType varchar(50),
		@LocalId int
	)
AS
BEGIN
	SET NOCOUNT ON;
	
	DECLARE @DialogHandle uniqueidentifier;
	SET @DialogHandle = NEWID();

	BEGIN DIALOG @DialogHandle 
		FROM SERVICE FusionApplicationService 
		TO SERVICE ''FusionConnectorService''
		ON CONTRACT TriggerFusionContract
		WITH ENCRYPTION = OFF;
		
	DECLARE @msg varchar(max);

	SET @msg = (SELECT	@MessageType AS MessageType, 
						@LocalId as LocalId,
						CONVERT(varchar(50), GETUTCDATE(), 126)+''Z'' as TriggerDate 
					FOR XML PATH(''SendFusionMessage''));	

	SEND ON CONVERSATION @DialogHandle
		MESSAGE TYPE TriggerFusionSend (@msg);
	 
	END CONVERSATION @DialogHandle;

END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pSendFusionMessageCheckContext]
	(
		@MessageType varchar(50),
		@LocalId int
	)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @ContextInfo varbinary(128)
 
	SELECT @ContextInfo = CAST( ''Fusion:''+@MessageType AS VARBINARY(128) );

	IF CONTEXT_INFO() IS NULL OR CONTEXT_INFO() <> @ContextInfo
	BEGIN	
		EXEC fusion.pSendFusionMessage @MessageType, @LocalId
	END
END'


EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pIdTranslateGetLocalId]
	(
		@TranslationName varchar(50),
		@BusRef uniqueidentifier,
		@LocalId varchar(25) output
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	SET @LocalId = null;
	
	SELECT @LocalId = LocalId from [Fusion].IdTranslation WITH (ROWLOCK) 
		WHERE TranslationName = @TranslationName and BusRef = @BusRef;
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pIdTranslateGetBusRef]
	(
		@TranslationName varchar(50),
		@LocalId varchar(25),
		@BusRef uniqueidentifier output,
		@DidGenerate bit = 0 output,
		@CanGenerate bit = 1
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	SET @BusRef = NULL;
	SET @DidGenerate = 0;
	
	SELECT @BusRef = BusRef from [fusion].IdTranslation WITH (ROWLOCK) 
		WHERE TranslationName = @TranslationName AND LocalId = @LocalId;
	
	IF @@ROWCOUNT = 0
	BEGIN
		IF @CanGenerate = 1
		BEGIN
			SET @BusRef = NEWID();

						
			INSERT fusion.IdTranslation WITH (ROWLOCK) (TranslationName, LocalId, BusRef) 
					VALUES (@TranslationName, @LocalId, @BusRef);
			
			SET @DidGenerate = 1;
					
			RETURN 0;
		END
		RETURN 1;
	END

	RETURN 0;
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate]
	(
		@MessageType varchar(50),
		@BusRef uniqueidentifier,
		@LastGeneratedDate datetime
	)

AS
BEGIN
	SET NOCOUNT ON;
		
	IF EXISTS (SELECT * FROM fusion.MessageTracking
			   WHERE MessageType = @MessageType AND BusRef = @BusRef)
	BEGIN	
		UPDATE fusion.MessageTracking
		   SET LastGeneratedDate = @LastGeneratedDate
		   WHERE MessageType = @MessageType AND BusRef = @BusRef
	END
	ELSE
	BEGIN
		INSERT fusion.MessageTracking (MessageType, BusRef, LastGeneratedDate)
			VALUES (@MessageType, @BusRef, @LastGeneratedDate)
	END		
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate]
	(
		@MessageType varchar(50),
		@BusRef uniqueidentifier,
		@LastProcessedDate datetime
	)

AS
BEGIN
	SET NOCOUNT ON;
		
	IF EXISTS (SELECT * FROM fusion.MessageTracking
			   WHERE MessageType = @MessageType AND BusRef = @BusRef)
	BEGIN	
		UPDATE fusion.MessageTracking
		   SET LastProcessedDate = @LastProcessedDate
		   WHERE MessageType = @MessageType AND BusRef = @BusRef
	END
	ELSE
	BEGIN
		INSERT fusion.MessageTracking (MessageType, BusRef, LastProcessedDate)
			VALUES (@MessageType, @BusRef, @LastProcessedDate)
	END		
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates]
	(
		@MessageType varchar(50),
		@BusRef uniqueidentifier
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	SELECT LastProcessedDate, LastGeneratedDate
		FROM fusion.MessageTracking
		WHERE MessageType = @MessageType AND BusRef = @BusRef;

END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml]
	(
		@MessageType varchar(50),
		@BusRef uniqueidentifier
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	SELECT LastGeneratedXml
		FROM fusion.MessageTracking
		WHERE MessageType = @MessageType AND BusRef = @BusRef;

END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml]
	(
		@MessageType varchar(50),
		@BusRef uniqueidentifier,
		@LastGeneratedXml varchar(max)
	)

AS
BEGIN
	SET NOCOUNT ON;
		
	IF EXISTS (SELECT * FROM fusion.MessageTracking
			   WHERE MessageType = @MessageType AND BusRef = @BusRef)
	BEGIN	
		UPDATE fusion.MessageTracking
		   SET LastGeneratedXml = @LastGeneratedXml
		   WHERE MessageType = @MessageType AND BusRef = @BusRef
	END
	ELSE
	BEGIN
		INSERT fusion.MessageTracking (MessageType, BusRef, LastGeneratedXml)
			VALUES (@MessageType, @BusRef, @LastGeneratedXml)
	END		
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageLogAdd]
	(
		@MessageType varchar(50),
		@MessageRef uniqueidentifier,
		@Originator varchar(50) = null
	)

AS
BEGIN
	SET NOCOUNT ON;
		
	INSERT fusion.MessageLog (MessageType, MessageRef, Originator, ReceivedDate) values (@MessageType, @MessageRef, @Originator, GETUTCDATE())
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pMessageLogCheck]
	(
		@MessageType varchar(50),
		@MessageRef uniqueidentifier,
		@ReceivedBefore bit output
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	IF EXISTS ( SELECT * FROM fusion.MessageLog WHERE MessageType = @MessageType AND MessageRef = @MessageRef )
	BEGIN
		SET @ReceivedBefore = 1
	END
	ELSE
	BEGIN
		SET @ReceivedBefore = 0
	END
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [fusion].[pSetFusionContext]
	(
		@MessageType varchar(50)
	)
AS
BEGIN
	SET NOCOUNT ON;
	
	DECLARE @ContextInfo varbinary(128)
 
	SELECT @ContextInfo = CAST( ''Fusion:''+@MessageType AS VARBINARY(128) );
 
	SET CONTEXT_INFO @ContextInfo
END'


GO


/* ----------------------------------------------------------------
-- Testing code
------------------------------------------------------------------*/
-- ALTER QUEUE fusion.qFusion WITH STATUS = ON;
--EXEC [fusion].[spClearFusionMessageQueue] 'staffChange'
--select * from fusion.qfusion