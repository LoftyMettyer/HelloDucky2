	-- Apply the stored procedures
	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessage]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessage];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessageCheckContext]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessageCheckContext];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastProcessedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastGeneratedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingGetLastMessageDates]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingGetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageLogCheck]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageLogCheck];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageLogAdd]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageLogAdd];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateSetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateSetBusRef];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateGetLocalId]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateGetLocalId];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateGetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateGetBusRef];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pConvertData]') AND xtype = 'P')
		DROP FUNCTION [fusion].[pConvertData]


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pIdTranslateGetBusRef
	--
	-- Purpose: Converts a local identifier into a uniqueidentifier for the bus, 
	--			returning consistent value for all future conversions.  
	--          This will create a new identifier where one is not found where
	--			@CanGenerate = 1
	--
	-- Returns: 0 = success, 1 = failure
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pIdTranslateGetBusRef]
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
	END';

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pIdTranslateGetLocalId
	--
	-- Purpose: Finds the local id equivelant for the given Bus reference number, 
	--          assuming it has previous been created through spIdTranslateSetBusRef
	--
	-- Returns: 
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pIdTranslateGetLocalId]
		(
			@TranslationName varchar(50),
			@BusRef uniqueidentifier,
			@LocalId varchar(25) output
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SET @LocalId = null;
	
		SELECT @LocalId = LocalId from [fusion].IdTranslation WITH (ROWLOCK) 
			WHERE TranslationName = @TranslationName and BusRef = @BusRef;
	END';

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spIdTranslateSetBusRef
	--
	-- Purpose: Sets the conversion of a given local reference into the given bus ref
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pIdTranslateSetBusRef]
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
	END	'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageLogAdd
	--
	-- Purpose: Adds fact that message has been processed to local message log
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageLogAdd]
		(
			@MessageType varchar(50),
			@MessageRef uniqueidentifier,
			@Originator varchar(50) = NULL
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		INSERT fusion.MessageLog (MessageType, MessageRef, Originator, ReceivedDate) VALUES (@MessageType, @MessageRef, @Originator, GETUTCDATE());

	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageLogCheck
	--
	-- Purpose: Checks whether message has been processed before
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageLogCheck]
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

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingGetLastGeneratedXml
	--
	-- Purpose: Gets the last generated XML for a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml]
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

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingGetLastMessageDates
	--
	-- Purpose: Gets the last processing date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates]
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

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingSetLastGeneratedDate
	--
	-- Purpose: Sets the last processed date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastGeneratedDate datetime
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM [fusion].MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE [fusion].MessageTracking
			   SET LastGeneratedDate = @LastGeneratedDate
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT [fusion].MessageTracking (MessageType, BusRef, LastGeneratedDate)
				VALUES (@MessageType, @BusRef, @LastGeneratedDate)
		END		
	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingSetLastGeneratedXml
	--
	-- Purpose: Sets the last generated XML for a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml]
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

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingSetLastProcessedDate
	--
	-- Purpose: Sets the last processed date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate]
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

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pSendMessage
	--
	-- Purpose: Triggers a message to be sent
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pSendMessage]
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

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pSendMessageCheckContext
	--
	-- Purpose: Triggers a message to be sent, checking context
	--          to see if we are in the process of updating according to
	--          this same message being received (preventing multi-master
	--          re-publish scenario)
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pSendMessageCheckContext]
		(
			@MessageType varchar(50),
			@LocalId int
		)
	AS
	BEGIN
		SET NOCOUNT ON;
	
		DECLARE @ContextInfo varbinary(128);
 
		SELECT @ContextInfo = CAST( ''Fusion:''+@MessageType AS VARBINARY(128) );
 
		IF CONTEXT_INFO() IS NULL OR CONTEXT_INFO() <> @ContextInfo
		BEGIN	
			EXEC fusion.pSendMessage @MessageType, @LocalId;
		END
	END'
