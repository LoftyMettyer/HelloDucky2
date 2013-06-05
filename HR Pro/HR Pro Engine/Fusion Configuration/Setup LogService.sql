/*
Run this script on:

        (local).blank    -  This database will be modified

to synchronize it with:

        (local).FusionLog

You are recommended to back up your database before running this script

Script created by SQL Compare version 10.1.0 from Red Gate Software Ltd at 28/03/2012 14:11:45

*/
SET NUMERIC_ROUNDABORT OFF
GO
SET ANSI_PADDING, ANSI_WARNINGS, CONCAT_NULL_YIELDS_NULL, ARITHABORT, QUOTED_IDENTIFIER, ANSI_NULLS ON
GO
IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE id=OBJECT_ID('tempdb..#tmpErrors')) DROP TABLE #tmpErrors
GO
CREATE TABLE #tmpErrors (Error int)
GO
SET XACT_ABORT ON
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO
BEGIN TRANSACTION
GO
PRINT N'Creating [dbo].[IdTranslation]'
GO
CREATE TABLE [dbo].[IdTranslation]
(
[ConnectorName] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL,
[TranslationName] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL,
[LocalId] [varchar] (25) COLLATE Latin1_General_CI_AS NOT NULL,
[BusRef] [uniqueidentifier] NOT NULL,
[Time] [datetime] NOT NULL,
[MessageId] [uniqueidentifier] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[pIdTranslationLogAdd]'
GO
---------------------------------------------------------------------------------
-- Name:    pIdTranslationLogAdd
--
-- Purpose: Adds fact that an id has been translated to message log
--
-- Returns: n/a
---------------------------------------------------------------------------------

CREATE PROCEDURE [dbo].[pIdTranslationLogAdd]
	(
		@ConnectorName varchar(50),
		@TranslationName varchar(50),
		@BusRef uniqueidentifier,
		@LocalId varchar(25),
		@Time datetime,
		@MessageId uniqueidentifier		
	)

AS
BEGIN
	SET NOCOUNT ON;
		
	-- If the translation varies for any reason, we want to see them all
	
	INSERT IdTranslation (ConnectorName, BusRef, TranslationName, Time, LocalId, MessageId)
			   VALUES (@ConnectorName, @BusRef, @TranslationName, @Time, @LocalId, @MessageId)
				   
END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[FusionLog]'
GO
CREATE TABLE [dbo].[FusionLog]
(
[Id] [uniqueidentifier] NOT NULL,
[MessageId] [uniqueidentifier] NULL,
[ConnectorName] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL,
[EntityRef] [uniqueidentifier] NULL,
[Time] [datetime] NOT NULL,
[LogLevel] [char] (1) COLLATE Latin1_General_CI_AS NOT NULL,
[Message] [varchar] (512) COLLATE Latin1_General_CI_AS NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[pFusionLogAdd]'
GO
---------------------------------------------------------------------------------
-- Name:    pFusionLogAdd
--
-- Purpose: Adds an entry to the fusion log
--
-- Returns: n/a
---------------------------------------------------------------------------------

CREATE PROCEDURE [dbo].[pFusionLogAdd]
	(
	@Id uniqueidentifier,
	@MessageId [uniqueidentifier],
	@ConnectorName [varchar](50),
	@EntityRef [uniqueidentifier],
	@Time [datetime],

	@LogLevel char(1),
	@Message varchar(512) 
	)

AS
BEGIN
	SET NOCOUNT ON;
	
	INSERT FusionLog (Id, MessageId, ConnectorName, EntityRef, Time, LogLevel, Message)
			   VALUES (@Id, @MessageId, @ConnectorName, @EntityRef, @Time, @LogLevel, @Message)
				   
END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
IF EXISTS (SELECT * FROM #tmpErrors) ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT>0 BEGIN
PRINT 'The database update succeeded'
COMMIT TRANSACTION
END
ELSE PRINT 'The database update failed'
GO
DROP TABLE #tmpErrors
GO
