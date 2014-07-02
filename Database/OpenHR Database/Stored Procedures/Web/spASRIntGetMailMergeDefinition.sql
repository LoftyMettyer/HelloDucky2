CREATE PROCEDURE [dbo].[spASRIntGetMailMergeDefinition] (	
			@piReportID 			integer, 	
			@psCurrentUser			varchar(255),		
			@psAction				varchar(255),
			@psErrorMsg				varchar(MAX)	OUTPUT,		
			@psReportName			varchar(255)	OUTPUT,		
			@psReportOwner			varchar(255)	OUTPUT,		
			@psReportDesc			varchar(255)	OUTPUT,		
			@piBaseTableID			integer			OUTPUT,		
			@piSelection			integer			OUTPUT,		
			@piPicklistID			integer			OUTPUT,		
			@psPicklistName			varchar(255)	OUTPUT,		
			@pfPicklistHidden		bit				OUTPUT,		
			@piFilterID				integer			OUTPUT,		
			@psFilterName			varchar(255)	OUTPUT,		
			@pfFilterHidden			bit				OUTPUT,		
			@piOutputFormat				integer			OUTPUT,		
			@pfOutputSave				bit				OUTPUT,		
			@psOutputFileName			varchar(MAX)	OUTPUT,		
			@piEmailAddrID 			integer			OUTPUT,		
			@psEmailSubject			varchar(255)	OUTPUT,		
			@psTemplateFileName		varchar(MAX)	OUTPUT,		
			@pfOutputScreen				bit				OUTPUT,		
			@pfEmailAsAttachment	bit				OUTPUT,		
			@psEmailAttachmentName	varchar(MAX)	OUTPUT,		
			@pfSuppressBlanks		bit				OUTPUT,		
			@pfPauseBeforeMerge		bit				OUTPUT,		
			@pfOutputPrinter			bit				OUTPUT,		
			@psOutputPrinterName	varchar(255)	OUTPUT,		
			@piDocumentMapID			integer		OUTPUT,		
			@pfManualDocManHeader		bit		OUTPUT,		
		 	@piTimestamp			integer			OUTPUT,		
			@psWarningMsg			varchar(MAX)	OUTPUT		
		)		
		AS		
		BEGIN		
			SET NOCOUNT ON;		
			DECLARE	@iCount		integer,		
					@sTempHidden	varchar(MAX),		
					@sAccess 		varchar(MAX),		
					@fSysSecMgr		bit;		
			SET @psErrorMsg = '';		
			SET @psPicklistName = '';		
			SET @pfPicklistHidden = 0;		
			SET @psFilterName = '';		
			SET @pfFilterHidden = 0;		
			SET @psWarningMsg = '';		
			exec [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;		
			/* Check the mail merge exists. */		
			SELECT @iCount = COUNT(*)		
			FROM [dbo].[ASRSysMailMergeName]		
			WHERE MailMergeID = @piReportID;		
			IF @iCount = 0		
			BEGIN		
				SET @psErrorMsg = 'mail merge has been deleted by another user.';		
				RETURN;		
			END		
			SELECT @psReportName = name,		
				@psReportDesc	 = description,		
				@psReportOwner = userName,		
				@piBaseTableID = tableID,		
				@piSelection = selection,		
				@piPicklistID = picklistID,		
				@piFilterID = filterID,		
				@piOutputFormat = outputformat,		
				@pfOutputSave = outputsave,		
				@psOutputFileName = outputfilename,		
				@piEmailAddrID = emailAddrID,		
				@psEmailSubject = emailSubject,		
				@psTemplateFileName = templateFileName,		
				@pfOutputScreen = outputscreen,		
				@pfEmailAsAttachment = emailasattachment,		
				@psEmailAttachmentName = isnull(emailattachmentname,''),		
				@pfSuppressBlanks = suppressblanks,		
				@pfPauseBeforeMerge = pausebeforemerge,		
				@pfOutputPrinter = outputprinter,		
				@psOutputPrinterName = outputprintername,		
				@piDocumentMapID = documentmapid,		
				@pfManualDocManHeader = manualdocmanheader,				
				@piTimestamp = convert(integer, timestamp)		
			FROM [dbo].[ASRSysMailMergeName]		
			WHERE MailMergeID = @piReportID;		
			/* Check the current user can view the report. */		
			exec [dbo].[spASRIntCurrentUserAccess]		
				9, 		
				@piReportID,		
				@sAccess OUTPUT;		
			IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 		
			BEGIN		
				SET @psErrorMsg = 'mail merge has been made hidden by another user.';		
				RETURN;		
			END		
			IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 		
			BEGIN		
				SET @psErrorMsg = 'mail merge has been made read only by another user.';		
				RETURN;		
			END		
			/* Check the report has details. */		
			SELECT @iCount = COUNT(*)		
			FROM [dbo].[ASRSysMailMergeColumns]		
			WHERE MailMergeID = @piReportID;		
			IF @iCount = 0		
			BEGIN		
				SET @psErrorMsg = 'mail merge contains no details.';		
				RETURN;		
			END		
			/* Check the report has sort order details. */		
			SELECT @iCount = COUNT(*)		
			FROM [dbo].[ASRSysMailMergeColumns]		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.sortOrderSequence > 0;		
			IF @iCount = 0		
			BEGIN		
				SET @psErrorMsg = 'mail merge contains no sort order details.';		
				RETURN;		
			END		
			IF @psAction = 'copy' 		
			BEGIN		
				SET @psReportName = left('copy of ' + @psReportName, 50);		
				SET @psReportOwner = @psCurrentUser;		
			END		
			IF @piPicklistID > 0 		
			BEGIN		
				SELECT @psPicklistName = name,		
					@sTempHidden = access		
				FROM [dbo].[ASRSysPicklistName]		
				WHERE picklistID = @piPicklistID;		
				IF UPPER(@sTempHidden) = 'HD'		
				BEGIN		
					SET @pfPicklistHidden = 1;		
				END		
			END		
			IF @piFilterID > 0 		
			BEGIN		
				SELECT @psFilterName = name,		
					@sTempHidden = access		
				FROM [dbo].[ASRSysExpressions]		
				WHERE exprID = @piFilterID;		
				IF UPPER(@sTempHidden) = 'HD'		
				BEGIN		
					SET @pfFilterHidden = 1;		
				END		
			END

			-- Columns
			SELECT ASRSysMailMergeColumns.[type],
				ASRSysColumns.tableID,
				ASRSysMailMergeColumns.columnID,
				ASRSysColumns.columnName AS [name], 
				ASRSysTables.tableName + '.' + ASRSysColumns.columnName AS [heading],
				ASRSysColumns.DataType,
				ASRSysMailMergeColumns.size,
				ASRSysMailMergeColumns.decimals,
				CASE WHEN ASRSysColumns.DataType = 2 or ASRSysColumns.DataType = 4 THEN '1' ELSE '0' END AS [isnumeric],		
				ASRSysMailMergeColumns.SortOrderSequence AS [sequence]
			FROM ASRSysMailMergeColumns		
			INNER JOIN ASRSysColumns ON ASRSysMailMergeColumns.columnID = ASRSysColumns.columnId		
			INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.type = 'C'		

			-- Expressions
			SELECT CASE WHEN ASRSysExpressions.access = 'HD' THEN 1 ELSE 0 END AS [ishidden],		
				ASRSysMailMergeColumns.[type],
				ASRSysExpressions.tableID,
				ASRSysMailMergeColumns.columnID,
				ASRSysExpressions.name AS [name],
				convert(varchar(MAX), '<Calc> ' + replace(ASRSysExpressions.name, '_', ' ')) AS [heading],
				ASRSysMailMergeColumns.size,
				ASRSysMailMergeColumns.decimals,
				ASRSysMailMergeColumns.SortOrderSequence AS [sequence]
			FROM ASRSysMailMergeColumns		
			INNER JOIN ASRSysExpressions ON ASRSysMailMergeColumns.columnID = ASRSysExpressions.exprID		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.type <> 'C'		
				AND ((ASRSysExpressions.username = @psReportOwner)	OR (ASRSysExpressions.access <> 'HD'))		

			-- Orders
			SELECT ASRSysMailMergeColumns.columnID,
				convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) AS [columnname],
				ASRSysMailMergeColumns.sortOrder,
				ASRSysTables.tableID,
				ASRSysMailMergeColumns.sortOrderSequence AS [sequence]
			FROM ASRSysMailMergeColumns		
			INNER JOIN ASRSysColumns ON ASRSysMailMergeColumns.columnid = ASRSysColumns.columnId		
			INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.sortOrderSequence > 0		
			ORDER BY ASRSysMailMergeColumns.type, [sequence] ASC;

			IF @fSysSecMgr = 0 		
			BEGIN		
				SELECT @iCount = COUNT(ASRSysMailMergeColumns.ID)		
				FROM [dbo].[ASRSysMailMergeColumns]		
				INNER JOIN ASRSysExpressions ON ASRSysMailMergeColumns.columnID = ASRSysExpressions.exprID		
				WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
					AND ASRSysMailMergeColumns.type <> 'C'		
					and ((ASRSysExpressions.username <> @psReportOwner) and (ASRSysExpressions.access = 'HD'));		
							
				IF @iCount > 0 		
				BEGIN		
					IF @iCount = 1		
					BEGIN		
						SET @psWarningMsg = 'A calculation used in this definition has been made hidden by another user. It will be removed from the definition';		
					END		
					ELSE		
					BEGIN		
						SET @psWarningMsg = 'Some calculations used in this definition have been made hidden by another user. They will be removed from the definition';		
					END		
				END		
			END		
		END
