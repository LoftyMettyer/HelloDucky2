/* --------------------------------------------------- */
/* Update the database from version 5.0 to version 5.1 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion int,
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @ownerGUID uniqueidentifier
DECLARE @newDesktopImageID	integer,
		@picname			varchar(255),
		@oldDesktopImageID	integer;

DECLARE @sSPCode nvarchar(max)

DECLARE @admingroups TABLE(groupname nvarchar(255))


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON
SET @DBName = DB_NAME()

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '5.0') and (@sDBVersion <> '5.1')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion < 9)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
/* Step - Data Cleansing */
/* ------------------------------------------------------------- */

	EXECUTE sp_executeSQL N'UPDATE ASRSysColumns SET lostFocusExprID = 0 WHERE (lostFocusExprID = - 1);';	
	EXECUTE sp_executeSQL N'UPDATE ASRSysColumns SET dfltValueExprID = 0 WHERE (dfltValueExprID = - 1);';

/* ------------------------------------------------------------- */
/* Step - Menu Enhancements */
/* ------------------------------------------------------------- */



	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = object_ID(N'tbsys_userusage') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [tbsys_userusage](
			[objecttype]	smallint, 
			[objectid]	integer,
			[username]	varchar(255),
			[lastrun]	datetime,
			[runcount]	integer)';
		GRANT INSERT, UPDATE, SELECT, DELETE ON dbo.[tbsys_userusage] TO [ASRSysGroup];
	END

	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = object_ID(N'tbsys_userfavourites') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [tbsys_userfavourites](
			[username]		varchar(255),
			[objecttype]	smallint, 
			[objectid]		integer,
			[dateset]		datetime)';
		GRANT INSERT, UPDATE, SELECT, DELETE ON dbo.[tbsys_userfavourites] TO [ASRSysGroup];
	END

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_updateobjectusage]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_updateobjectusage];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_updateobjectusage](@objecttype integer, @objectid integer)
	AS	
	BEGIN

		DECLARE @sUsername varchar(255)
		
		SET @sUsername = SYSTEM_USER;

		IF NOT EXISTS(SELECT [objectid] FROM dbo.[tbsys_userusage] WHERE [objecttype] = @objecttype AND [objectid] = @objectID AND [username] = @sUsername)
		BEGIN
			INSERT tbsys_userusage (objecttype, objectid, username, lastrun, runcount)
				VALUES (@objecttype, @objectID, @sUsername , GETDATE(), 1)
		END
		ELSE
		BEGIN
			UPDATE dbo.[tbsys_userusage] SET [lastrun] = GETDATE(), [runcount] = [runcount] + 1
				WHERE [objecttype] = @objecttype AND [objectid] = @objectID AND [username] = @sUsername
		END

	END';
	GRANT EXECUTE ON dbo.[spstat_updateobjectusage] TO [ASRSysGroup];

	IF EXISTS (SELECT * FROM sys.views WHERE object_id = object_ID(N'[dbo].[ASRSysAllobjectNames]'))
		DROP VIEW [dbo].[ASRSysAllobjectNames]
	EXEC sp_executesql N'CREATE VIEW ASRSysAllobjectNames
	AS
		SELECT 2 AS [objectType], ID, Name FROM ASRSysCustomReportsName
		UNION
		SELECT 1 AS [objectType], CrossTabID AS ID, Name FROM ASRSysCrossTab
		UNION		
		SELECT 14 AS [objectType], MatchReportID AS ID, Name FROM ASRSysMatchReportName
		UNION
		SELECT 15 AS [objectType], 0 AS ID, ''Absence Breakdown''
		UNION
		SELECT 16 AS [objectType], 0 AS ID, ''Bradford Factor''
		UNION		
		SELECT 17 AS [objectType], ID AS ID, Name FROM ASRSysCalendarReports
		UNION		
		SELECT 20 AS [objectType], RecordProfileID AS ID, Name FROM ASRSysRecordProfileName
		UNION
		SELECT 23 AS [objectType], 0 AS ID, ''Succession Planning''
		UNION
		SELECT 24 AS [objectType], 0 AS ID, ''Career Progression''
		UNION
		SELECT 30 AS [objectType], 0 AS ID, ''Turnover''
		UNION
		SELECT 31 AS [objectType], 0 AS ID, ''Stability Index''';
	GRANT SELECT ON dbo.[ASRSysAllobjectNames] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_recentlyrunobjects]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_recentlyrunobjects];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_recentlyrunobjects]
	AS
	BEGIN

		SELECT TOP 10 ROW_NUMBER() OVER (ORDER BY [lastrun] DESC) AS ID, u.[objectid], o.[Name], o.[objectType]
			FROM tbsys_userusage u
			INNER JOIN ASRSysAllobjectNames o ON o.[objectType] = u.objecttype AND o.[ID] = u.objectid
			WHERE [username] = SYSTEM_USER
			ORDER BY u.[lastrun] DESC

	END';
	GRANT EXECUTE ON dbo.[spstat_recentlyrunobjects] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getfavourites]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getfavourites];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_getfavourites]
	AS
	BEGIN

		SELECT TOP 10 o.[objectType], f.[objectid], o.[Name]
			FROM tbsys_userfavourites f
			INNER JOIN ASRSysAllobjectNames o ON o.[objectType] = f.[objecttype] AND o.[ID] = f.objectid
			WHERE [username] = SYSTEM_USER

	END';
	GRANT EXECUTE ON dbo.[spstat_getfavourites] TO [ASRSysGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_addtofavourites]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_addtofavourites];
	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spstat_addtofavourites](@objecttype integer, @objectid integer, @count tinyint OUTPUT)
		AS
		BEGIN

			DECLARE @now datetime;
			SET @now = GETDATE();
			
			IF NOT EXISTS(SELECT [username] FROM dbo.tbsys_userfavourites 
								WHERE [username] = SYSTEM_USER AND @objectid = [objectid] AND @objecttype = [objecttype])
			BEGIN
				INSERT dbo.tbsys_userfavourites (username, objecttype, objectid, dateset)
					VALUES (SYSTEM_USER, @objecttype, @objectid, @now);
			END

			SELECT @count = COUNT(*) FROM dbo.tbsys_userfavourites WHERE [username] = SYSTEM_USER;

		END';
	GRANT EXECUTE ON dbo.[spstat_addtofavourites] TO [ASRSysGroup];


/* ------------------------------------------------------------- */
/* Step - Management Packs */
/* ------------------------------------------------------------- */

	IF NOT EXISTS(SELECT * FROM ASRSysFileFormats where ID = 923)
	BEGIN
		INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default])
			VALUES (923, 'Word', 'PDF (*.pdf)', 'pdf', NULL, 17, 0);
		INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default])
			VALUES (924, 'Word', 'Rich Text Format (*.rtf)', 'rtf', NULL, 6, 0);
		INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default])
			VALUES (925, 'Word', 'Plain Text (*.txt)', 'txt', NULL, 2, 0);
		INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default])
			VALUES (926, 'Word', 'Web Page (*.html)', 'html', NULL, 8, 0);		
		INSERT ASRSysFileFormats (ID, Destination, [Description], Extension, Office2003, Office2007, [Default])
			VALUES (927, 'Excel', 'Web Page (*.html)', 'html', NULL, 44, 0);
	END


/* ------------------------------------------------------------- */
/* Step - Updating workflow stored procedures */
/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
/* Step - Updating User SEttings with data/columns for Omit spacer DEV */
/* ------------------------------------------------------------- 
	IF NOT EXISTS(SELECT * FROM ASRSysUserSettings where section = 'Output')
	BEGIN
		INSERT ASRSysUserSettings ([UserName],[Section],[SettingKey],[SettingValue])
			VALUES ('HRPro','Output','ExcelOmitSpacerRow','0');
		INSERT ASRSysUserSettings ([UserName],[Section],[SettingKey],[SettingValue])
			VALUES ('Admin','Output','ExcelOmitSpacerCol','0');	
		INSERT ASRSysUserSettings ([UserName],[Section],[SettingKey],[SettingValue])
			VALUES ('HRPro','Output','ExcelOmitSpacerRow','0');
		INSERT ASRSysUserSettings ([UserName],[Section],[SettingKey],[SettingValue])
			VALUES ('Admin','Output','ExcelOmitSpacerCol','0');				
	END		
	----------------------------------------------------------------------
	-- spASRSubmitWorkflowStep
	----------------------------------------------------------------------*/

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSubmitWorkflowStep]
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psFormInput1		varchar(MAX),
				@psFormElements		varchar(MAX)	OUTPUT,
				@pfSavedForLater	bit				OUTPUT,
				@piPageNo	integer
			)
			AS
			BEGIN
				DECLARE
					@iIndex1			integer,
					@iIndex2			integer,
					@iID				integer,
					@sID				varchar(MAX),
					@sValue				varchar(MAX),
					@iElementType		integer,
					@iPreviousElementID	integer,
					@iValue				integer,
					@hResult			integer,
					@hTmpResult			integer,
					@sTo				varchar(MAX),
					@sCopyTo			varchar(MAX),
					@sTempTo			varchar(MAX),
					@sMessage			varchar(MAX),
					@sMessage_HypertextLinks	varchar(MAX),
					@sHypertextLinkedSteps		varchar(MAX),
					@iEmailID			integer,
					@iEmailCopyID		integer,
					@iTempEmailID		integer,
					@iEmailLoop			integer,
					@iEmailRecord		integer,
					@iEmailRecordID		integer,
					@sSQL				nvarchar(MAX),
					@iCount				integer,
					@superCursor		cursor,
					@curDelegatedRecords	cursor,
					@fDelegate			bit,
					@fDelegationValid	bit,
					@iDelegateEmailID	integer,
					@iDelegateRecordID	integer,
					@sTemp				varchar(MAX),
					@sDelegateTo		varchar(MAX),
					@sAllDelegateTo		varchar(MAX),
					@iCurrentStepID		int,
					@sDelegatedMessage	varchar(MAX),
					@iTemp				integer, 
					@iPrevElementType	integer,
					@iWorkflowID		integer,
					@sRecSelIdentifier	varchar(MAX),
					@sRecSelWebFormIdentifier	varchar(MAX), 
					@iStepID			int,
					@iElementID			int,
					@sUserName			varchar(MAX),
					@sUserEmail			varchar(MAX), 
					@sValueDescription	varchar(MAX),
					@iTableID			integer,
					@iRecDescID			integer,
					@sEvalRecDesc		varchar(MAX),
					@sExecString		nvarchar(MAX),
					@sParamDefinition	nvarchar(500),
					@sIdentifier		varchar(MAX),
					@iItemType			integer,
					@iDataAction		integer, 
					@fValidRecordID		bit,
					@iEmailTableID		integer,
					@iEmailType			integer,
					@iBaseTableID		integer,
					@iBaseRecordID		integer,
					@iRequiredRecordID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iTempElementID		integer,
					@iTrueFlowType		integer,
					@iExprID			integer,
					@iResultType		integer,
					@sResult			varchar(MAX),
					@fResult			bit,
					@dtResult			datetime,
					@fltResult			float,
					@sEmailSubject		varchar(200),
					@iTempID			integer,
					@iBehaviour			integer;
		
				SET @pfSavedForLater = 0;
		
				SELECT @iCurrentStepID = ID
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
				SET @iDelegateEmailID = 0;
				SELECT @sTemp = ISNULL(parameterValue, '''')
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_DelegateEmail'';
				SET @iDelegateEmailID = convert(integer, @sTemp);
		
				SET @psFormElements = '''';
						
				-- Get the type of the given element 
				SELECT @iElementType = E.type,
					@iEmailID = E.emailID,
					@iEmailCopyID = isnull(E.emailCCID, 0),
					@iEmailRecord = E.emailRecord, 
					@iWorkflowID = E.workflowID,
					@sRecSelIdentifier = E.RecSelIdentifier, 
					@sRecSelWebFormIdentifier = E.RecSelWebFormIdentifier, 
					@iTableID = E.dataTableID,
					@iDataAction = E.dataAction, 
					@iTrueFlowType = isnull(E.trueFlowType, 0), 
					@iExprID = isnull(E.trueFlowExprID, 0), 
					@sEmailSubject = ISNULL(E.emailSubject, '''')
				FROM ASRSysWorkflowElements E
				WHERE E.ID = @piElementID;
		
				--------------------------------------------------
				-- Read the submitted webForm/storedData values
				--------------------------------------------------
				IF @iElementType = 5 -- Stored Data element
				BEGIN
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
					SET @sValue = LEFT(@psFormInput1, @iIndex1-1);
					SET @sTemp = SUBSTRING(@psFormInput1, @iIndex1+1, LEN(@psFormInput1) - @iIndex1);
		
					SET @sValueDescription = '''';
					SET @sMessage = ''Successfully '' +
						CASE
							WHEN @iDataAction = 0 THEN ''inserted''
							WHEN @iDataAction = 1 THEN ''updated''
							ELSE ''deleted''
						END + '' record'';
		
					IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
					BEGIN
						SET @sValueDescription = @sTemp;
					END
					ELSE
					BEGIN
						SET @iTemp = convert(integer, @sValue);
						IF @iTemp > 0 
						BEGIN	
							EXEC [dbo].[spASRRecordDescription] 
								@iTableID,
								@iTemp,
								@sEvalRecDesc OUTPUT
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
						END
					END
		
					IF len(@sValueDescription) > 0 SET @sMessage = @sMessage + '' ('' + @sValueDescription + '')'';
		
					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue, 
						ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
						AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
				END
				ELSE
				BEGIN
					-- Put the submitted form values into the ASRSysWorkflowInstanceValues table. 
					WHILE (charindex(CHAR(9), @psFormInput1) > 0)
					BEGIN
		
						SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
						SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1);
						SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''''''', '''''''''''');
						SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1);
						SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2);
		
						--Get the record description (for RecordSelectors only)
						SET @sValueDescription = '''';
		
						-- Get the WebForm item type, etc.
						SELECT @sIdentifier = EI.identifier,
							@iItemType = EI.itemType,
							@iTableID = EI.tableID,
							@iBehaviour = EI.behaviour
						FROM ASRSysWorkflowElementItems EI
						WHERE EI.ID = convert(integer, @sID);
		
						SET @iParent1TableID = 0;
						SET @iParent1RecordID = 0;
						SET @iParent2TableID = 0;
						SET @iParent2RecordID = 0;
		
						IF @iItemType = 11 -- Record Selector
						BEGIN
							-- Get the table record description ID. 
							SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
							FROM ASRSysTables 
							WHERE ASRSysTables.tableID = @iTableID;
		
							SET @iTemp = convert(integer, isnull(@sValue, ''0''));
		
							-- Get the record description. 
							IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iTemp > 0)
							BEGIN
								SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(MAX), @iRecDescID) + '' @recDesc OUTPUT, @recID'';
								SET @sParamDefinition = N''@recDesc varchar(MAX) OUTPUT, @recID integer'';
								EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp;
								IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
							END
		
							-- Record the selected record''s parent details.
							exec [dbo].[spASRGetParentDetails]
								@iTableID,
								@iTemp,
								@iParent1TableID	OUTPUT,
								@iParent1RecordID	OUTPUT,
								@iParent2TableID	OUTPUT,
								@iParent2RecordID	OUTPUT;
						END
						ELSE
						IF (@iItemType = 0) and (@iBehaviour = 1) AND (@sValue = ''1'')-- SaveForLater Button
						BEGIN
							SET @pfSavedForLater = 1;
						END
		
						IF (@iItemType = 17) -- FileUpload Control
						BEGIN
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.fileUpload_File = 
								CASE 
									WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_File
									ELSE null
								END,
								ASRSysWorkflowInstanceValues.fileUpload_ContentType = 
								CASE 
									WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_ContentType
									ELSE null
								END,
								ASRSysWorkflowInstanceValues.fileUpload_FileName = 
								CASE 
									WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_FileName
									ELSE null
								END
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @piElementID
								AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
						END
						ELSE
						BEGIN
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = @sValue, 
								ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription,
								ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID,
								ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
								ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID,
								ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @piElementID
								AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
						END
					END
		
					IF @pfSavedForLater = 1
					BEGIN
						/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 7
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
						
						/* Remember the page number too  */
						UPDATE ASRSysWorkflowInstances
						SET ASRSysWorkflowInstances.pageno = @piPageNo
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
						RETURN;
					END
				END
					
				SET @hResult = 0;
				SET @sTo = '''';
				SET @sCopyTo = '''';
		
				--------------------------------------------------
				-- Process email element
				--------------------------------------------------
				IF @iElementType = 3 -- Email element
				BEGIN
					-- Get the email recipient. 
					SET @iEmailRecordID = 0;
					SET @sSQL = ''spASRSysEmailAddr'';
		
					IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						SET @iEmailLoop = 0
						WHILE @iEmailLoop < 2
						BEGIN
							SET @hTmpResult = 0;
							SET @sTempTo = '''';
							SET @iTempEmailID = 
								CASE 
									WHEN @iEmailLoop = 1 THEN @iEmailCopyID
									ELSE isnull(@iEmailID, 0)
								END;
		
							IF @iTempEmailID > 0 
							BEGIN
								SET @fValidRecordID = 1;
		
								SELECT @iEmailTableID = isnull(tableID, 0),
									@iEmailType = isnull(type, 0)
								FROM ASRSysEmailAddress
								WHERE emailID = @iTempEmailID;
		
								IF @iEmailType = 0 
								BEGIN
									SET @iEmailRecordID = 0;
								END
								ELSE
								BEGIN
									SET @iTempElementID = 0;
		
									-- Get the record ID required. 
									IF (@iEmailRecord = 0) OR (@iEmailRecord = 4)
									BEGIN
										/* Initiator record. */
										SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID,
											@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
											@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
											@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
											@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
										FROM ASRSysWorkflowInstances
										WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
										SET @iBaseRecordID = @iEmailRecordID;
		
										IF @iEmailRecord = 4
										BEGIN
											-- Trigger record
											SELECT @iBaseTableID = isnull(WF.baseTable, 0)
											FROM ASRSysWorkflows WF
											INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
												AND WFI.ID = @piInstanceID;
										END
										ELSE
										BEGIN
											-- Initiator''s record
											SELECT @iBaseTableID = convert(integer, ISNULL(parameterValue, ''0''))
											FROM ASRSysModuleSetup
											WHERE moduleKey = ''MODULE_PERSONNEL''
												AND parameterKey = ''Param_TablePersonnel'';
		
											IF @iBaseTableID = 0
											BEGIN
												SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
												FROM ASRSysModuleSetup
												WHERE moduleKey = ''MODULE_WORKFLOW''
												AND parameterKey = ''Param_TablePersonnel'';
											END
										END
									END
		
									IF @iEmailRecord = 1
									BEGIN
										SELECT @iPrevElementType = ASRSysWorkflowElements.type,
											@iTempElementID = ASRSysWorkflowElements.ID
										FROM ASRSysWorkflowElements
										WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
											AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
		
										IF @iPrevElementType = 2
										BEGIN
											 -- WebForm
											SELECT @sValue = ISNULL(IV.value, ''0''),
												@iBaseTableID = EI.tableID,
												@iParent1TableID = IV.parent1TableID,
												@iParent1RecordID = IV.parent1RecordID,
												@iParent2TableID = IV.parent2TableID,
												@iParent2RecordID = IV.parent2RecordID
											FROM ASRSysWorkflowInstanceValues IV
											INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
											INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
											WHERE IV.instanceID = @piInstanceID
												AND IV.identifier = @sRecSelIdentifier
												AND Es.identifier = @sRecSelWebFormIdentifier
												AND Es.workflowID = @iWorkflowID
												AND IV.elementID = Es.ID;
										END
										ELSE
										BEGIN
											-- StoredData
											SELECT @sValue = ISNULL(IV.value, ''0''),
												@iBaseTableID = isnull(Es.dataTableID, 0),
												@iParent1TableID = IV.parent1TableID,
												@iParent1RecordID = IV.parent1RecordID,
												@iParent2TableID = IV.parent2TableID,
												@iParent2RecordID = IV.parent2RecordID
											FROM ASRSysWorkflowInstanceValues IV
											INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
												AND IV.identifier = Es.identifier
												AND Es.workflowID = @iWorkflowID
												AND Es.identifier = @sRecSelWebFormIdentifier
											WHERE IV.instanceID = @piInstanceID;
										END
		
										SET @iEmailRecordID = 
											CASE
												WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
												ELSE 0
											END;
		
										SET @iBaseRecordID = @iEmailRecordID;
									END
		
									SET @fValidRecordID = 1;
									IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
									BEGIN
										SET @fValidRecordID = 0;
		
										EXEC [dbo].[spASRWorkflowAscendantRecordID]
											@iBaseTableID,
											@iBaseRecordID,
											@iParent1TableID,
											@iParent1RecordID,
											@iParent2TableID,
											@iParent2RecordID,
											@iEmailTableID,
											@iRequiredRecordID	OUTPUT;
		
										SET @iEmailRecordID = @iRequiredRecordID;
		
										IF @iRequiredRecordID > 0 
										BEGIN
											EXEC [dbo].[spASRWorkflowValidTableRecord]
												@iEmailTableID,
												@iEmailRecordID,
												@fValidRecordID	OUTPUT;
										END
		
										IF @fValidRecordID = 0
										BEGIN
											IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
											BEGIN
												SELECT @sTempTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
												FROM ASRSysWorkflowQueueColumns QC
												INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
												WHERE WFQ.instanceID = @piInstanceID
													AND QC.emailID = @iTempEmailID;
		
												IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
											END
											ELSE
											BEGIN
												IF @iEmailRecord = 1
												BEGIN
													SELECT @sTempTo = rtrim(ltrim(isnull(IV.value , '''')))
													FROM ASRSysWorkflowInstanceValues IV
													WHERE IV.instanceID = @piInstanceID
														AND IV.emailID = @iTempEmailID
														AND IV.elementID = @iTempElementID;
		
													IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
												END
											END
										END
		
										IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
										BEGIN
											-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
											EXEC [dbo].[spASRWorkflowActionFailed] 
												@piInstanceID, 
												@piElementID, 
												''Email record has been deleted or not selected.'';
													
											SET @hTmpResult = -1;
										END
									END
								END
		
								IF @fValidRecordID = 1
								BEGIN
									/* Get the recipient address. */
									IF len(@sTempTo) = 0
									BEGIN
										EXEC @hTmpResult = @sSQL @sTempTo OUTPUT, @iTempEmailID, @iEmailRecordID;
										IF @sTempTo IS null SET @sTempTo = '''';
									END
		
									IF (LEN(rtrim(ltrim(@sTempTo))) = 0) AND (@iEmailLoop = 0)
									BEGIN
										-- Email step failure if no known recipient.
										-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
										EXEC [dbo].[spASRWorkflowActionFailed] 
											@piInstanceID, 
											@piElementID, 
											''No email recipient.'';
												
										SET @hTmpResult = -1;
									END
								END
		
								IF @iEmailLoop = 1 
								BEGIN
									SET @sCopyTo = @sTempTo;
		
									IF (rtrim(ltrim(@sCopyTo)) = ''@'')
										OR (charindex('' @ '', @sCopyTo) > 0)
									BEGIN
										SET @sCopyTo = '''';
									END
								END
								ELSE
								BEGIN
									SET @sTo = @sTempTo;
								END
							END
						
							SET @iEmailLoop = @iEmailLoop + 1;
		
							IF @hTmpResult <> 0 SET @hResult = @hTmpResult;
						END
					END
		
					IF LEN(rtrim(ltrim(@sTo))) > 0
					BEGIN
						IF (rtrim(ltrim(@sTo)) = ''@'')
							OR (charindex('' @ '', @sTo) > 0)
						BEGIN
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.userEmail = @sTo,
								ASRSysWorkflowInstanceSteps.emailCC = @sCopyTo
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Invalid email recipient.'';
						
							SET @hResult = -1;
						END
						ELSE
						BEGIN
							/* Build the email message. */
							EXEC [dbo].[spASRGetWorkflowEmailMessage] 
								@piInstanceID, 
								@piElementID, 
								@sMessage OUTPUT, 
								@sMessage_HypertextLinks OUTPUT, 
								@sHypertextLinkedSteps OUTPUT, 
								@fValidRecordID OUTPUT, 
								@sTo;
		
							IF @fValidRecordID = 1
							BEGIN
								exec [dbo].[spASRDelegateWorkflowEmail] 
									@sTo,
									@sCopyTo,
									@sMessage,
									@sMessage_HypertextLinks,
									@iCurrentStepID,
									@sEmailSubject;
							END
							ELSE
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC [dbo].[spASRWorkflowActionFailed] 
									@piInstanceID, 
									@piElementID, 
									''Email item database value record has been deleted or not selected.'';
										
								SET @hResult = -1;
							END
						END
					END
				END
		
				--------------------------------------------------
				-- Mark the step as complete
				--------------------------------------------------
				IF @hResult = 0
				BEGIN
					/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 3,
						ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.userEmail = CASE
							WHEN @iElementType = 3 THEN @sTo
							ELSE ASRSysWorkflowInstanceSteps.userEmail
						END,
						ASRSysWorkflowInstanceSteps.emailCC = CASE
							WHEN @iElementType = 3 THEN @sCopyTo
							ELSE ASRSysWorkflowInstanceSteps.emailCC
						END,
						ASRSysWorkflowInstanceSteps.hypertextLinkedSteps = CASE
							WHEN @iElementType = 3 THEN @sHypertextLinkedSteps
							ELSE ASRSysWorkflowInstanceSteps.hypertextLinkedSteps
						END,
						ASRSysWorkflowInstanceSteps.message = CASE
							WHEN @iElementType = 3 THEN @sMessage
							WHEN @iElementType = 5 THEN @sMessage
							ELSE ''''
						END,
						ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
			
					IF @iElementType = 4 -- Decision element
					BEGIN
						IF @iTrueFlowType = 1
						BEGIN
							-- Decision Element flow determined by a calculation
							EXEC [dbo].[spASRSysWorkflowCalculation]
								@piInstanceID,
								@iExprID,
								@iResultType OUTPUT,
								@sResult OUTPUT,
								@fResult OUTPUT,
								@dtResult OUTPUT,
								@fltResult OUTPUT, 
								0;
		
							SET @iValue = convert(integer, @fResult);
						END
						ELSE
						BEGIN
							-- Decision Element flow determined by a button in a preceding web form
							SET @iPrevElementType = 4; -- Decision element
							SET @iPreviousElementID = @piElementID;
		
							WHILE (@iPrevElementType = 4)
							BEGIN
								SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
									@iPrevElementType = isnull(WE.type, 0)
								FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPreviousElementID) PE
								INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
								INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
									AND WIS.instanceID = @piInstanceID;
		
								SET @iPreviousElementID = @iTempID;
							END
					
							SELECT @sValue = ISNULL(IV.value, ''0'')
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
							WHERE IV.elementID = @iPreviousElementID
								AND IV.instanceid = @piInstanceID
								AND E.ID = @piElementID;
		
							SET @iValue = 
								CASE
									WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
									ELSE 0
								END;
						END
				
						IF @iValue IS null SET @iValue = 0;
		
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
			
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.completionDateTime = null
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT SUCC.id FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, @iValue) SUCC)
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 6
								OR ASRSysWorkflowInstanceSteps.status = 8
								OR ASRSysWorkflowInstanceSteps.status = 3);
					END
					ELSE
					BEGIN
						IF @iElementType <> 3 -- 3=Email element
						BEGIN
							IF @iElementType = 2 -- WebForm
							BEGIN
								SELECT @sUserName = isnull(WIS.userName, ''''),
									@sUserEmail = isnull(WIS.userEmail, '''')
								FROM ASRSysWorkflowInstanceSteps WIS
								WHERE WIS.instanceID = @piInstanceID
									AND WIS.elementID = @piElementID;
							END;
									
							-- Do not the following bit when the submitted element is an Email element as 
							-- the succeeding elements will already have been actioned.
							DECLARE @succeedingElements TABLE(elementID integer);
		
							EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
								@piInstanceID, 
								@piElementID, 
								@superCursor OUTPUT,
								'''';
		
							FETCH NEXT FROM @superCursor INTO @iTemp;
							WHILE (@@fetch_status = 0)
							BEGIN
								INSERT INTO @succeedingElements (elementID) VALUES (@iTemp);
							
								FETCH NEXT FROM @superCursor INTO @iTemp;
							END
							CLOSE @superCursor;
							DEALLOCATE @superCursor;
		
							-- If the submitted element is a web form, then any succeeding webforms are actioned for the same user.
							IF @iElementType = 2 -- WebForm
							BEGIN
								-- Return a list of the workflow form elements that may need to be displayed to the initiator straight away 
								DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
								SELECT ASRSysWorkflowInstanceSteps.ID,
									ASRSysWorkflowInstanceSteps.elementID
								FROM ASRSysWorkflowInstanceSteps
								INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
								WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
									AND ASRSysWorkflowInstanceSteps.elementID IN 
										(SELECT suc.elementID
										FROM @succeedingElements suc)
									AND ASRSysWorkflowElements.type = 2
									AND (ASRSysWorkflowInstanceSteps.status = 0
										OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3);
		
								OPEN formsCursor;
								FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
								WHILE (@@fetch_status = 0) 
								BEGIN
									SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
									DELETE FROM ASRSysWorkflowStepDelegation
									WHERE stepID = @iStepID;
		
									INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
										(SELECT WSD.delegateEmail, @iStepID
										FROM ASRSysWorkflowStepDelegation WSD
										WHERE WSD.stepID = @iCurrentStepID);
								
									-- Change the step status to be 2 (pending user input). 
									UPDATE ASRSysWorkflowInstanceSteps
									SET ASRSysWorkflowInstanceSteps.status = 2, 
										ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
										ASRSysWorkflowInstanceSteps.completionDateTime = null,
										ASRSysWorkflowInstanceSteps.userName = @sUserName,
										ASRSysWorkflowInstanceSteps.userEmail = @sUserEmail 
									WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
										AND (ASRSysWorkflowInstanceSteps.status = 0
											OR ASRSysWorkflowInstanceSteps.status = 2
											OR ASRSysWorkflowInstanceSteps.status = 6
											OR ASRSysWorkflowInstanceSteps.status = 8
											OR ASRSysWorkflowInstanceSteps.status = 3);
								
									FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
								END
								CLOSE formsCursor;
								DEALLOCATE formsCursor;
		
								UPDATE ASRSysWorkflowInstanceSteps
								SET ASRSysWorkflowInstanceSteps.status = 1,
									ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
									ASRSysWorkflowInstanceSteps.completionDateTime = null
								WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
									AND ASRSysWorkflowInstanceSteps.elementID IN 
										(SELECT suc.elementID
										FROM @succeedingElements suc)
									AND ASRSysWorkflowInstanceSteps.elementID NOT IN 
										(SELECT ASRSysWorkflowElements.ID
										FROM ASRSysWorkflowElements
										WHERE ASRSysWorkflowElements.type = 2)
									AND (ASRSysWorkflowInstanceSteps.status = 0
										OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3);
							END
							ELSE
							BEGIN
								DELETE FROM ASRSysWorkflowStepDelegation
								WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
									FROM ASRSysWorkflowInstanceSteps
									WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
										AND ASRSysWorkflowInstanceSteps.elementID IN 
											(SELECT suc.elementID
											FROM @succeedingElements suc)
										AND (ASRSysWorkflowInstanceSteps.status = 0
											OR ASRSysWorkflowInstanceSteps.status = 2
											OR ASRSysWorkflowInstanceSteps.status = 6
											OR ASRSysWorkflowInstanceSteps.status = 8
											OR ASRSysWorkflowInstanceSteps.status = 3));
							
								INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
								(SELECT WSD.delegateEmail,
									SuccWIS.ID
								FROM ASRSysWorkflowStepDelegation WSD
								INNER JOIN ASRSysWorkflowInstanceSteps CurrWIS ON WSD.stepID = CurrWIS.ID
								INNER JOIN ASRSysWorkflowInstanceSteps SuccWIS ON CurrWIS.instanceID = SuccWIS.instanceID
									AND SuccWIS.elementID IN (SELECT suc.elementID
										FROM @succeedingElements suc)
									AND (SuccWIS.status = 0
										OR SuccWIS.status = 2
										OR SuccWIS.status = 6
										OR SuccWIS.status = 8
										OR SuccWIS.status = 3)
								INNER JOIN ASRSysWorkflowElements SuccWE ON SuccWIS.elementID = SuccWE.ID
									AND SuccWE.type = 2
								WHERE WSD.stepID = @iCurrentStepID);
		
								UPDATE ASRSysWorkflowInstanceSteps
								SET ASRSysWorkflowInstanceSteps.status = 1,
									ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
									ASRSysWorkflowInstanceSteps.completionDateTime = null,
									ASRSysWorkflowInstanceSteps.userEmail = CASE
										WHEN (SELECT ASRSysWorkflowElements.type 
											FROM ASRSysWorkflowElements 
											WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @sTo -- 2 = Web Form element
										ELSE ASRSysWorkflowInstanceSteps.userEmail
									END
								WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
									AND ASRSysWorkflowInstanceSteps.elementID IN 
										(SELECT suc.elementID
										FROM @succeedingElements suc)
									AND (ASRSysWorkflowInstanceSteps.status = 0
										OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3);
							END
						END
					END
			
					-- Set activated Web Forms to be ''pending'' (to be done by the user) 
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 2
					WHERE ASRSysWorkflowInstanceSteps.id IN (
						SELECT ASRSysWorkflowInstanceSteps.ID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.status = 1
							AND ASRSysWorkflowElements.type = 2);
		
					-- Set activated Terminators to be ''completed'' 
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 3,
						ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
					WHERE ASRSysWorkflowInstanceSteps.id IN (
						SELECT ASRSysWorkflowInstanceSteps.ID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.status = 1
							AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowElements.type = 1);
		
					-- Count how many terminators have completed. ie. if the workflow has completed. 
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.status = 3
						AND ASRSysWorkflowElements.type = 1;
							
					IF @iCount > 0 
					BEGIN
						UPDATE ASRSysWorkflowInstances
						SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
							ASRSysWorkflowInstances.status = 3,
							ASRSysWorkflowInstances.pageno = @piPageNo
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
					
						-- Steps pending action are no longer required.
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
								OR ASRSysWorkflowInstanceSteps.status = 2); -- 2 = Pending User Action
					END
		
					IF @iElementType = 3 -- Email element
						OR @iElementType = 5 -- Stored Data element
					BEGIN
						exec [dbo].[spASREmailImmediate] ''OpenHR Workflow'';
					END
				END
			END';

	EXECUTE sp_executeSQL @sSPCode;



/* ------------------------------------------------------------- */
/* Step - Image updates */
/* ------------------------------------------------------------- */

	-- Create system tracking column
	IF NOT EXISTS(SELECT ID FROM syscolumns	WHERE ID = (SELECT ID FROM sysobjects where [name] = 'ASRSysPictures') AND [name] = 'GUID')
		EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysPictures] ADD [GUID] [uniqueidentifier] NULL;';

	-- Generic image update routine
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_writepicture]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_writepicture];
	EXECUTE sp_executeSQL  N'CREATE PROCEDURE spadmin_writepicture(@guid uniqueidentifier, @name varchar(255), @pictureID integer OUTPUT, @pictureHex varbinary(MAX))
	AS
	BEGIN

		IF NOT EXISTS(SELECT [guid] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid)	
		BEGIN

			SELECT @pictureID = ISNULL(MAX(PictureID), 0) + 1 FROM dbo.[ASRSysPictures];

			INSERT [ASRSysPictures] (PictureID, Name, PictureType, [guid], [Picture]) 
				SELECT @pictureID, @name, 1, @guid, @pictureHex;

		END
		ELSE
		BEGIN
			SELECT @pictureID = [PictureID] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid;
			UPDATE [ASRSysPictures] SET [Name] = @name, Picture = @pictureHex WHERE [guid] = @guid;
		END

	END';

	-- Add/update images
	EXEC dbo.spadmin_writepicture '7410CCC5-01EF-46F0-9D9F-9323A93B4573', 'Default Background.jpg', @newDesktopImageID OUTPUT, 0xFFD8FFE000104A46494600010101004800480000FFDB0043000503040404030504040405050506070C08070707070F0B0B090C110F1212110F111113161C1713141A1511111821181A1D1D1F1F1F13172224221E241C1E1F1EFFDB0043010505050706070E08080E1E1411141E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1E1EFFC0001108017201FE03012200021101031101FFC4001F0000010501010101010100000000000000000102030405060708090A0BFFC400B5100002010303020403050504040000017D01020300041105122131410613516107227114328191A1082342B1C11552D1F02433627282090A161718191A25262728292A3435363738393A434445464748494A535455565758595A636465666768696A737475767778797A838485868788898A92939495969798999AA2A3A4A5A6A7A8A9AAB2B3B4B5B6B7B8B9BAC2C3C4C5C6C7C8C9CAD2D3D4D5D6D7D8D9DAE1E2E3E4E5E6E7E8E9EAF1F2F3F4F5F6F7F8F9FAFFC4001F0100030101010101010101010000000000000102030405060708090A0BFFC400B51100020102040403040705040400010277000102031104052131061241510761711322328108144291A1B1C109233352F0156272D10A162434E125F11718191A262728292A35363738393A434445464748494A535455565758595A636465666768696A737475767778797A82838485868788898A92939495969798999AA2A3A4A5A6A7A8A9AAB2B3B4B5B6B7B8B9BAC2C3C4C5C6C7C8C9CAD2D3D4D5D6D7D8D9DAE2E3E4E5E6E7E8E9EAF2F3F4F5F6F7F8F9FAFFDA000C03010002110311003F00FB2E8A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A2909001248007526800A535C36BFF1134CB7BA6D3B4589F58BF1C1580E6343EEDFE1F98AC873E21D64EFD6F54FB240DFF2E969C0C7A13DFF005AD95093577A13CCBA1DDEA3E20D1EC095B8BE8838EA8A7737E42B3C78A7ED1FF20FD26F2E076765DABF9F358BA75A6956001B7B58F78FE37F99BF335A5FDA23FBD4FD92E8172E0D43C432FDDB0B5807FD349327F434E0DE226EB71609F407FC2A8FF68FFB547F68FF00B553ECD85CBC5FC443A5C5837D41FF000A69BFF1145F7AC6D271FEC4983FA9AA7FDA3FED51FDA3FED51ECD85CB07C506DFFE421A4DE5B8EEEA372FE7C55FD3FC43A3DF90B6F7D16F3FC0E76B7E47AD647F68FF00B559DA8DAE957E0FDA2D632C7F8D7E56FCC557B24F70B9DF0A4AF3051AFE8C77E87AA9B8847FCBA5DFCC31E80F6FD2B5742F88FA6CD74BA76BD03E8F7A780263FBA7FA376FC78F7A4E8492BAD4399753BBC521A6AB0650CA410464107834C96E218466599107FB4C0573CA4A1AC996937B12D19ACB9F5CD321FBD7418FFB00B7F2AA5378B2C107C914F27D140FE66B82AE6B83A5F1545F7DCDA185AD3DA0CE871457272F8C917EE5831FF79F1FD2A03E367079D3811ED37FF5AB827C519641D9D4FC1FF91D0B2CC53FB3F8A3B3EA293F0AE560F1B5931027B69E1F718615B9A6EAB61A8AE6D6E6373DD73861F81E6BB3099D60716ED46A26FB6CFEE66157095A92BCE2D23428A28AF50E70A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A280138C51F851E95C77C51F1EE91E01D01B51D418CF732E52CECD0FEF2E1FD07A01DCF6F724035084A72518ABB626D257669F8CBC51A2F84B489354D76F52DA01C22E72F237F7517B9FF002702BC6F53F137897E20B979DA5D07C38C7F776C87135CAFAB9F43F97B1EB5C646356F14EBDFF095F8DE71717679B4B01FEA6D13A8017D7FFD6727A74DFDA3FED57B14B071A4BBCBBF45E9FE660E6E5E874DA4AD86956C2DEC20485075C756F727A9AB4FAA2A296690281DC9AF32D6BC6B65601D239564917EF1DDF2AFD4D79D788FE26DD48CC967F3BF691FEE8FA2FF008D74D3C0D4AAF44673AD1A6AF267D0977E28B2B78DA479D422F5766DAA3F135C9EB1F18BC35605945EB5D38FE1B64DFF00F8F1C2FEB5F385CEA1AEF886F5622F777F3B9F922405CFFC0547F415D6683F06FC7BAB8591F4D8F4F8DBA3DECA10FF00DF232C3F2AEC59651A6AF567638A598393B538B6773A87C7F8D49165A2CB20ECD2CE17F400FF003AC997E3F6BA4FEEB49B151FED33B7F515A1A6FECE97AE01D47C536D09EEB6F6AD27EA597F956CC3FB3AE8607EFBC4BA839FF62045FE79A5C996C3CFEF27DAE365B248E5A2F8FDAE83FBDD2AC587FB2CE3FA9AD5B0F8FE8C40BDD1648C7768AE037E840FE75A937ECEBA1907C9F12EA087FDA851BFC2B1F52FD9D2F1149D3BC536D31ECB3DA98FF50CDFCA8E4CB65E5F78BDAE363BA4CEBB47F8C7E1ABF2AAD7AF6AE7F86E536FFE3C32BFAD75B69E29B2B98D648EE1191BA3AB0653F88AF9CB5EF833E3DD243491E9D16A31AF57B29839FF00BE4E18FE55C8DBDEEBBE1DBD68B75E585C29F9E375287F153D7F114FFB328D457A53B951CC251D2A45A3EC94D515D77248181EE0D56D57EC3AA5B1B6BF81278CF4DDD47B83D41AF9D7C39F136EE3654BDF91BFE7A20E0FD57FC2BD1744F1BD95F04496548DDBEEB06F95BF1ED5C753035293D51DD0AD1A8AF1675163AFF88FC079368D2EBBE1DFF96B67237EFA05EE50FB7E5EDDEBD23C2375E1DF19696356D0B556B8809C3C640F3216FEEB8EC7F9F6CD796FF688FEF5733731EABE1CD7878AFC13702D3505E6E6CFFE58DE2752ACBD327FFAE3079AF371B9450C725ED62B996CFF00CEC74D1C4CE8FC0F43E946F0B5B11FF1F12FE42A097C24A47EEEF083FED479FEB54FE1678FF4AF1FE862F6C8FD9AF60C25F58B9FDE40FF00D54F383FC8822BB4CFB57CC56E1EC126E33A766BCDFF0099DB0CC2BEEA7F91C15FF86AFAD8164559D47F74F3F95624B09562ACA548E0823915EB040F6AC8D77478750899D5424EA3861DFD8D7CEE67C254DC1CF0CF5ECFF467A585CDE49A8D5FBCF3878BDAA355922956589D91D4E5594E0835A52C2C8C55948607047A542F17B57E7D5306E12D373E82355491DA7837587D4AD1E1B8399E1C066FEF03D0FD6BA2EF5C4FC3F8592FEE241F756300FD49E3F91AEDABF5CE1DAF56BE0212AAEF2D55FBD8F90C7D3853AF250D85A28A2BDD38828A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A29ACCA8A59880A06492780280307C75E27D33C1FE1ABAD7755936C308C2203F34AE7EEA2FB9FF13DABE4EBCD6350F1578965F167889C35D49C5A5BE7E4B58FF85547F9EE7A9AB7F1BFC79FF09AF8C4A5BCA4E87A6394B44CF133FF0014A7EBDBDB1EA6B91FED2FF6ABE9B0182F634F99FC4FF05D8F3E557DA4AEB65FD5CEC5B535552CD20000C924F4AE1BC57E36B8B866B4D358A4038697382FF4F41597E20D5E49D0DBC6C445FC58FE23E9F4AC11196396AF568E192D6461531165A04F3DDDEC8B182ED938551CE4FB0AF55F879F05E6BD48F50F154925AC2D865B48CE2561FED1FE1FA75FA56CFC19F065BE9F691788752855EF251BAD5187FAA5ECDFEF1FD057A8FDA7DEB0C4E2E51F72969E6631A2A6F9AA6A4DE1DD1B44F0F5A8B6D1B4CB7B34C6098D7E66FF0079BA9FC4D6AFDA6B13ED3EF47DA7DEBCA717277675A924AC8DBFB4D1F69AC4FB4FBD1F69F7A9F663E736FED347DA6B13ED3EF47DA7DE8F661CE6DFDA6B33C41A468BE20B436BACE996F7B1E38F3132CBFEEB751F81AAFF0069F7A3ED3EF54A328BBA13927A33C63E217C157B549350F0A4925C4432CD6721CC8A3FD86FE2FA1E7EB5E4D0CD77632B44DBD0A9C32B0C1047A8AFB03ED3EF5E65F197C196FAA59CBAFE9B0AA5F42BBAE1547FAE41D4FF00BC3F51F857AB85C649FB9535F339254541F3D3D0F3DF0AF8D6E2D596D3506325B1E164EAD1FF0088AEED3544750CB2065232083C1AF133195395ADDF0F6AF25BAFD9DD898BB67F80FF00856F5B0C9EB136A788BAD4F418355D47C37E2387C59E1B90477D0F1710FF0005D47FC4AC3BFF00FA8F502BEAEF87DE2BD33C67E18B6D734C7F925F965898FCD0C83EF237B8FD460F7AF8BBFB4BFDAAEABE0BF8ECF823C66B24B211A2EA2C23BE8FB467B4A07A8CF3ED9F6AF2F1B81F6D4F997C4B6F3F2FF2378D5F672BF47BFF0099F65507A54514892C6B246CAE8C032B039041E8453A5758D19D8E028C9AF9793E55767A0B5388D7620355B8DA382F9FC6B39A2F6AD5BACCD3C92375762D57B43D33CE985CCCBF221CA83FC47FC2BF3A965F2C5E25A82F89B7E8AE7D2AC42A14939744687866C0D9580DE3124A7737B7A0AD7A2815F7D86A11C3528D28EC91F3B56A3A937396EC5A28A2BA080A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A28012BC83F69EF1A7FC239E0B3A459CBB2FB540C8C41E5201F7CFE390A3EA7D2BD789C03F4AF8B3E32EB5278CBC77A8EA08E5ED237FB3DAF3C794848047D4E5BFE055E965787556B734B68EBFE471E2EA351508EEFF002EA7991BD6CF008A9219DE446724AA2F526B4DB490012578159BA8E1592D23E07DE7FA57D846D2D4E0AB2E5B538F5FC8A7832BF984607F08F6AD4F0E58A5F6B96366E3E49A7447FF00749E7F4AA8895A5A14E2CB57B4BB3D21995CFD01E6AA4DD9D8CADA9F4324EA8A15405503000EC297ED3EF58D15C34C5445972DF742F39AD38B4ABF650CFB22CF663CFE95F299866982CB9278AA8A37EEFF004DCF430F85AF89D2945B26FB4FBD1F69F7AAF71A6DFC485D556551D761C9FCAB34DCE0E09AAC066184C7C79F0D5149793BFDFD89AF86AD8795AA45A7E66D7DA7DE8FB4FBD62FDA7DE8FB4FBD77FB330E636BED3EF47DA7DEB17ED3EF47DA7DE8F661CC6D7DA7DE8FB4FBD62FDA7DEB42C34FD42F10491C5B233D1DCE01FA572E2F1387C1C39EBCD463DDBB1AD1A552B4B969C5B7E45AFB4FBD21B804107041EA2926D1EFA35C8689CFA026B2AE2492DDCA4EA6361D41AE5C066F80C7C9C70D56326BA27AFDDB9AE230788A0AF520D23C2FC5362963E21BFB4886238E76083D173C7E95918313F980647F10F515B9E22B817DADDE5DAF2B2CCC57E99E3F4ACC74AFAE837CAAE79CD6A3A59DE38D5D49643D08A8FEDAFE84D49A760BB5AC9C83F327F8569AE92194305C83532B47534A32E6BD3974FC8FA93F65BF1B1F10783CE8379317BCD2D408CB1E5E03C2FF00DF27E5FA6DAF51D62E372F90878FE23FD2BE3BF845ACCBE0CF1BE9DAA162B6A64F26E876313F0DF970DF502BEC98ECD5E52EE772672BCF5AF85CFF000B38CFF75B4BFA67AD97558ABC67BC7F2E850D3F4EF3984928C463B7F7AB75142A85500003000A50000001814B5E66170B0C3C6D1DFAB3AEAD59557762D14515D66414514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450072BF15B567D1FC05AADDC4C567784C1091D43BFCA08FA649FC2BE5DB7D136408A579C73F5AFA0FE3816BAB6D2B4A5E44B706671ECA303FF423F95705FD8FFECD7B7819AA547D5DFF00438251E7AEE5D95BF56793F89ED96CACC0231BB24FB28EB5E791666779D8732367E83B57A87C625FB2594AA386629027D4F27F4CD79BC4802803A015F47877FBA4CF35DE55653F92F9022D4CAB428A7819AB6ECAE5AD59F447C25D166B4F08D95E5F0DF733C7B9091CA447EE8FCB07F2AEAA687DAAF693144746B23063CA36F194C74DBB4629D2C55FCCDC40EAE3F1752AD5D5B6FE4BA2F91FA665F1850A3184364623C6C872B5CB78C6CD638C6A112ED39DB281DF3D0D7713438ED5CF78DA348BC397B23F0020C7D72315C3C31531196E6D46545E92924D774DA4FF00CCACD29D3C4E126A7D136BD51C07DA3DE8FB47BD51D321B9D46ED6D6D54BBB727D147A9F6AEF749F07E9C91837B2CB7327701B6AFE9CFEB5FBBE73C4F9764D250C44FDF7AA8AD5DBBF97CCF83C0E5788C626E9AD1757B1C87DA3DE8FB47BD77975E0BD1AE2322033DB3F66572C3F106B82F12E937DA0DD88AEB0F13F314ABF75C7F43ED4651C4F97E6D3E4A32B4BB3D1FCBA31E372AC4611734D5D77474DE02D2D755BE92E275DD6F6F8241E8CC7A0FA77AF4192200600C0AE67E0E3C52F876E8646F174777FDF2B8AED248EBF36E31A953158F9C24F48E8976FF873E9F24846961A2D6F2D59912C5ED5CEF8C3443ABE87756B0BF9570D1911483B1F4FA1E95D84917B556962F6AF89A11AD83C4C6BD17692774CF6EA2856A6E13574CF916689A391A3752ACA48607A822A074ADEF1888FFE12BD5BCAC797F6C9718FF7CD63B0AFEA9C255756842A356BA4FEF47E5B560A3371EC519B744CB3A7DE8CE7F0EF5E83E16B75BDB4200CE30CBEEA6B869101041E86BD17E0E2FDAAD6343CB46EF037F31FD2AF10FF0076CC55E356335E8FD1FF00C12CCFA26F85942F24715F4DFC20D55B56F87FA5CD33969E08FECF293D7727CA09F72003F8D794FF0063FF00B35DD7C10DD67FDADA5370AB2ACE83FDE183FC96BE771D25568BF277FD0F4631E4AEA5DD5BF547A6514515E19E8051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140076A4E2919828C938155DA4691B6A8E2B394D2F51C62D9C7F8C2D4DF78B2DC632B0DB67F124FF0088AABFD91FECD75334110D7A69E5E9E5051FA54E45A63AD772ACD24976308C2CDBF33E46FDA0C797AAE9B6A3FE5A4B34C47D3007F335E72A2BD37F69884C1E3DD320C600D38BFE7230FE95E68057D8611DE8C59E4CA3CAEC2A8CD4AAB48A2B77C1DE1AD53C53AC2699A543BE43CBBB70912F7663D856956A469C5CE6EC91318B93B2DCF6AF81BE2FB6D5B4187C3F7932A6A3649B220C7FD7443A63DC0E08F400FAD7A2CB0FA0AE43C27F06FC3DA308AE2FA6B9D46F5086DE2468911BD542907F335DEBC01542AE70063939AFC5F8828612AE265530D7B3D5E9A5FCBC8FB5CBAA558D251ABBA322588E3A5790FC5FF13C523AE85A74AB26C7DD74EBC8C8E89F8753F87BD7B45ED9C7736D24128631C836B0562A48FA8E6BCABC79F0DF4AB2D22E754D23CE85E05DED096DEACB9E704F23039EA7A563C334303471F0A989BF327EEE9A5DECDBFC8ACD275E78771A56B75EF6F225F859A588BC32BA84883CEBB62C4E3A282401FA13F8D74CF1B21CAD52F86924773E0BB354C6E83744E3D08627F9115BB2C3ED5F1DC5986AB88CCEBD4ABF15DFDC9D97E07AF94CA10C2C230DACBFE09462988383C1AA5E2FB24D5BC37776C541916332427D1D4647F87E357E683DB154354B9FB169B733C9F752263FA5787956271384C6D270DD495ADDEFA1D98AA74EA519296D6773CD3E1678AC687AD3DBDFC812CAEC0566ED1B8E8C7DB920FF00F5ABDDADE70D1AB070EAC3208390457897813C1B65AC5BBDF6A4F288964289121DBBB18C927D39ED5EADA25A5A69764967631B47027DD52ECD8FCC9AFD338CB1D97CF1BFB993F6AB492B69A79DF75D4F9CC928621505CE972BD577374A861915C8FC48F135A785F449252EA6FA552B6B0E792DFDE23D075FD2BA68A5F7AE5FC55F0F7C3DE249A4BA9C5CC17B27FCB74999BE99562463D862BC8CA25819E2632C5DF956F657BF93F2F43B716ABAA4D51DFCCF9A252D23B3B92CCC7249EE6A1618AEBBC7BE0BD4FC2578A97589ED2527C9B941F2B7B1F43ED5CABAD7F40E131347114954A32BC5ECD1F9F56A73A73719AB3457615E85F007E7D6F53B53FC0619D47E241FE95C011DABD1BF66C84CDF11AFA0C643698CDF9489FE34F14ED464C98479B43E87FEC8FF0066AC7846D0D8F8BE4E3027B523F1047F85750A2D368C9E7150470C475DB69E2E81194FE46BE3FDB369A7D8F5E50BD9F99B745145729B05145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400532591635DCC714D9E558632EC700564CD72D3C9B8F4EC3D2B9311898D256EA6B4A8B9EBD0B4D2BCF20007D055D863F2D71D4F7350D9C1E526E6FBE7AFB559AAA106973CB762A925B476303C613ADA5B24E0E0B4814FE47FC2B981AD0C8F9EB4BE2C4A61D01A51C08DD1FF5DBFD6BC97FB6BFDBAF6B0D479E1739A52B339AFDAD5107C45D1E68FEE49A56D1F512B9FEB5E4A82BD57F69093EDD1F8535753906392073EE4291FC9ABCB5057D2E015A845763CAAEEF3B8F515F527C1CF0F45E19F0940AD1817D78A27B96C73923E55FF8083F9E6BE6DF0C5AADEF8874EB471959AEA3461EC58035F57C33720D7C471C6692A0A9E1E2FE2BB7F2D8F6F21C2AA8E751F4D11D0C530239E69E555864564C33FBD5D867F7AF8AA3898CD599EF4E938EC12C0339C7359DAAADAC7633B5F34696DB0894C870BB4F0735B1BD194962063A935E51A8C97BF10BC50F61672B43A2D9B659C7461D377B93CE0761F8D7A584CB238A939B972C63AB7DBD3CDF438EBE29D24A295E4F44BFAE8737E15D59340F115CDB69E65BFD3269300221DC476603D474F7AF5252B342B2A2B8471901D4A9FC41E455ED2B41D3B48B61069D6A91003E66C659BDC9EF52CB17A8AE0E22AB431D514E9C1A6B46DBD5DBAB56DCE8CB29D4C3C6D295D76EDE862CD1ED04115E69E3FD5E4BA9069B1453416C1B323C8854BFE07B0AF5B962F6ACFD42C2DEF21686E6049A33FC2C335E2E515B0D9762D622B53E6B6DADACFBDADAB5D0EDC6D3AB8AA2E9D395AFBF9F91CCF8796C93498134C99658635C6475CF7C8EC6B5629F0707835C76B5A6DCF84B534D474F2CF65236D7427A7FB27FA1AEAEDDE2BCB48AEADDB31C8A194D78FC4993CF0D358FC34DCE9D46DDDEF7DDA66F96E31544F0F563CB38F45B5BA35E468C537BD5B8A5F7AC55768CE0F4AB50CD9EF5E361730D6CF73D0A9449BC4BA55A788741B9D2AF14149930AD8E51FF008587B835F2BEA569358DF4F6570BB6682468DC7A10706BEAD8E6E3AD7CFBF192D96DFC7D7CC830B30497F12A33FA835FB278759ACA7567856F46AEBD5349FDF73E3F88F0A9423556F7B1C438AF55FD93E346F8A3A8CB20F963D1DC1FA9963AF2D715E9FF00B36BFD8AEFC51ABB1C08E08E153FF7D31FE42BF51C72BE1E48F96A0ED3B9EECFAD0DED87E335D2783AE16F2192627255F68FCBFF00AF5E25FDB5FEDD7ABFC2394CFA1F9A4E43C8EC3F0C2FF4AF99C4D0E485CF5632BB3B8A28A2BCE350A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A6310AA589C01C934EAC8F135E7D9ED161538694E3F01D6B9F15888E1E94AA4BA174A9BA93515D4A57D7E6E263B4E235FBA3FAD4DA4279B7433CAA0DC6B0D25ADEF0C10C67F51B7FAD7CC60710F15898F33DF5FBB53D6C452F6345F29BB451457D79E31C87C56B56BAF095E220CB185C28F560372FEAB5F2BFF6D7FB75F63F8820371A45C228CB2AEF5FA8E6BE0EF1EAC9A0F8BF52D2C821229C98BFEB9B7CCBFA115EF64E94D4A07357D2CCED7C637435BF85CC54EE9B4DB85947D01E7FF1D63F95703110E8AC3A11915A1E07D711EE6E74BB93986F222A41EE7078FC89ACBB48DADE49ACA4397B77299F51D8FE55EFD2872268F2652F7DC7E7F79D4FC3C507C67A567B5C06FCB9FE95F46C13FBD7CB56EEF1B878DD91C72194E08AD18350D438FF004EBAFF00BFCDFE35F11C51C255B3AC4C2B42AA8A8AB5AD7EB7BEE8F6F2ACDE182A5284A17BBBEE7D430CFEF57219FDEBE64B6BFD438FF4EBAFFBFADFE35A56F7D7F81FE9B73FF7F5BFC6BC5A5C055E9FFCBF5F77FC13B67C454E5F63F1FF00807B8FC43D55EC3C2772626DB25C620523B6EEBFF8E8356BE1869D1E97E14B7CA812DD0F3E43DCE7EE8FC063F5AF1259EE675549AE269467203B9233F8D7BE69E4436B0C2BC08D1547E031559AD0793E1218772BB9B6DBDB64B4FC49C15558EAF2AB6B28A497CCDB640C322ABCB0FB52433FBD590CAE3DEBC0F72AA3D5F7A0CCC962F6AAB2C35B32C555658BDAB82BE10E8A758E6F5CD363D434EB8B2900C488403E87B1FCEB8FF871348D6777A6C99DD6F26E507B039C8FCC7EB5E932C58EA2BC37C5025B2F116A290CAF17FA43FDC6238CE7B57AF91E57FDA586AF97C9D93B34F7B34ECDFCD1E76638BFAAD5A7894AF6BA7E68F4F9A1F6AA8F1B4672B5E3D717B7DFF3F971FF007F5BFC6B36E6FAFB9FF4DB9FFBFADFE35955F09E551DD62127E8FF00CC71E2D8AFF977F8FF00C03DDA29B3F29E0D78EFC7000F8BA271FC56684FFDF4C2B9A9AFEFC1FF008FEB9FFBFADFE3542EE69A77DF3CB24AC06017624E3F1AFA5E14E06AF91E33EB13AEA6ACD5AD6DFE679B9AE7B0C6D1F66A16D6FBFF00C0294842A963D00C9AEF7C0F74344F86324CC76CDA9DCB49EF82703F4527F1AF3DBD5798C76717FACB87118F607A9AD6F1B6B71C0F69A45A9C436710181EB8007E807E75FA355873A513E7A32F7E31F9FDC745FDB5FEDD7D47F08AD5ADFC21661C61BC95247A337CC7FF0042AF8BBC1625D77C55A7694B92279D449EC83963FF007C835F797866DFECFA2DBA95DA597791E99E9FA62BC1CE12A718C0F5A86BA9A74514578074851451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400DC571BE329C8D4D109E163047E24D7675C9F8F2C2478E3BF854B79631281D76F507F0E6BC1E238D57809BA6AF6B37E8B73BB2D9463885CDD4C249BDEB5FC3FA82DB5E7EF0E11C6D63E9E86B968E7CF7AB31CDEF5F9E60B3574AA46A45EA8FA4AD85538B8BEA7A82B065041C8ED4B9C8AE134CD6EEACF0AAC1E3FEEB738FA7A56EDBF896D9C0F3629233FECE1857E8784E20C26222B9A5CAFB3FF33E6EB65D5A9BD15D1BDD4735F227ED5FE156B1D6A2D6618FE407C99081FC072633F87CCB9FA57D509AEE9CDFF2DD87D54D719F18344B1F16F84EE92322568A3225014E7CB3DC67BA901BF035F4995667423888F24D3BF668E1AF87A9C8F9A2D1F0BDBCAF04E934676BA30653EE2BA3BE98C8F06B083F7722849B1DBD0FE1D2B1357B0B8D2F53B8D3EE976CB03956F43E847B11CD6A7842EE03336977B836F75F2827A2B1FF001FE78AFBEBA6AE78189A6D5A71DD7E2BA9AB160804735720EA2B384136977CDA5DD64E3E681CFF001A7F88AD080F22B2645D3D51A76DDAB4ADFB5665B1E95A56E7A562C4CD5B3204A85BA02335EAF078B347ED72DFF7EDBFC2BC9206E95A16ED5E06719150CD5C1D6935CB7B59AEB6EE8EEC16655706A4A9A4EF6DFCBE67ACC1E2BD20FF00CBCB7FDFB6FF000AB91F8B7475EB72FF00F7EDBFC2BCA219315696504735E643847074F694BEF5FE46F2E20C43DE2BEE7FE67AAA78C341C61AE9FF00EFD37F85325F18F86C75BC7FFBF2DFE15E58F28C7155267AE85C33856ACDCBEF5FE466B3BC45F65F89EA5378D3C343ADEBFF00DF97FF000AF23F195D5BDEF882F6EED1CBC32C9B9188233C7A1A8AE1AA85C1EB5DD96E4B430151D4A4DDDAB6B6FF00233C566157150519A5A6BA19F71DEB36E7BD68DC1E2B36E4F5AF751C48CE9BBD5593A7356A6EF59ED0CFA95EA697699DCE33338E91A77FC6B543D16ACAF632ED9A6D5DC7EEA1529167B9EE47F2AE76EA692E6E249E5397918B135BBE2FBA82391348B1C7916DC391FC4C3B7E1FCEB1F4BB2B8D475082C6D5374D3B8451F5EE7D875AD6E922F0D06DB9CB77B7923D9BF655F0B36A3E2093589633B01F22238EDC191BF2C0FC4D7D86000B803815E77F037C2B0F873C236FB530CF1858C918257A96FAB1C9FCABD101AF89CC711EDEBB6B647B74A3CB1168A28AE1340A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A42011834B450D5C0E6B53F09E9B78E65837DAC8793E5FDD3FF0001FF000C56637832ED0FEEEFA261FED291FE35DBE290D78388E1ACB6BC9CDD3B37D9B5F82D0EDA7996229AB2969E7A9C4AF84B5207FE3E2DB1F56FF0A9E3F0B5EFF1DD423E809AEC28C5443863010DA2FEF65BCD310FAAFB8E6A2F0C15FBF779FA27FF005EB42C346B7B572E2491C952A4311820F62315AB47E35DF432AC2E1E4A54E1AAF36FF339EA62EAD456948F92FF00699F877269FA89D5AC212632A59081F7A3EEBF55FD45780D7E8DF8AF43B3F10E8B369D76301C663900E637ECC2BE22F8BFE04BEF09EBF70AF6FB230D960A3E500F475FF64FE878AFBBCA71CAAC7D94DEA8F2ABD3B3BA24F0DDCDAF8BF4A1A26A1308755806EB4B83D5F1FCCFA8EE39AAB17DAACEF9F4DD4A2305EC5D41E8E3B32FA835C5412CB04C9343234722306475382A47420D7AA787F57D17C7F63168BE20916C75C8F8B4BD5C2EF6EDF8FAAF43DB9AF52A5E1AF4FC8F3DD2E5DB6FC8AB6EDD2B4ADDAB2354B3D57C317C2C3C41079618E20BB41FBA987D7B1F6357ADE404020E476A8D1ABA336AC6C40F576192B26193A1AB7149EF50D10D1AF14B532CBEF59692D4A25F7A8B10E25F697DEABCB2D57697DEA2797DE8B0288B33D529DE9D2C9552692AD22D220B86ACEB96EB566E2418249AA3A65A6A9E25BF3A7F87E0F34A9C4D74DFEAA11EE7B9F6157A25765A57284A6E6EAF23D3B4E88CF7B37DD41D147F79BD00ABBE219ED7C1BA4B69367309F59B91BAE671D533FCBD87E35ABAEEA9A2FC3DB19749D1245D435F9462EAF1B07CB3EFEFE8BDBBD7955C4D2DC4EF3CF23492C8C59DD8E4B13DCD5D3BCF5E9F996A9736FB7E646492727935EF1FB347C3B9355D4D755BE898461433123EE467B7FBCDFCB35C0FC25F035F78B35EB78E3B732465B2AAC3E56C7566F451DFD7A57DBDE11D0ACBC39A2C5A75A0CEDF9A5908C191FB93FE1D857999B6395287B383D59E8D0A777766C468B1C6B1A28545002803803D29F4515F2A76851451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140082B96F889E0DD3BC63A3B5A5DAAC77280F913EDC9427B1F553DC575345384E5092945EA26AFA33E04F893F0FF0056F09EB735A5C5B34783B82F552BFDE43FC4BFCBBD7222CE707214835FA15E33F0AE8FE2BD29AC356B5F3179314ABC49137AA9EDF4E87BD7CBFF00127E186AFE0DB97B89A037BA516F92F614E17D9D7F84FE9EF5F5383CD5564A33D25F99C7528F2EAB6307C17F10E68AC3FB07C69A7FF6D690E366F750D2C63F1FBC07E047AD7472FC3A5BEB36D5FE1AEB506A367D5B4EB993E64FF6558F2A7D9B1F535C8DBE9115C462484ABAFA8ABDA5D9EA1A55DADDE9B733DA4EBD24898A9FA71D47B56F3693BC1D9F6E9F718FB24CAD737D71A4DE7D8B5FD3EEB48BAFEE5CC642B7BAB7423DEB4ADAE239503C722BA9E854E41AEEB4DF1FDF4F67FD9FE2CD12CB5DB43C37991A863EE41054FE4298FE12F83FAEB996CE7D57C2D74FC9F25D9573F43B971F4C567F5BE5F8E3F76ABFCCCE585BFC2CE4526F7A904DEF5D57FC29BBD9BE6F0F7C4AD3EF50FDD4B88149FC59589FD2A293E0C7C4D43FBAD4FC3138EC4BCCBFFB2D52C661DFDAFCD19BC254EC73466A89E5F7AEA53E0CFC4D7389752F0C403B90F337FECB52FF00C29ABF87E6F107C49D3AC507DE4B78141FC19981FD2878CC3AFB5F9B0584A9D8E22E2E12342F23AAA8EA58E00ACDB7BE9B54BCFB16856375ABDD9FF9676B19603DCB7403DEBD293C1FF08343612DF5DEABE29BA4E40924664CFD06D5C7D49A7EA1E3EBBB5B33A7F84742B1D0AD070A52252DF50000A0FE06A7EB7CDF045FCF45FE66B1C2DB7673707C396B4B31ABFC48D660D2AC47234FB7932F27FB2CC3927D973F515CFF008CBE21BFD80E81E07D3FFB1B4A4057CD55DB2C83BE31F773EBCB1F514ED5AD751D5AEDAEF53BA9EEE76EAF2B163F41E83DAA84FA4450466498AA28EE6AE0D3779BBBEDD0BF6491E7C6D2E09248249E4935D57C3BF01EABE29D6A1B4B7B569771CEDE836F7663FC2A3D7F2AEF7E1CFC32D5FC6772B35ADB9B4D295BE7BD993E53EC83F88FE9EA457D45E0AF09E8DE12D2858E956FB49C19A76E6499BD58FF004E82B2C6E691A2B961ACBF236A7479B57B14FE1BF82B4FF0668CB6B6C165BB751E7CFB71BB1FC2A3B28EC2BACED4502BE567394E4E527A9D8924AC85A28A290C28A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A002A396349A368E4457461865619047A115251401E43E38F827A65FCF26A5E13BB3A0EA0793101BADA43EEBFC3F8647B5795EB965E25F084863F17F872E16D94E06A1663CD84FB923A7E383ED5F58D35D15D4AB2865230411906BB6963E715CB2F797E3F799BA69EDA1F2DE8D3687AC01FD9FA8DBCCC7FE59EEC3FFDF279AD5FEC23FDDFD2BD4FC53F087C01E21679AEB428AD2E5B9F3EC4981C1F5C2FCA4FD41AE3EEBE08EBBA712DE15F889A95BA0FB906A1109D7E99C803FEF9AE9589A73FB56F55FAAFF223924BA1CE7F611FEEFE952A6957083092CABF4622AF4BE0EF8DBA77FAA9BC29AC28E858BC6C7F0C28FD6A06B4F8D5170FE00D1E73EB15FA28FD64A776F69C5FCD7EB60F9103E9572E30F2CADF5626A2FEC23FDDFD2AE2DAFC6A9785F0069101F593504207E5254F1783FE36EA1C49278534753D482F230FC30C28BB5BCE2BE6BF40F9199FD847FBBFA565EB12E89A429FED1D42DE061FC05B2FFF007C8E6BB6B5F825E20D4483E29F88BA8CC87EFC1A7C2205FA67383FF7CD761E17F83FE00F0F32CD6FA1C77B72BCF9F7C7CF627D70DF283F40293C4D386F2BFA2FD587249F43C3344B4F11F8BA511F83FC39713404E0DFDDAF9502FB827AFE193ED5EA5E08F823A6D9CF1EA5E2FBDFEDCBE1C8802EDB68CFFBBD5BF1C0F6AF5E4454408802A818000C014EAE6AB8F9C95A3EEAFC7EF2D524B7D48E18A286158A0448E3401551060281D801D2A5A28AE2340A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A0D20A6CFFEA9FF00DD353295A3704AFA0CF3A2070644F6E697CC8C73E62FE75E1F147AA5FEB76FA4E936D6F3DC4D0CB3133CE625554280F215B27E71DAB57FE112F1BFFD03F45FFC18BFFF001AAF9EC366B8EC4D35529E1D38BDBDE4B6763D2AB84C3519B84EABBAFEEB67AE8618CE463D6A836B3A524BB1B50B50DD3FD68FF1AE0647D5744F065BE85AA5DABEA724B2492AC4ECEA91348E5503103200DA3F0F4AE5F4DB1BFBE13CB7BAED8E909E632C309B17B8728380CCC1D40CF27033DA8C56752A7887421CA9A5AB6ECAFD977614301CD47DAC9377D92577EAFB1EDB7779696A035CDC45086E9BDC0CFD292CEF6CEF031B5B9867DBF78A386C7D715E21731EA115B69FA659DEC73084476A6F2E55B6A448B8F3360393903EE83D4F5EB5AD65E1E97EC5AD5EC1E37B36B88B46B983296ED6A2DCC8A08998EF621418F820763E95B65B9ACB1D88E48D946ED6FAB4BADBB5CCB1984585A3CCD372B2E9A6BD2E7B04722383B1D5F1D70738A78FAD7857ECA5A7D8595BEBA61F125B6B13B3C419604942C6A3760E6455C9273C01C63DEBB9F881F15FC1FE0BB8FB26A57B25C5F60136B6ABBE4407A16E405FA139F6AFA8AB85946B3A54EEDFA1E251C6C254156A968A7E773BC181CFEB4CF313CCD9BD777F773CFE55E63E10F8E1E0AF12EAB0E970BDED9DCCEC122173100AEC4E02E413C9F7E2B82F046996E9FB485E5DDE78D2CEF6E7ED572D15A426591F243623662A106D53838270571570C04FDEF6978D95F6BDCCA799537CBECAD2BB4B7B58FA3C52F6AE6FC6DE32F0FF008334E17DAFEA096CAD9114606E92523B2A8E4FD7A0CF26B82D27F685F025F5FC76AE9A95AABB6D12CD0AED1EE70C4FE59AC29E12B558F3422DA3A6AE36852972CE6933D828AC9D575ED1F4BD0A4D76F6FE08B4D4884A6E37650A9E8463AE72318EB9AE0757F8F7F0EEC2D1268751B8BF772408ADEDDB70F73BB000FC6952C355A9F045B2AAE2E8D2F8E497CCF54228EDCD703F0EFE2B785BC6F76F65A64B3C176A0B082E502B381D4A90483F4EB5D2F8ABC45A2F8674A7D4B5CD421B2B65E3739E58FF007540E58FB0A53C3D484F925177EC3862694E1ED2325CBDCD906826BC613F68CF02B5CF95E46ADB33812181707DF1BBA57531FC5DF87EF73696C3C410996EC298D446E71B8E006E3E53F5EDCF4AD5E07111DE0CC2398E165B4D7E4779DFDA97DE938201EB5CA68FF10FC23AC789DFC39A66AE975A8A170C9146CCBF2FDEF9F1B78FAD73C612926E2AF6DCE9955841A52695F6F33AEA28A2A4D488BA2E32C07D4D3D4E466B95F88F6267D1D6ED33BAD9F271FDD3C1FD70699F0CB5396FB4492D2EA5692E6CA53116739668CFCC87F23B73DF61AE08E32F8B961A4AD649A7DD75FB8E8961ED41554EF7767E475B518914E4875C0EA6B27C677E2C3429D95B124A3CB4FA9EBFA66BCBBC2B15D6A30789F52FB6345656C91D8DBA3BB79524E0891D98283D328B9C1C7CD5957CC1D3AF2A508DF962E4DB764BB2F9950C3274E3393B734925A5FD5FC8F6943919C83EE29C07A5709F0ABED56BA36A736A5A9D94E9F6BDE1209DDD2D94449953BD54A9382D8C7F167BD694FE36D2A390A2457328071B954007E9935A4F31A1469C255E4A2E493DEFBF6B7E64430D52A4A4A9C5BB36B63A9E3F1A8FCC4F30A798A1BFBB9E6B2F46F10699AB3B436D36260A5BC97187C0E0903B8E474F51EB5E13E09D32DE3FDA46F2EAF3C67677B73F6BB968ED2132C8F921B11B3150836A9C1C1382B8AF53031A78CA72AB095D24DAB2BDCF3F1B567859C6128EADA5ABB5BFCCF7ABCF1068963A943A6DE6AD6505F4D8F2EDDE75591F27030A4E4E4D6A718AF903E26DEDBD87ED3F35FDECC22B6B6D42D649646C90A8B1C649FCABD6FF00E1A2BC09F6D16EB16A8D1EEDBE7790A14FBE0B6715DD532CA9C909534DDD5D9C14B36A7ED271AAD4795D91EC6DD2991C89264C722B63AE0E6B9BF105FDB6BFF0E753BCD175982DA2B9D3E530DF994A243943F3330E576F7EE315E59FB2FD8D958E95E24683C496FAC5C304DE2DD250B1A856C1CC8AA49273C01C63DEB9E186E6A729B6D356D2DF9F63AAA62F96AC21149A95DDEFE5D1753D9EC35FD0F50D426D3AC756B2B9BB873E6C114EACE98383900E460F15A3E647E66CDEBBFF00BB9E7F2AF93FF64024FC48D4327FE61CFF00FA1A5749E0BD32DE2FDA42F2F2FBC6B6579722EAE5A2B588CB23E486C46CC5422ED5C838271B715D3572E50A928A93D237D8E3A39A4A708CDC56B2B6F6FCFA9F47E28038F5AF93BC67F13EFA6F8E6B3D9F8A2E7FE11CB5BD85008252B098976F99C0FBC33BB9E73F4C57B6C7F197E1CC914B28F11C616250CFBA17079206071C9E7A0F73DAB3AB97578C63249BBF65B1BD1CD284E728CA4A367D5AD7CD1E8B45637857C47A4789F4A4D5344BC5BBB562543852A430EA083C8EDF9D51F1A78EBC2FE106B74F106AB1DA4970098A3DACEEC0753850481EF5C8A94DCB9145F376EA773AF4D43DA392E5EF7D0E9E96AAD85CC37B6505E40C5A19D1648C9520956191C1E4706AD76ACDAB3B3344D35743481D3D690B617270052F27E95E63F14BC51F6533408D29820210A4432D3484E02003A9C9000F5AF3F1F8E5848295B9A4DA492DDB67461F0FEDA4F5B24AEDF648F409758D2A27D926A16AAC3A8328E3F5AB16F730DC26F8254954FF00123061FA5795691E07F185EDAADD5FDEE99A533AEE5B5F21EE1D33D9DC328CFAE01FAD3AEF48F16787AFEC5ED22967325EDBC6D3D9A174319917CCDEBD546DDDC9C81EB9AE58627318D482AB4572C9A4ECEED5FABF25E46B28615C25284DDD2BEAAD7F43D647BD07A5636B7E23D3B4A7F2A677926C64A463247D7B0A8349F166977F7096FBE482573851300031F4041C66BADE6585557D8B9AE6EDFD7532585ACE1CFCAEDDCE83F2341AA5AAEA567A640B35E49E5AB36D0402727F0FA562DD78D7498982C62698633B91703F522AABE6185C3BB549A4FB5F5FB82961AAD5578C5B474E7348703A9C566689ADD86AF1B1B49B2E80178D861973D323D3DC57CDDFB45FC43D517E23C3A6787BC4B7105958C4B1DC25A4C51567DE778623A90368F6E7DEBD2CBE8AC7B5ECA5A3D6E79D986256060E5516BB58FA9A901F518AE1BC2DF14BC15AFEA16DA4E9BAD09EFA643B53CA7192AB93C918E80D16FF0015FE1F4F71790C5E25B41F6452D33B06098071F2B1186E4F00673DAADE16B276E57F709632838A929AB7AA3B91D7AFE94BDAB90F09FC46F0778A6FE5B1D0F561753C51195C794E802020139603D4572BE23F8FBE03D1F527B1867BBD4DA338796D230D103E81891BBEA323DE947095E52E451770963B0F187339AB7A9EB359579E20D0ECB51874DBCD5EC60BE9B1E5DBC932AC8F938185272726B1FC01E3EF0EF8DACDA5D12E99A48FFD64128DB220F5232723DC135F39F8F7FE4EB94649FF0089A59FFE811D7461B02EAD4942A5E2D2B9CF8BCC152A51A94ED24DD8FAEA8A45FBA3E94B5C07A6828A28A00074A8E7FF0054FF00EE9A9074A8E7FF0054FF00EE9A8A9F031C773C3127BD4D66DE0D2EC2EEEEFE486464FB332ABAC60AEEE59978C95E335AB9F1DFFD0BFAFF00FE0543FF00C76A2F0E5CA699E2EB6D5E6567862B49E02A9F789768C83F4F90D771FF0009D69DFF003EB73F92FF008D7C26570CB7EAD1F6F51A96B7576ADABB68BCAC7D0E31E2FDB4BD9C7DDD2DA2EC8E52FEC7566F0AE9FA9F889678B53F3A685A291D09F28BB18F76DC82C1401C1EE739350782BC1CFE25D0DF54B8F10EA76A5AEEE22114090ED558E564182C84F451DEBA2F17EAF67ADF85AD6FEC199E16BA68CE570559372B03F4208AB1F063FE449FFB7FBCFF00D2892BD8A185C3D4CC6ADE09AE54D5D5FE7AF53CFAB5AA470B4D2935ABD9DBB76383BCD36493C4DA6E830EA173024F7EF6CF70A11A42AB1C8D9E576E4941DAB7FC4DE0E87C35E03F185EAEA77B7D35DE8D3C4CD702301552198800228EEE7AD516FF00929DA2FF00D85E6FFD133D76FF0016BFE49AF88FFEC1575FFA21EBAB866853853E7E45CDCCD5ECAF6BED739B3AA939DD393B28AD2EEDB763C07F660D4DF44F0878EF598D43C96566B3A29E859524207E60565FECF7E07D3BE22F88359D5FC5524D7A96E56478FCC2A679642C4B330E71C1E98E4FE7ADFB2F696DAE7847C75A3A3847BEB35B7563D14B24801FCCD725F09BC737DF0AFC4FA8596A9A74AC92379377013B5D594F51EE39FAE7F1AFD1EA45CE759527EFD95BEE47C15394614E8CAAABC2EEFF007FFC31EE917C05F0A5978BF4CD7B48B8BCB28ECA6599AD377988E5795C16E579EBD73ED5E49F0C3FE4E92EBFEC277DFCE4AEFECBE3F9D77C61A4685E1CD05FCBBBBB8E29A6B86CB6C27E6DA8BD38E724FE15C07C2FE7F6A4BA1EBA9DF7F392B9F0F1C446153DBBFB2EDF89D3899E1A73A7F575F695F4B76313E2A6BF67E24F8D57F2789A7BC5D1EC6EDAD0476F82E228895C267805882493FDE3547E23DDFC2FBCD3206F05699AA69D7D1B80EB33168E54C1C9E5890D9C74E3AD5EF8C7E1E7F08FC5CBEB8D674EFB6E9B7B74F7B082CCAB346EDB8AEE041041254F3DB3DC55D97C45F05574F1227807517BAC7FAAFED190267FDECE7F4AEFA7CAA9D39534DAB746ADF357479D5B99D4A91A8D277EA9DFE4D267470EAF36ABFB236A51DC485DACEF23B7527AEDF36361FF00A11A77ECDBF0C7C29E2DF0DDDEB5E20B496F244B93024466648D405073F2E093CFAE3DA892FEDF50FD9635D96CB43B7D22CD7518D608A22EDBD7CD8FE62EE49739C8CFB63B576DFB1E7FC938BEC7FD049FFF00404AE0AF370C3D571D1F37F91E861A9C6A6268C67EF2E5FCAFDCF1FB2D323F05FED1F6DA3E9AF325B5BEAD146819BE6F2DCA9C13DF86C7BD6DFED037777E2FF8E967E117B8686CEDE582CE2EE15A5DACEF8F5F980FF808AAFE3DFF0093AE1FF615B3FF00D023A97F692D3354F0BFC5F87C5B0237917862B88250BC2C91AAA95CFAFCA0FE22BA20E33AB4E527AB8BFBCE79C5C29D58C57BAA4AFE87AAEA7FB3E7802E7475B3B382F2CAE9530B78B70CEE5BD595BE53EE001ED8AF27F8FDF0D22F0368BE1DD474B9A49E3893EC975332805A504BAB6074C82C31D828AEC8FED2DA6A68CB20F0FCF26A58C18C4A163CFAE719C7B63F1AED238AEFE2BFC0A66D56D22B6BED42079605452151D1C9888CE4E0E073DC135C309E330B38CAB3F76FDEE7A13A782C5C251A0BDEB5F6B6C6037C5409FB3BB6BEB718D5847FD9C8777CDF682301BEA13E7FA8359FFB1FF850DAE917DE2FBB8FF7B7ADF66B52473E5A9CBB0FAB607FC00D7CF1A3D8EAFAB6A769E1484CBE64F7A2358189004A4EDC91DB1CFD39AFBCFC2DA35AF87BC3B63A2D9AE20B381624E31BB03963EE4E49FAD3CC1430B4DC29EF377F976272C954C6558CEA6D056F9F7FB8D4C734BD2BE787F8E3E209BE2D7FC22F0E99A7C560356FB0EE70CD215126C2D9C8193827A71EF5F43A9CF35E457C3D4A0A3CFD4F730F8BA7887250E9A10DDC09716D2C120CA4885187B118AF32F0A4F2685E395B598E23B92D6931E837824C6C7F1CA8FF00AE95EA79AF37F8A3A63A5F477D6EDE534C0624033B254C156FAF43FF0001AF9BCE3F712A78C5F61EBE8F73DAC0FEF14E83FB4B4F55B157E316B8B68B2EDCB8B384B6C1D5E46E8A3DCFCA3F1AD283433E1DF84F0695290D728A925D3FF7E777DF21FF00BE98FE18AE4F4E824F147C47D36CE61BA28243AA5E8ED843FBA5FC6420E3D10D7A6F8F47FC531739EC53FF004215C74A2EA65F88C54B7A89B5E89591A39258BA5496D0B2F9B69B382F04433DC786BC5F1DBAB33FF6846DB57A902084903F01591A3D8E8735F4F79AD3EA578926DFB3ADB5DB451C400E7E5523712727249F4C56EFC3EBF974DD13C55790223BC7A94670E091830420E7145F5EF85F5367BABFF0F18EF18EE69AD6631191BD58A904FE39AE2C43A54D509FB4519FB38AD6374D5BF0674D1539AA91E4BC79DECECD3B9ABE0BD034B4F13A6B9A66B334F1416B241F649A31E62798C8725B8C81B001C7AF26BC1FE17FFC9D25D7FD852FBF9C95EA1E0ABC924F891A6D8E9CC5CC31CD25F104131C050855623B973191D33B6BCBBE17FF00C9D2DCFF00D852FBF9C95F75C275655301294A0A378BD12B2DDEA979EE7C8711C231C5D38C64DDA4B56EEFA7E451F8B7A641AD7ED2777A4DC3C890DE5FDB4323210182B471838CF7AF46F8C7F057C27A4FC3ABDD57C3B633DB5FE9F1898B79ECFE6A0237EE0C48E993C63A579C7C60D51345FDA3AFB5792232AD9DEDB4E501C16DB1C6719EDD2BA8F8A9F1E6C7C47E0ABAD0346D2EE619EF904734B330C4699048503A938C76C66BE9250C43545D2DACAFAFA6E78319E194ABAABBB6EDA7AEC37E0AEA97377F01BC7DA6CCECD159D948F164F4DF13E47FE3A2AE7EC77FF1E1E2CFFAE717F27AD0F871E11BDF0CFECE9E2CBCD4A1782EF55B09A63138C32462221323B139271E84567FEC77FF001E1E2CFF00AE717F27A8AD38CA9D771DAEBF4BFE25D084E1568467BD9FEB6FC0C2FD8FFF00E4A46A1FF60F7FFD0D2A0F86AAAFFB505E46D9DAFA9DFA9C7A1F32A7FD8FFF00E4A46A1FF60F7FFD0D2A2F85FF00F274B73FF614BEFE7256D51FEF6AFF0083FCCC292FDD52FF001FF91CC78A3C19A669DF1B7FE10AB69AEBEC06F20B7F31D81942C8A849CE319F98F6AF6CD4FE0078774FF076BB0E9335F5DEA7716F9B66B9756D8E87700A0003E6C6DC9F5AF30F8D5752F873F68C9F5AB9819922B8B6BC8D738F315513A1FAA91F857B8FC30F8BD61E3DF14CDA369DA4CF0470DA19DA799C6490CA368503A73D73F8573E2A789F634EA537A2577F81D384A78575EA53A8B56ECB4F53CBBF647F159B1D6EF7C27792158AE54CB002785914723DB2B9FF00BE456536FF008BFF00B44704CBA45B4B8FF645AC27F4DEDFABD637C7DD06E7C11F152F2E74E67B6B7D401BBB678F2BB77E77A823A61B3F862BD7FF00648F09FF0065F8427F135CC605D6AAFB61247220438FD5B27F015A62254E1078B8EF2492F5FEBF233C353AB52A2C1CBE18C9B7E9FF000E7B8A80A0000003A52F15E4BFB41FC4DD53E1F26931695656B3CB7CD23179F2555536F000C727775CF6F7AEAFE11788EF7C59E01D3F5ED423862B9B9DFBD210428C39031927B0AF9F9616A2A2AB3D9BB1F4B1C5D375DD05F12573AD3D0D78EDCC0973F1234749C031AEB2EEDB877549597FF1F0B5EC7D2BCD7C7FA5CDA7EACBAB421A3864956459547FA99410467EA403CF04F1F5F9BCE54E9BA5884AEA12BB4B7B3D2FF23D9C0F2CD4E9376725A7AF6F99E92473DA91CED566F419AE2ED7C7B11807DA6C5FCD039F2DC6D27F1E47EB583ABFC43BB8F55D3EDEDDADE3373770C0B6CC773CAAF22AB1F5E1493C74C735A473BC2D571A749F34A5A2493EBD5F6B112C056A71729AB24AFAB317EC573E2AF19D9E92F7B3DAC37226BBBC9616024289B46C5241C659D79F40715BFE2FF0B69FA1DA4074E371E4484C6D1CD334BCE3AEE625BD7393E98C5550AFE18F1AADCC9116681648C2F43243263907EAAA7EAA4543F113C616F3DAA4B246F1C5112228B3BA49A46E02803AB1E800AF9AF6943EA0F0AE3FBFBED6D5BBE8EFE87ADC951625D7BFEEEDBDF44ADB7DE695D6A936ADF0E349BCB87324EB3BC12B9EAED13491963EE4A67F1AA3F0F3C1D63E20B0935CD66F350B966BB9E186DA3B8686289639590708416276E4924FD2AD9D2AEB45F85FA1D85F2EDBCDE66B85CE76C921791D7F02E47E15B5F06B3FF08463FE9FEF3FF4A24AF6E85084B32ABED126D4574BFAD8F327524B094D45B49B7FA6E72FA2CB269BF10F49B68A46092DD4D692E792E9E548C01FF81221FC2BC3BF68AF0C58787FE293DB59493B2EA282F66323024492C8FB80C01C71C57B7FFCD4ED17FEC2F37FE899EBCC3F6BB865B7F89BA65FBC6DE449A7C611BB12923EE1F5191F98AF7781DB84397A5E5F833C4E2E8F3C5CADAFBBF91EABE05F82FE18F086A0BE21B4BBD42EEED2D9C46B70EBB10B2E09C2A8C9C1239F5AF9DFE0A781ADBC7BE2F9748BDBE9AD2DA180CF21854177C10368CF03AF5C1AF79F067C6FD3BC55ACD8F86ECB46B98E5B9B793CE9A4906D8CAC64F031F374F6AF36FD8FFF00E4A46A1C67FE25EFFF00A1A57D5529622953AD2AAFDEB69F89F3F56385AB3A31A4BDDBBBEFBE86E7C5EF04689F09FE1FDE5CF8726D41AF35A9134E926B898131C2433B85DA07DED801CE783563E067C1EF09788BE1D43AD6BB6F3DD5DDF99363A4EC9E40562A3681C13C67E6CFD2BB9FDA83C3979E20F8652BD844F34DA75C2DE18D464B205656C0F60DBBF0AF25F837F1B6DBC1DE15FEC0D5B4E9EE6280B3DBBC2C030CF3B483DB3CE7DCF5ACE8CAB57C2374A579DF5D75B1AD7850C3E312A91B42DA69A5CC9F014373E01FDA1A2D02DE791E25D43EC64B1FF59139C027DF0C0FD6AE78F3FE4EB47FD852CFFF00408EA6F83763A97C44F8E13F8CAE2D4C5656F726EE56EAA8DFF2CA3071C9E9F829350F8F78FDAB80FF00A8AD9FFE811D76295EBD9FC4A1A9C5C96A1CD1F85CF43EB91F747D2978AF24FDA07E266A9F0FA3D222D2ACAD6796F8C8C5E7CED554DBC0031C9DDD73DBDEBACF845E23BDF1678074ED7B508E18AE6E77EF48410A30E40C649EC2BE6A585A91A2AB3F85BB1F530C5D39567417C495CEC28A28AC0EB0A08C8C1A28A00E42E3E1D7852E6E66B892CAF03CAE5D826A3708B9273C287000F61C533FE159F843FE7CEFBFF06973FF00C72BB1A39ACBD8D3FE543BB39F5F086809E1C1E1F4B29174E0E64082E240E18B1627CCDDBF2493DFBD5ED0347D3F41D3174DD2E0305AA3338532339CB316625989249249E4D690A2B4B2BB64DB6F239F3E10D03FE1228FC406CE43A8239911FED326C562A54911EED80E09EDDCD6A6AD616BAA69B73A7DF45E75B5CC4D14B19246E4652AC3239E413570D1447DDB581A4EF7EA73BE0EF07786FC2304D6FE1DD323B249D83CB87676620606598938F6E9D6A1F14F80BC21E27B8173AEE81697770001E69051C81D01652091F5AEA08F6A31C569EDAA7373733BF7BEA67EC69F2F272AB76B68733E18F02F847C33379BA1E816567360AF9CA9BA4C1EA37364E3F1A8B47F00784749F134FE23D3F458A1D52767792E0C8EC7739CB10189009F603A9AEAF9A31F851ED67ABE67AF98950A4925CAB4DB4D8CCD7743D235FB0363ACE9D6D7F6E4E7CB9E30C01F51E87DC572D6BF07FE1BDB4A258FC276658741233B8FC9988AEF08F7A3B510AB520AD1934BD425429CDDE514DFA189ACF86741D5FC3ADE1EBDD3216D28AA8FB2C798D00520803663182074A77853C39A3785F4A5D3342B08ECAD558B6C56272C7A9249249FA9AD81CD078A5CF2B72DF42BD9C39B9ACAFDCE56E3C01E129FC5C3C57368B149AC6F0FF683239F98285076E76E4003B76CF5ADBD6F48D2F5AD3DEC356B0B7BDB56E5A29A30CB9F5E7A1F7ABF8A3143A9276BB7A6C254A0AE9456BBF99C3597C23F87369702E21F09D89707204A5A45FF00BE5891FA56CF8DB5EB3F07F83AFF0059991043656E4C710F943374441E99240AE87B553D4F4FB0D4ECDACF51B382F2DDC82D14F1874241C8C83C75AAF6B29C93A8DB5EA4FB18C22D524A2DF91F357ECA7E1DB8D7BC67A9F8EB535321B77711C8CB8DF712E4BB0FA293FF007D0AFA83EB5534CB0B1D32CD2CF4EB482D2DD32562823088327270071D6AD8FE55AE2F12F11579ED65B246582C2AC353E4BDDEEDF99C65AFC32F045B7898F88E2D0621AA79ED71E7196461E63124B6D2DB73924F4E3B57678C76A419CF6A776AC25394BE2773A214A10BF2A4AFD85ACFD6F49B1D6AC1EC75185A581C8242C8C8C08EE194823F035A1456724A5A334307C33E14D0FC392DC4BA459BC325C85134924F24AEC173B46E762703278F7AD3D52C6DB50B09ACAEE3F321994A3A862A707D08E41F71CD5AA5A6D26AC25A6C62681E1AD1741B1B9B3D3ACD920BA7324EB2CAF36F2542F25C92780062B3AF3E1E783EE642EDA3AC608C6C8279614FFBE51828FCABABFC697F1A1C53DC0C9F0FE85A3E836AD6DA3E9B6D6313B6E711260B9F563D49F73591A3F803C23A4F89A7F11D868D143AA4ECEF25C191D8EE739620312013EC07535D663E9455A9CA37B37A90E94256BA5A6DE471BABFC33F03EAFAE4FAE6A7E1FB7BBD42E06D964919C86F942FDDCEDCE00E71EF4DD1BE16F80347BF4BFB0F0C58A5C467746EFBA4DA7B1018900FBD7698A3157EDEA5ADCCEDEA47D5A95F9B955FD0A5ABD85A6A9A5DCE9B7D089AD6EA268A68F246E461823239E9593E11F06786FC2969716BA0697159C572DBA6C33317E31CB31271ED9EE6BA3E693B75A85392564F42DD38392935AAEA72DE10F01784BC25753DCF87F468ACE6B85092C9B9DD8A839C02C4E07D3AF1E949A3F803C23A4F8967F125868B143AA4ECEF25C191D8EE739620312013EC07535D5E28A6EB4DB6DC9EBE64AA1492494569B69B1CEF8B7C17E17F15089B5FD12D6F9E2188DDC10EA3D3729071ED9C52F84FC1FE1AF0AC72AF87F47B6D3CCB8F31A304B3E3A02C49247B66BA0141A5ED67CBCB776EC3F630E6E6E557EF6D4F93BE36DFDC7C48F8DD65E12D2DB7DBD94A2C9594640727333FD0631FF00AFA9747B0B6D2B4AB5D36CD025BDAC4B0C4BE8AA303F9556D3F40D0EC3509351B1D1EC2DAF25DDE65C456C8923E4E4E580C9C9E6B56BA31189F6B4E14D2B28FE672E1706E8D49D593BCA4FEE5D8E6BC67E06F0BF8C1AD5BC45A62DF1B50C21CCAE9B77633F748CF415ABA0E9561A26950E99A65A476B696E8122893A28FE64FB9E4D681CE38A0573BA9271516F4EC75AA505273495DF5EA3AA29A28E689A29515E3705595864303D41152D1526872575F0EFC21733195B493193FC10DCCD120FA2A3851F80AD0D03C27E1CD02469749D1ED2D667FBD304DD2B7B176CB1FCEB77814952A2A3B21593773375BD134AD6AD85BEA7631DC22FDD272AC9FEEB0C15FC08ACBD23C09E14D27504D46CB47885E46311CD348F3347FEE9763B7F0C574DF8D2FE34F9527707AE8CCED7349B1D6EC1AC75185E5818838491A36041EA194823F0349A068FA7683A5A69BA5C060B546670A646739662CC4B312492493C9AD1E9475A1248673FF00F088681FF091C7E2036721D411CC88FF0069936AB152A488F76C0704F6EE6A5F14786740F14D98B4D7B4AB6BF8D0E504A9CA1F5561C8FC0D6E521A706E16E5D2C4CA31926A4AF7396F0C7807C1DE1A9DE7D1341B5B49E4428D28DCCFB4F501989201F6A4F07F80BC25E12BB9EEBC3DA3456534EA164937BBB1507380589C0FA75E3D2BAAA2ADD5A8EF793D77D4CD50A6AD68AD36D360EA3A5711AAFC29F87BA9DF3DF5DF85EC9A790E5D90B2027D76A9033F857714D3F8528D49D3D62DAF41D4A50A8AD349FA99FA3691A6E8BA7C7A7E91636F65691FDC8A140AA3D4F1DFDEB167F0078467F177FC257368B149AC6F0FF683239F98285076E76E4003B76CF5AEAF0293A50AA4936D37760E941A49C5591CDF8CFC0FE17F1835AB788B4C5BE36A1843995D36EEC67EE919E82B5341D2AC344D2A1D334CB48ED6D2DD0245127451FCC9F73C9AD039C71CD0287524E2A2DE9D86A94149CD2577D7A8EA28A2A4D028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A29A49EC3F13400BEF514D71143FEB6445F4C9E4D298F70F9998FB0381FA52A431C7F71157E82A1F33D86ADD4A52EA3C620B5B99BD0AC640FCCE2A9CF7FAE303E469007A17901FD38ADB19A5AE7A942A4FEDB5E897EA99A46A423F653F5B9C6DFEA5E2A854B35B08C0EA523DC07EA6B1D7C4FAF46D9372AE3D1A35C7E8057A4103D066B97F19E95035B7DB638D5250C03E063703DFEB9AF9FCCF2EC5D3A6EB51AF276D6CDFF969F81EA60F15425354EA535AF5B10E83E2FF00B45CA5BEA30A44CE76AC884EDCFB83D3EB5D7F515E53058BDC5CC70463E791828AF548C6D50BE95BF0F6331389A7255DDED6B333CD70F468CE2E9697E849451457D21E505145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400868C507F95657897503A6E8F35D290240B84CFF78F03FC7F0AC6B558D1A72A93D92BBF9150839C9423BB34C641E3A52D796FFC25DAF7FCFE0FFBF49FE15DF7863513A9E8F0DD31064C6D7C7F78707FC7F1AF2B2ECF70D985474E9269A57D6C76E2B2DAD858A94ED6F235A8A28AF6CE010F5A6E41E322A3BC622D6520E085383F85794695A85FB6A968AD7B72CA66404195882370F7AF1733CE2380A94E128DF9FF00E01DD84C0CB1319493B729EBD4520E8296BD94708514514C04AC3F14CA0DB2DB0E59CE48F61FFD7AD999C2216C64F61EB54ADECC34C6E2E3E6909C81D96B8F17195583A51EBBF9236A1250973BE850F0E691F676FB54EB8908C2A9FE11EBF5ADFE28ED462B4C361E187A6A10D89AB565567CD2168A28AE8330A28A2800A28A2800A28A2800A28A2801323A9A0608E2B96F8AF3CF6FF0CFC49716F2C90CD1E993B47246C55948438208E41AE33F654D42FF0052F8672DC6A37D757938D4654F32E25691B0153032C49C735B4683749D5BECEC733C4255952B6EAE7AED1451589D214514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145003715C1FC4DBEDD3DBE9C8DC20F31C7B9E07E99FCEBB59271F36C2005192C7A015E45ACDE35FEA771764922472573FDD1C0FD00AF93E2BC72A7855462F59BFC16E7B392E1B9EBFB47F67F3263A5B8F0EAEABCE0CDB31FECE3AFE7915D0FC32BFD97371A73370E3CC41EE383FA63F2A7BEB7A09F0CFF6406973E4ED07CAE37F5CFF00DF5CD72BA2DE369FAA5BDD8CE23705B1DD7A11F966BE669CE865D8BA1568CD3565CD677D7667B128D4C5E1EAC2A45A7776BFE07B19200E6BCF7C49E2EBB96E5E0D364F2A053B7781967F7F615D878866DBE1DBC9A16FF960C548F71D6BCCBC3B6F1DD6B96704A03234A3703D081CE2BE8B88F1F5954A585A12E573DDAF3765A9E565385A528CEB5557E5E8685B6A1E291079EBF6D9612A725E32CA47D48FD4564E91FF00215B3FFAEF1FFE842BD7EEC016728000FDD9FE55E41A47FC85ACFF00EBBC7FFA10AF1338CBE782AD4232A8E777D7A6AB63D0C0626388A751A828DBB7CCEFFC7379AA69F6705DD85C98943EC940456EBD0F23D8FE7557C07AEDEEA1753DADFCFE6B840F19DA0700E08E07B8AE8F5AB24D434A9ED0E3F7884293D9BA83F9E2BCBF40BA6D335D8269328124D9283D81E0D7B19AD7AF80CCA955E67ECE5BABBB767A6DB6A7060A953C4E1274F9573AEB6D7BAFF23BFF001B6A93697A5ABDB3F973CAE111B00E3B9383F4C7E3591E08D4F59D5352737376D25BC29961B14649E00E07D4FE159BF11AF85CEB096C8D94B74E7FDE6E4FE98AE9BC0563F63D092565C3DC1F31BE87A7E9CFE34E9E26B63B397184DAA70DD26ECFD7E6C52A54F0F8052925CD2F2D7FAB1C47C5D5F8B97BE2BB6D3FC0534769A5FD8D5A7B871100252EE08DCE0B740BF74572971E12FDA234E85AF6DFC67677D2A0DDE425C6E66F60248827EB5D1FC50F8C92E87E261E11F08E8CDAE6BB908E08631C6E4676055E5CE393C803D7AE32165FDA47544F3162D1F4756190ADE51C7E0779FCEBF48A31A91A71BA8A5E76BBFD4F87AEE9CEA4B96526FCAF65FA1BFF00007E25DF78D62BFD1F5FB58EDF5BD3706528BB04AB9DA495FE1604608E9C8E9D2A0FDA73C65E24F06E89A3DCF87350FB14B7172E92B7931C9B9428207CEA71CFA5703FB32A6A317C6EF1547AB4A92EA2B6F722EDD000AD28B98F79180060B67B0ADEFDB47FE45CF0FF00FD7DC9FF00A053F634D639412567D3A6C4AC454965F2936F99697EBB90C5AC7C71F88B6915CF87A18BC35A498D7CB9E6611C971C7DFCED2D83D46D50BCF53597A4F8EBE257C35F1E58687F10AEBFB434DBD650647659308C76F9892000FCA7AAB76EC320D7D15E1AC7FC23BA671FF2EB17FE802BC07F6D445DBE15971F36EBA5CFB7EEA961EB42B55F62E0945DFA6BF78F15427428FB7539392B75D37EC7D1173710DADB4B733C8B1C3121791D8E02A81924FB015F3CDC7C45F88FF133C4977A5FC348D34DD2AD8ED7BE95543104F0CE581DB9C70AA3775F7C7A47C7EBC9ACBE096B9342D877B68A2247F75E4446FD18D657ECA9636F6BF082CEE625024BCB99E598F72C1CA0FD105614230A546559ABBBD95F6EF73A31139D6AF1A11938AB5DDB7F4391BED2FF00689F0A47FDA316B76BE21893E696DE2226247FBAE8AC7E88735EB0DE32B7D33E185AF8CBC449F66CE9F15CCF0C6A41F35D47EED4139C9638009FAD75F8E31D6BC23F6CBBC962F05E8F6284AC53DF97931DF646703FF1EFD28A7358BA9084A296BBA56D07560F074E752326F4D9BBEA66E8F7DF1A3E29C6FACE91A9DB785B437722D97254C801C643052EDE84FCA339C0A8756F11FC5AF847A8D9DD78B2FA1F1268171288DDD4EE20F5C062A195F00900E54E0FE17FC35E26F8DF63E1ED3ACF4CF877A5B58C16D1C76E778E630A029FF5DDC60D677C4193E36F8D3C2F71A06A7F0FEC61B79991F7C322EF42AC08C66523B63A7426BB6315ED3964A3C9EAAF6F5DEE7039354F9E2E5CFDECED7F4DAC7D0765AA595E68D06B30DC2358CD6EB729293853195DC18FA71CD7CF97DF107E247C4EF13DDE95F0D7FE25BA4DB1C3DDB00ACCB9203BBB025738E154671EBCE3A6D6A1D6BC33FB28DC596A703DAEA50587D9658CB02515E6D98C8247DC6AD0FD942C2DED7E125BDCC48A25BCBA9A595BB921B60FD10572D38428539D6B2959D95F6F53B2AD4A988A90A37714D5DDB47E8705E30D0FE3AF87FC1BAB9D575DB1D6F4892CE55BD4F3448D1C45486605D55B2064F04FD0D769FB217FC92B9BFEC252FF00E811D771F187FE49678A3FEC1771FF00A2CD70FF00B217FC92B9BFEC252FFE811D54EB3AB846DA4B55B2B130C3AA38D8A4DBF75EEEE7B3514515E59EC8514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514005145140051451400514514008719ACC9AE8DCCDF67B73C776F5FFEB557F12EA3F6755B68DB0F20CB1F45FF00EBD2785C078E49BA9DDB47F3AF2EAE2D54AEB0F07EBFE475C287252F6B2F91078D2ED74EF0F4B1A1C4937EED4F739FBC7F2CD70FE17D286B1A9FD99DDD23085D997A81D07EA457A76A3A6D8EA1B45E5BACC133B7713C67AFF2A6E9DA5E9FA7B3B59DB2445C00C57BE2BCCC7E493C6E36152A35ECE3A5B5BFF4D9D785CC6387C3CA114F9DF539F3E03D3C0E2F2EBFF1DFF0AE47C4BA5FF656A66D95D9E328190B75C1FF00EB835EBBC8154350D2B4FBF747BBB5495907CA5BB52CC786B0D5A8F2E1E2A32BEFA8F0B9BD5A752F55B92317C2928D63C24F67237CC91B5B93E831F29FC88FCAB815FB4E95A982CBB27B7901C11DC1FE55EB5A7E9B63601C59DB2C41F1BB693CE3A543AA68BA6EA4435D5B2BB8180C0956FCC5658EC86BE270F4BDF4AA415AFD1F62B0D9952A3567EEBE4974EC72F3F8DFCEB43147A7B79CEBB797CA827F0C9AE4F48FF90AD9FF00D778FF00F4215EA3A7F87F4AB124C166A1C8C176259B1F53D3F0A23F0EE8D1489247611ABA3065233C11D3BD615F23CC7172A752BD44DC7CADDBB2D4D69E6385A2A51A50693FEBB9A6076EC2BCC3C77606CB5D9240B88EE0798BF5FE2FD79FC6BD46A9EA5A6D96A0A82EEDD25D99DBBBB67AD7B79CE57FDA387F669DA49DD3679F97E33EA9579F74F73CA74BB79B55D62185CB3B4AE3CC63D76F527F215EBF1288E354500051818ED54AC346D32CA6F3ADAD238E4C6370EB8AD0C563916512CBA9CBDA34E4DEEBB1A6638E58B92E55648F95B44D62CBE1EFED2DE20BBF162BC36F7725CF9572632C23599C489271C91B46DC8E993E86BD875FF8D9F0EF4BD3DAE23D71350971F25BD9A33BB9F4E4003F122BA4F19F81FC2DE30B78D3C43A3C378D10F924C94913D83A9071ED9C5647863E137807C377AB7DA6F87A137487292DC48F3143D8A872403EE066BEC2788C3D64A534F992B696B3B1F2F4F0D89A2E50A6D72B6DDDDEEAFF0099E3BFB31DFBEA9F1BBC53A9C96EF6AF776F733B40E7E688BDCC6DB4F03919C74ADDFDB473FF0008DF87F3FF003F72FF00E802BD8346F08786F46D72EF5CD334982D751BCDFF00699D376E9373066CE4E39600D64F89CF80BC597F168BAFC115FCB6F74D0C493432AC627C728B260296C0E99A3EB519625558C5D92FD05F539470AE84A4AEDEFF003B9D3786FF00E45DD33FEBD62FFD0057827EDADFEA3C2BFF005D2EBF9455F43DBC31DBC11C312858E35088A3B00300573FE30F0D7863C4D2D843E23D2E3BF68E4716C1D1C852572DCAF03213BFA63AD73E1AB2A559546B4573AB1741D6C3FB24ECDDBF34278EBC3E3C51E00D4740DCA8F7767B2366E8B2000A13EC180AF0CF801F116CFC0B1DF7813C6E24D2A4B6B966865950958D8E3746D8E8323706E87279E95F4B2A850028C003000ED5CE78BBC0DE12F15ED6D7F43B4BD9546D594829201E9BD486C7B66AE86220A32A7555E2F5D374C8C461A6E71AB49A525A6BB34739E25F8D7F0FB45B169D35C8B539F194B7B1064673E99FBABF89AC8F8C5A1DDFC4DF82D61AAE9966CB7E22875482D95B733068FE68C1E3276B1C71C951EB5D068FF073E1BE9776B736DE18B79245E57ED32C93A8FF0080BB11FA574DA9EBDA3E88F0D9DD4AE9218F72416F6D24CCB18E37158D49551D32401473D2A738BA09B69DF5FF0080274AAD584A3886926ADA7E7767907C15F8D1E1E1E19B3D03C5979FD95A969E82D84B3A37972AA0DAA4903E5600004363919EF81D9788FE35FC3BD1AD0CABAE26A52E32B0D8A9959BF1E147E24568F883E1EF807C605354D4F41B5B992E14482E622D13C808C824A104F1EB599E15F857F0B622D7DA5F872DEE8C5349096BA69260248DCA38DB21238652338EDC56929E126DCDA92F2D2DF799C218C845414A2FB377BDBD0BFAB463E247C1C9BC980DBBEB3A6F99044EF9D8E46E404FFBC1735E41FB3AFC4BD23C21A55EF83BC5F349A54B6D74EF03CB1B6D527878DB009521813CF1C9F4E7E8837BA759C5716EAE882C2157961850B189083B70AA33FC27000ED5CE78C7E18F81FC5979F6DD6B438A4BB200371148D13B63FBC548DDF8E6A28D7A6A12A5553E56EEADBA2EB61EACA71AB4A4B992B3BECFFA679F7C62F8CBE10BAF076A9E1FD067975ABDBEB492026DE3611C2AC872ECCC39C0C9C0CF4E48AD1FD90BFE495CC3FEA252FF00E811D763A07C34F03E85A75D5869BE1FB648AF22686E4BB33C92230C152EC4B007D0102B73C31E1FD1BC35A69D3B42B18AC6D4C86431464E3710013C93E829D4AF455174A9A7BDEEC54B0F5DD7556AB5B35646C514515C27A61451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450014514500145145001451450079EF8C26923F10CBBC90B8529F4DA3FAE6ACF8675A8AD2478AE3222739DC0E769AE975AD1ECB558C25CA1DCBF76453865FC6B04F82F9CC5A8B28F468B27F98AF8FAF97E6187C63AF412926DBDD75E8EF63DCA58BC355C3AA5574B2B7F563AB82E6DE74DF0CD1C8BEAAC0D4C08AE561F08B2105B5263FEEC78FEB5A36FE1F863C799757327B17C0AF7686231925FBCA567EABFE09E6D4A5417C33BFC99AB24D14632F2A27FBCC0546B728FF00EAB327FBA38FCFA5476FA75A41CA4084FA9E4FEB57315DB1F6AFE2B2FC7FC8C1F2ADB51A379EA028FCCD3B81EF4B456A91170A28A298051451400514514005145140086BCAFF00E11DD73FB7E4921D37544986BE6F62B892F22364212E371316F2C58C65C0C26431072319AF55149570A8E1732A94954B5FA1E5D0E8DE209ADA0D32E749D5228ACED751825B986E615694CA7F766125F3B88E858280719EF49A2F87F5D8A1B48D34686CA086FA77511C515BBBA359C881A48E3764CF98C17208CF0481D6BD4BBD2569F5895AD6465F568DD3BB3CA66F056AF6BA30B7D1EC7EC924FA45AC77A2291035C4D1CA8D2231270CED1F98BB9B83BB04E2A2BAF0A5D7F63AA5B689AB3817924D1595CDB58BDBA93122F300955029209054821B71E3764FAE0A5A1622481E12079B693A26B1078DECF516D11433F966EE6992278ED80B60845B4A24F342EE006C6520E58F19CD6978DAC2E65D662BDB0D375F1782DBCA4BDD2AE204FE22447224CC14807904A91C9E9DFB6A4150EB36D3B6DA14A84545C6FBBB9E5B7BA0789AE2712EA7A58BDD6A58EC8DBEA70BC623B1740BE7E32C1946E0EDF2A9DE1803E82DC3A678834BBDB7BC4D1AE6F5776B1118A09A20CBF69BB496273B9D46D2A873C923238AF48A2ABDBB7D10961A29DEEFFA773C88784353834BD423FF008478CBA9DE786A0B58AF15A1CC570903C6E8CC583027E4191907D78AD8FEC1D53FE12CFB47F64C9F6CFED7FB4FF6BF989B7EC9B71E4FDEDFD3E4D9B76E7E6F7AF45FC68A1D793DC161A2B6679458F826F2CF46D3E24D36E209A5F0F4F6BAA9B49A313C9705ADCA0259B6BB0026C12718C8C804575BF0E6C6FB4ED1EE20BDD32DF4F5FB531B78E1B7480BC7B53E778D1D915B76E1F29E400700922BABA294AB392B32A1423097320A28A2B2370A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2900514514C028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A0028A28A40145145300A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2800A28A2803FFD9

	IF EXISTS(SELECT * FROM ASRSysPictures p
		INNER JOIN ASRSysSystemSettings s ON s.SettingValue = p.PictureID
		WHERE s.Section = 'desktopsetting' AND s.SettingKey = 'bitmapid'
			AND p.[name] IN ('Advanced Business Solutions Wallpaper 1024x768.jpg', 'Advanced Business Solutions Wallpaper 1280 x800.jpg',
			'Advanced Business Solutions Wallpaper 1440x900.jpg', 'Advanced Business Solutions Wallpaper 2560x1600.jpg',
			'ASRDesktopImagePersonnelnPost.bmp', 'ASR Splash.jpg', 'ASRDesktopImage 1024x768.jpg',
			'COASolutionsDesktopImage-1024x768.jpg', 'HRProP.bmp', 'HRProPP.bmp', 'HRProPR.bmp', 'HRProPRP.bmp',
			'HRProPRT.bmp',	'HRProPRTP.bmp', 'HRProPRTS.bmp', 'HRProPS.bmp', 'HRProPT.bmp', 'HRProPTP.bmp', 'HRProPTS.bmp',	'HRProT.bmp'))
	BEGIN
		-- Set backcolour to white, image to our newly inserted one and tile in the centre
		EXEC spsys_setsystemsetting 'desktopsetting', 'backgroundcolour', '16777215';
		EXEC spsys_setsystemsetting 'desktopsetting', 'bitmapid', @newDesktopImageID;	
		EXEC spsys_setsystemsetting 'desktopsetting', 'bitmaplocation', 2;	
	END


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '5.1';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v5.1')


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
EXEC dbo.spsys_setsystemsetting 'database', 'refreshstoredprocedures', 1;


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF;

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v5.1 Of OpenHR'
