
/* --------------------------------------------------- */
/* Update the database from version 4.0 to version 4.1 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion int,
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16)

DECLARE @sSQL varchar(max)
DECLARE @sSPCode nvarchar(max)
DECLARE @sSPCode_0 nvarchar(4000)
DECLARE @sSPCode_1 nvarchar(4000)
DECLARE @sSPCode_2 nvarchar(4000)
DECLARE @sSPCode_3 nvarchar(4000)
DECLARE @sSPCode_4 nvarchar(4000)
DECLARE @sSPCode_5 nvarchar(4000)
DECLARE @sSPCode_6 nvarchar(4000)
DECLARE @sSPCode_7 nvarchar(4000)
DECLARE @sSPCode_8 nvarchar(4000)

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
IF (@sDBVersion <> '4.0') and (@sDBVersion <> '4.1')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2005 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion <> 9 AND @iSQLVersion <> 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of HR Pro', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 - Modifying Workflow procedures'

	----------------------------------------------------------------------
	-- spASRActionActiveWorkflowSteps
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRActionActiveWorkflowSteps]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRActionActiveWorkflowSteps];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			-- Return a recordset of the workflow steps that need to be actioned by the Workflow service.
			-- Action any that can be actioned immediately. 
			DECLARE
				@iAction			integer, -- 0 = do nothing, 1 = submit step, 2 = change status to ''2'', 3 = Summing Junction check, 4 = Or check
				@iElementType		integer,
				@iInstanceID		integer,
				@iElementID			integer,
				@iStepID			integer,
				@iCount				integer,
				@sStatus			bit,
				@sMessage			varchar(MAX),
				@iTemp				integer, 
				@iTemp2				integer, 
				@iTemp3				integer,
				@sForms 			varchar(MAX), 
				@iType				integer,
				@iDecisionFlow		integer,
				@fInvalidElements	bit, 
				@fValidElements		bit, 
				@iPrecedingElementID	integer, 
				@iPrecedingElementType	integer, 
				@iPrecedingElementStatus	integer, 
				@iPrecedingElementFlow	integer, 
				@fSaveForLater			bit;
		
			DECLARE stepsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT E.type,
				S.instanceID,
				E.ID,
				S.ID
			FROM ASRSysWorkflowInstanceSteps S
			INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
			WHERE S.status = 1
				AND E.type <> 5; -- 5 = StoredData elements handled in the service
		
			OPEN stepsCursor;
			FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @iAction = 
					CASE
						WHEN @iElementType = 1 THEN 1	-- Terminator
						WHEN @iElementType = 2 THEN 2	-- Web form (action required from user)
						WHEN @iElementType = 3 THEN 1	-- Email
						WHEN @iElementType = 4 THEN 1	-- Decision
						WHEN @iElementType = 6 THEN 3	-- Summing Junction
						WHEN @iElementType = 7 THEN 4	-- Or	
						WHEN @iElementType = 8 THEN 1	-- Connector 1
						WHEN @iElementType = 9 THEN 1	-- Connector 2
						ELSE 0					-- Unknown
					END;
				
				IF @iAction = 3 -- Summing Junction check
				BEGIN
					-- Check if all preceding steps have completed before submitting this step.
					SET @fInvalidElements = 0;	
				
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID;
		
					OPEN precedingElementsCursor;			
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
		
					WHILE (@@fetch_status = 0)
						AND (@fInvalidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*) 
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID;
		
								IF @iCount = 0 SET @fInvalidElements = 1;
							END
							ELSE
							BEGIN
								SET @fInvalidElements = 1;
							END
						END
						ELSE
						BEGIN
							IF (@iPrecedingElementType = 2) -- WebForm
							BEGIN
								IF @iPrecedingElementStatus = 3 -- 3 = completed
									OR @iPrecedingElementStatus = 6 -- 6 = timeout
								BEGIN
									SET @iTemp3 = CASE
											WHEN @iPrecedingElementStatus = 3 THEN 0
											ELSE 1
										END;
		
									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID;
								
									IF @iCount = 0 SET @fInvalidElements = 1;
								END
								ELSE
								BEGIN
									SET @fInvalidElements = 1;
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingElementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END;
		
										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID;
									
										IF @iCount = 0 SET @fInvalidElements = 1;
									END
									ELSE
									BEGIN
										SET @fInvalidElements = 1;
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus <> 3 SET @fInvalidElements = 1;
								END
							END
						END
		
						FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
					END
					CLOSE precedingElementsCursor;
					DEALLOCATE precedingElementsCursor;
					
					IF (@fInvalidElements = 0) SET @iAction = 1;
				END
		
				IF @iAction = 4 -- Or check
				BEGIN
					SET @fValidElements = 0;
					-- Check if any preceding steps have completed before submitting this step. 
		
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID;
		
					OPEN precedingElementsCursor;	
		
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
					WHILE (@@fetch_status = 0)
						AND (@fValidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID;
							
								IF @iCount > 0 SET @fValidElements = 1;
							END
						END
						ELSE
						BEGIN
							IF (@iPrecedingElementType = 2) -- WebForm
							BEGIN
								IF @iPrecedingElementStatus = 3 -- 3 = completed
									OR @iPrecedingElementStatus = 6 -- 6 = timeout
								BEGIN
									SET @iTemp3 = CASE
											WHEN @iPrecedingElementStatus = 3 THEN 0
											ELSE 1
										END;
		
									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID;
							
									IF @iCount > 0 SET @fValidElements = 1;
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingElementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END;
		
										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID;
		
										IF @iCount > 0 SET @fValidElements = 1;
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus = 3 SET @fValidElements = 1;
								END
							END
						END
		
						FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
					END
					CLOSE precedingElementsCursor;
					DEALLOCATE precedingElementsCursor;
		
					-- If all preceding steps have been completed submit the Or step.
					IF @fValidElements > 0 
					BEGIN
						-- Cancel any preceding steps that are not completed as they are no longer required.
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @iInstanceID, @iElementID;
		
						SET @iAction = 1;
					END
				END
		
				IF @iAction = 1
				BEGIN
					EXEC [dbo].[spASRSubmitWorkflowStep] @iInstanceID, @iElementID, '''', @sForms OUTPUT, @fSaveForLater OUTPUT;
				END
		
				IF @iAction = 2
				BEGIN
					UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
					SET status = 2
					WHERE id = @iStepID;
				END
		
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID;
			END
		
			CLOSE stepsCursor;
			DEALLOCATE stepsCursor;
		
			DECLARE timeoutCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT 
				WIS.instanceID,
				WE.ID,
				WIS.ID
			FROM ASRSysWorkflowInstanceSteps WIS
			INNER JOIN ASRSysWorkflowElements WE ON WIS.elementID = WE.ID
				AND WE.type = 2 -- WebForm
			WHERE ((WIS.status = 2) OR (WIS.status = 7)) -- Pending user action/completion
				AND isnull(WE.timeoutFrequency,0) > 0
				AND CASE 
						WHEN WE.timeoutPeriod = 0 THEN 
							dateadd(minute, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 1 THEN 
							dateadd(hour, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 2 AND WE.timeoutExcludeWeekend = 1 THEN 
							dbo.udfASRAddWeekdays(WIS.activationDateTime, WE.timeoutFrequency)
						WHEN WE.timeoutPeriod = 2 THEN 
							dateadd(day, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 3 THEN 
							dateadd(week, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 4 THEN 
							dateadd(month, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 5 THEN 
							dateadd(year, WE.timeoutFrequency, WIS.activationDateTime)
						ELSE getDate()
					END <= getDate();	
		
			OPEN timeoutCursor;
			FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID;
			WHILE (@@fetch_status = 0)
			BEGIN
				-- Set the step status to be Timeout
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 6, -- Timeout
					ASRSysWorkflowInstanceSteps.timeoutCount = isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID;
		
				-- Activate the succeeding elements on the Timeout flow
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT id 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1))
					AND (ASRSysWorkflowInstanceSteps.status = 0
						OR ASRSysWorkflowInstanceSteps.status = 3
						OR ASRSysWorkflowInstanceSteps.status = 4
						OR ASRSysWorkflowInstanceSteps.status = 6
						OR ASRSysWorkflowInstanceSteps.status = 8);
					
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
						AND ASRSysWorkflowElements.type = 1);
					
				-- Count how many terminators have completed. ie. if the workflow has completed.
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1;
										
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @iInstanceID;
					
					-- NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					-- is performed by a DELETE trigger on the ASRSysWorkflowInstances table.
				END
		
				FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID;
			END
		
			CLOSE timeoutCursor;
			DEALLOCATE timeoutCursor;
		END';

	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
PRINT 'Step 2 - Version 1 Integration Modifications'


	-- Create document management map table
	IF OBJECT_ID('ASRSysDocumentMapping', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysDocumentMapping]
                    ( [DocumentMapID]			integer			NOT NULL IDENTITY(1,1)
                    , [Name]					nvarchar(255)
                    , [Description]				nvarchar(MAX)
                    , [Access]					varchar(2)
                    , [Username]				varchar(50)
                    , [CategoryRecordID]		integer
                    , [TypeRecordID]			integer                    
                    , [TargetTableID]			integer
                    , [TargetKeyFieldColumnID]	integer
                    , [TargetColumnID]			integer
                    , [ParentTableID]			integer
                    , [ParentKeyFieldColumnID]	integer
                    , [ManualHeader]			bit
                    , [HeaderText]				nvarchar(MAX))
               ON [PRIMARY]'
	END	

	-- Add columns to ASRSysMailMergeName
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'OutputPrinterName')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName
								 ADD [OutputPrinterName] nvarchar(255), [DocumentMapID] integer';
	END


	-- Add columns to ASRSysControls
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysControls', 'U') AND name = 'NavigateTo')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysControls
								ADD [NavigateTo] nvarchar(MAX), [NavigateIn] tinyint, [NavigateOnSave] bit';
	END	

	-- Insert the system permissions for Document Management
	IF NOT EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 43)
	BEGIN
		INSERT dbo.[ASRSysPermissionCategories] ([CategoryID], [Description], [ListOrder], [CategoryKey], [Picture])
			VALUES (43, 'Document Types', 10, 'VERSION1',0x00);
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (154,43,'New', 10, 'NEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (155,43,'Edit', 20, 'EDIT');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (156,43,'View', 30, 'VIEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (157,43,'Delete', 40, 'DELETE');
	END

	-- Update the system permission image for Document Management
	IF EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 43)
	BEGIN
		SELECT @ptrval = TEXTPTR([picture]) 
		FROM dbo.[ASRSysPermissionCategories]
		WHERE categoryID = 43;

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000000000000000000000000000000000000000000000000070306000F060D000E0D0E00100F0F001411130024162000281523002A1B26003B162D003D172F002C2C2C0032282E00302F2F00372131003B3B3A00401F3200491A360061224E007329540072396100404040004F4F4E005648510051505000575655005E5E5E007C496F00606060006766650077767600787776007E737A0081376D008F467A008D507C00964F82009A5586009F5E8D0090648300A8759900AD759D00AB7F9900448BF000468CF000488DF0004F92F0005093F1005295F1005798F2005E9CF30070A7F30071A9F50074ABF50076ACF500918F8E00979796009D9C9A009E9E9E00A9829C00A09F9D00B684A700B68DA500B786A900B887A900BB8CAD00B799AD00A2A2A100A8A6A500A8A8A800B3B2B100C195B300C297B400C399B600C59BB700C7A2BB00C8A1BC00CAA5BE00B6B3D90080AFF20080B1F5008BB9F70092BEF800D5B5C900D4B8CA00D7B9CC00B0D1FA00D1D0CF00D4D3D300D7D6D500D9D8D800DFDEDE00E1D2DC00E0DFDE00E5E4E400EFE1E900EBEAEA00ECEBEB00F5F5F500FBFAFA00FDFDFD004CB0000059CF000067F0000078FF11008AFF31009CFF5100AEFF7100C0FF9100D2FFB100E4FFD100FFFFFF0000000000262F0000405000005A700000749000008EB00000A9CF0000C2F00000D1FF1100D8FF3100DEFF5100E3FF7100E9FF9100EFFFB100F6FFD100FFFFFF00000000002F26000050410000705B000090740000B08E0000CFA90000F0C30000FFD21100FFD83100FFDD5100FFE47100FFEA9100FFF0B100FFF6D100FFFFFF00000000002F1400005022000070300000903E0000B04D0000CF5B0000F0690000FF791100FF8A3100FF9D5100FFAF7100FFC19100FFD2B100FFE5D100FFFFFF00000000002F030000500400007006000090090000B00A0000CF0C0000F00E0000FF201200FF3E3100FF5C5100FF7A7100FF979100FFB6B100FFD4D100FFFFFF00000000002F000E00500017007000210090002B00B0003600CF004000F0004900FF115A00FF317000FF518600FF719C00FF91B200FFB1C800FFD1DF00FFFFFF00000000002F0020005000360070004C0090006200B0007800CF008E00F000A400FF11B300FF31BE00FF51C700FF71D100FF91DC00FFB1E500FFD1F000FFFFFF00000000002C002F004B0050006900700087009000A500B000C400CF00E100F000F011FF00F231FF00F451FF00F671FF00F791FF00F9B1FF00FBD1FF00FFFFFF00000000001B002F002D0050003F007000520090006300B0007600CF008800F0009911FF00A631FF00B451FF00C271FF00CF91FF00DCB1FF00EBD1FF00FFFFFF000000000008002F000E005000150070001B0090002100B0002600CF002C00F0003E11FF005831FF007151FF008C71FF00A691FF00BFB1FF00DAD1FF00FFFFFF0000001800040400083B4A260000000000001C0F3944371F0D0C4A26000000000000154360615A3C1F004D400000000000001545646462571D022122000000000000001A6F6F635D160E214B4F332D50000000013A5A5846051B2800312F2D360000000A185B591907230000302D2B51000000091E64633806545F002B2B2B52000000110B5A5E15204C41554E0000000000001310030017003F254953000000000000132A12140000292448530056000000003E42275C0000003D47530035353400000000000000000000000000322C2B000000000000000000000000002B2B2B000000000000000000000000002E2B2B00000000000000000000000000000000C01F0000801F0000801F0000801F0000C0010000C0210000C0610000C0210000C00F0000C10F0000C30B0000C3880000FFF80000FFF80000FFF80000FFFF000000
	END





/* ------------------------------------------------------------- */
/* ------------------------------------------------------------- */

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
     INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
    OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
    OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
    AND (sysusers.name = 'dbo')
--IF (@@ERROR <> 0) goto QuitWithRollback

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
    IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
    BEGIN
        SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
        --IF (@@ERROR <> 0) goto QuitWithRollback
    END
    ELSE
    BEGIN
        SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
        --IF (@@ERROR <> 0) goto QuitWithRollback
    END

    FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects

/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '4.1')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'ssintranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('ssintranet', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'outlook service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('outlook service', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'workflow service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('workflow service', 'minimum version', '4.1.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v4.1')


SELECT @NVarCommand = 
	'IF EXISTS (SELECT * FROM dbo.sysobjects
			WHERE id = object_id(N''[dbo].[sp_ASRLockCheck]'')
			AND OBJECTPROPERTY(id, N''IsProcedure'') = 1)
		GRANT EXECUTE ON sp_ASRLockCheck TO public'
EXEC sp_executesql @NVarCommand


SELECT @NVarCommand = 'USE master
GRANT EXECUTE ON sp_OACreate TO public
GRANT EXECUTE ON sp_OADestroy TO public
GRANT EXECUTE ON sp_OAGetErrorInfo TO public
GRANT EXECUTE ON sp_OAGetProperty TO public
GRANT EXECUTE ON sp_OAMethod TO public
GRANT EXECUTE ON sp_OASetProperty TO public
GRANT EXECUTE ON sp_OAStop TO public
GRANT EXECUTE ON xp_StartMail TO public
GRANT EXECUTE ON xp_SendMail TO public
GRANT EXECUTE ON xp_LoginConfig TO public
GRANT EXECUTE ON xp_EnumGroups TO public'
--EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE ['+@DBName + ']
GRANT VIEW DEFINITION TO public'
EXEC sp_executesql @NVarCommand


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.1 Of HR Pro'
