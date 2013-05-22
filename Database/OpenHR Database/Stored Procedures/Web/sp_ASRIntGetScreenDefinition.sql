CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenDefinition] (
	@piScreenID 		integer,
	@piViewID			integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the given screen's definition and table permission info. */
	DECLARE @iTabCount 		integer,
		@sTabCaptions		varchar(MAX),
		@sTabCaption		varchar(MAX),
		@fSysSecMgr			bit,
		@fInsertGranted		bit,
		@fDeleteGranted		bit,
		@sRealSource		sysname,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@iTableID			integer,
		@iTableType			integer,
		@sTableName			sysname,
		@iTempAction		integer,
		@iChildViewID 		integer,
		@sActualUserName	varchar(250);

	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT
					
	/* Get the table type and name. */
	SELECT @iTableID = ASRSysScreens.tableID,
		@iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysScreens
	INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.screenID = @piScreenID

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
	OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'

	/* Get the real source and insert/delete permissions for the table. */
	IF @fSysSecMgr = 1 
	BEGIN
		/* Permission must be granted for System or Security mangers. */
		SET @fInsertGranted = 1
		SET @fDeleteGranted = 1	

		/* Get the realSource of the table. */
		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			IF @piViewID > 0 
			BEGIN
				/* RealSource is the view. */	
				SELECT @sRealSource = viewName
				FROM ASRSysViews
				WHERE viewID = @piViewID	
			END
			ELSE
			BEGIN
				/* RealSource is the table. */	
				SET @sRealSource = @sTableName
			END 
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sRealSource = left(@sRealSource, 255)
			END
		END
	END
	ELSE
	BEGIN

		/* Permission must be read from the database  for Non-System and Non-Security mangers. */
		SET @fInsertGranted = 0
		SET @fDeleteGranted = 0	

		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			IF @piViewID > 0 
			BEGIN	
				/* RealSource is the view. */	
				SELECT @sRealSource = viewName
				FROM ASRSysViews
				WHERE viewID = @piViewID
			END
			ELSE
			BEGIN
				SET @sRealSource = @sTableName
			END 

			/* Get the insert/delete permissions for the realSource. */
			DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT p.action
				FROM #SysProtects p
				INNER JOIN sysobjects ON p.id = sysobjects.id
				WHERE p.action  IN (195, 196)
					AND sysobjects.name = @sRealSource
					AND ProtectType <> 206

			OPEN tableInfo_cursor
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
			WHILE (@@fetch_status = 0)
			BEGIN
				IF @iTempAction = 195
				BEGIN
					SET @fInsertGranted = 1
				END
				ELSE
				BEGIN
					SET @fDeleteGranted = 1	
				END
				FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
			END
			CLOSE tableInfo_cursor
			DEALLOCATE tableInfo_cursor
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sRealSource = left(@sRealSource, 255)

				/* Get appropriate child view if required. */
				DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT p.action
					FROM #SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					WHERE sysobjects.name = @sRealSource
						AND p.Action IN(193, 195, 196)
						AND ProtectType <> 206

				OPEN tableInfo_cursor
				FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iTempAction = 195
					BEGIN
						SET @fInsertGranted = 1
					END
					ELSE
					BEGIN
						IF @iTempAction = 196
						BEGIN
							SET @fDeleteGranted = 1	
						END
					END
					FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
				END
				CLOSE tableInfo_cursor
				DEALLOCATE tableInfo_cursor
			END
		END
	END
	
	/* Get the tab page captions info. */
	SET @iTabCount = 0
	SET @sTabCaptions = ''
	
	DECLARE captions_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT caption 
		FROM ASRSysPageCaptions
		WHERE screenID = @piScreenID
		ORDER BY pageIndexID

	OPEN captions_cursor
	FETCH NEXT FROM captions_cursor INTO @sTabCaption
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTabCount > 0 SET @sTabCaptions = @sTabCaptions + char(9) 

		SET @iTabCount = @iTabCount + 1
		SET @sTabCaptions = @sTabCaptions + @sTabCaption
			
		FETCH NEXT FROM captions_cursor INTO @sTabCaption
	END
	CLOSE captions_cursor
	DEALLOCATE captions_cursor

	SELECT @sTableName AS tableName,
		@sRealSource AS realSource,
		@fInsertGranted AS insertGranted,
		@fDeleteGranted AS deleteGranted,
		height,
		width,
		fontName,
		fontSize,
		fontBold,
		fontItalic,
		fontStrikethru,
		fontUnderline,
		@iTabCount AS tabCount,
		@sTabCaptions AS tabCaptions
	FROM ASRSysScreens
	WHERE screenID = @piScreenID
	
END