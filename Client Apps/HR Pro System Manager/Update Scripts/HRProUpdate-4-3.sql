
/* --------------------------------------------------- */
/* Update the database from version 4.2 to version 4.3 */
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

DECLARE @ownerGUID uniqueidentifier
DECLARE @nextid integer
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
IF (@sDBVersion <> '4.2') and (@sDBVersion <> '4.3')
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
PRINT 'Step 1 - System Functions'

	SELECT @ownerGUID = [SettingValue] FROM asrsyssystemsettings
		WHERE [Section] = 'database' AND [SettingKey] = 'ownerid'

	IF @ownerGUID IS NULL
	BEGIN
		SET @ownerGUID = NEWID();
		INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES ('database', 'ownerid', @ownerGUID)
	END			

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_getmodulesetting]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getmodulesetting];

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[udfsys_getownerid]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getownerid]

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spsys_setsystemsetting]')	AND xtype = 'P')
		DROP PROCEDURE [dbo].[spsys_setsystemsetting];

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getownerid]()
		RETURNS uniqueidentifier
		AS
		BEGIN
			DECLARE @returnval uniqueidentifier;
			SELECT @returnval = [SettingValue]
				FROM dbo.[ASRSysSystemSettings]
				WHERE [Section] = ''database'' AND [SettingKey] = ''ownerid''
			RETURN @returnval
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getmodulesetting](
			@module AS nvarchar(255),
			@modulekey AS nvarchar(255))
		RETURNS nvarchar(255)
		WITH SCHEMABINDING
		AS
		BEGIN
			DECLARE @result nvarchar(255);
			
			SELECT @result = [ParameterValue] FROM dbo.[asrsysmodulesetup] WHERE [ModuleKey] = @module AND [parameterkey] = @modulekey;

			RETURN @result;			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spsys_setsystemsetting](
			@section AS nvarchar(255),
			@settingkey AS nvarchar(255),
			@settingvalue AS nvarchar(MAX))
		AS
		BEGIN
			IF EXISTS(SELECT [SettingValue] FROM [asrsyssystemsettings] WHERE [Section] = @section AND [SettingKey] = @settingkey)
				UPDATE ASRSysSystemSettings SET [SettingValue] = @settingvalue WHERE [Section] = @section AND [SettingKey] = @settingkey;
			ELSE
				INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES (@section, @settingkey, @settingvalue);	
		END';
	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
PRINT 'Step 2 - Scripted Updates Date Effective Module'

	-- Date effective table
	IF OBJECT_ID('tbstat_effectivedates', N'U') IS NULL	
	BEGIN
		EXECUTE sp_executeSQL N'CREATE TABLE tbstat_effectivedates ([type] tinyint, [date] datetime);'
		EXECUTE sp_executeSQL N'INSERT tbstat_effectivedates ([type], [date]) VALUES (1, DATEADD(D, 0, DATEDIFF(D, 0, GETDATE())))'
	END

/* ------------------------------------------------------------- */
PRINT 'Step 3 - New admin system security'

	IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmins' AND type = 'R')
	BEGIN
		SELECT @NVarCommand = 'CREATE ROLE [ASRSysAdmins] AUTHORIZATION [dbo];'
		EXECUTE sp_executesql @NVarCommand;
		
		INSERT @admingroups
			SELECT DISTINCT gp.groupName --, o.name
				FROM ASRSysGroupPermissions gp
				INNER JOIN ASRSysPermissionItems pi ON pi.itemID = gp.itemID
				WHERE pi.itemID IN (1, 3) AND gp.permitted = 1;
		
		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand + 'EXEC sp_addrolemember ''ASRsysAdmins'', ''' + groupName + ''';' + CHAR(13)
			FROM @admingroups;

		EXECUTE sp_executesql @NVarCommand;

	END

/* ------------------------------------------------------------- */
PRINT 'Step 4 - Create object tracking system'

	IF OBJECT_ID('tbsys_scriptedobjects', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_scriptedobjects](
		[guid] [uniqueidentifier] NOT NULL,
		[parentguid] [uniqueidentifier] NULL,
		[objecttype] [int] NOT NULL,
		[targetid] [int] NULL,
		[ownerid] [uniqueidentifier] NOT NULL,
		[effectivedate] [datetime] NULL,
		[disabledate] datetime NULL,
		[revision] [int] NOT NULL,
		[lastupdated] [datetime],
		[lastupdatedby] nvarchar(255),
		[locked] [bit] NOT NULL,
		[tag] [xml] NULL)';
		
		-- Insert table defintions into script base table
	    SET @NVarCommand = 'INSERT [tbsys_scriptedobjects] ([guid],[objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
								SELECT NEWID(), 1, tableID, ''' + convert(nvarchar(64),@ownerGUID) + ''', ''01/01/1900'', 1, 0, [lastUpdated] FROM dbo.[ASRSysTables];';
		EXECUTE sp_executesql @NVarCommand;

		-- Insert column defintions into script base table
		SET @NVarCommand = 'INSERT [tbsys_scriptedobjects] ([guid], [parentguid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked])
								SELECT NEWID(), o.[guid], 2, c.columnID,''' + convert(nvarchar(64),@ownerGUID) + ''', ''01/01/1900'', 1, 0 FROM dbo.[ASRSysColumns] c
									INNER JOIN ASRSysTables t ON t.TableID = c.tableID
									INNER JOIN tbsys_scriptedobjects o ON t.tableid = o.targetid AND o.objecttype = 1;';
		EXECUTE sp_executesql @NVarCommand;

		-- Insert view defintions into script base table
	    SET @NVarCommand = 'INSERT [tbsys_scriptedobjects] ([guid], [parentguid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked])
								SELECT NEWID(), o.[guid], 3, v.viewID, ''' + convert(nvarchar(64),@ownerGUID) + ''', ''01/01/1900'', 1, 0  FROM dbo.[ASRSysViews] v
									INNER JOIN ASRSysTables t ON t.TableID = v.viewtableID
									INNER JOIN tbsys_scriptedobjects o ON t.tableid = o.targetid AND o.objecttype = 1;';
		EXECUTE sp_executesql @NVarCommand;

		-- Insert workflow defintions into script base table
	    SET @NVarCommand = 'INSERT [tbsys_scriptedobjects] ([guid],[parentguid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked])
								SELECT NEWID(), o.[guid], 10, w.ID, ''' + convert(nvarchar(64),@ownerGUID) + ''', ''01/01/1900'', 1, 0  FROM dbo.[ASRSysWorkflows] w
									INNER JOIN ASRSysTables t ON t.TableID = w.basetable
									INNER JOIN tbsys_scriptedobjects o ON t.tableid = o.targetid AND o.objecttype = 1
								UNION SELECT NEWID(), NULL, 10, w.ID, ''' + convert(nvarchar(64),@ownerGUID) + ''', ''01/01/1900'', 1, 0  FROM dbo.[ASRSysWorkflows] w
									WHERE w.basetable = 0';
		EXECUTE sp_executesql @NVarCommand;
		
	END

	-- Object modelling table
	IF OBJECT_ID('tbsys_systemobjects', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE tbsys_systemobjects ([objecttype] integer
				, [tablename] nvarchar(255), [viewname] nvarchar(255), [description] nvarchar(MAX)
				, [nextid] integer, [allowselect] bit, [allowupdate] bit)';

		-- Add table definition
		SELECT @nextid = MAX([tableid]) + 1 FROM dbo.[ASRSysTables];
		SET @NVarCommand = 'INSERT tbsys_systemobjects ([objecttype], [tablename], [viewname], [description], [nextid], [allowselect], [allowupdate])
			VALUES (1,''tbsys_tables'',''ASRSysTables'', ''Table definitions '' ,'  + convert(nvarchar(255),@nextid) + ', 1, 0);';
		EXEC sp_executesql @NVarCommand;

		-- Add column definition
		SELECT @nextid = MAX([columnid]) + 1 FROM dbo.[ASRSysColumns];
		SET @NVarCommand = 'INSERT tbsys_systemobjects ([objecttype], [tablename], [viewname], [description], [nextid], [allowselect], [allowupdate])
			VALUES (2,''tbsys_columns'',''ASRSysColumns'', ''Column definitions '' ,'  + convert(nvarchar(255),@nextid) + ', 1, 0);';
		EXEC sp_executesql @NVarCommand;

		-- Add screen definition
		--SELECT @nextid = MAX([screenid]) + 1 FROM dbo.[ASRSysScreens];
		--SET @NVarCommand = 'INSERT tbsys_systemobjects ([objecttype], [tablename], [viewname], [description], [nextid], [allowselect], [allowupdate])
		--	VALUES (14,''tbsys_screens'',''ASRSysScreens'', ''Screen definitions '' ,'  + convert(nvarchar(255),@nextid) + ', 1, 0);';
		--EXEC sp_executesql @NVarCommand;

		-- Add view definitions
		SELECT @nextid = MAX([viewid]) + 1 FROM dbo.[ASRSysViews];
		SET @NVarCommand = 'INSERT tbsys_systemobjects ([objecttype], [tablename], [viewname], [description], [nextid], [allowselect], [allowupdate])
			VALUES (3,''tbsys_views'',''ASRSysViews'', ''View definitions '' ,'  + convert(nvarchar(255),@nextid) + ', 1, 0);';
		EXEC sp_executesql @NVarCommand;

		-- Add workflow definitions
		SELECT @nextid = MAX([id]) + 1 FROM dbo.[ASRSysWorkflows];
		SET @NVarCommand = 'INSERT tbsys_systemobjects ([objecttype], [tablename], [viewname], [description], [nextid], [allowselect], [allowupdate])
			VALUES (10,''tbsys_workflows'',''ASRSysWorkflows'', ''Workflow definitions '' ,'  + convert(nvarchar(255),@nextid) + ', 1, 0);';
		EXEC sp_executesql @NVarCommand;


	END

	-- Object identity procedure
	IF OBJECT_ID('spASRGetNextObjectIdentitySeed', N'P') IS NULL	
	BEGIN

		EXEC sp_executesql N'CREATE PROCEDURE spASRGetNextObjectIdentitySeed (@viewname nvarchar(255), @nextid integer OUTPUT)
			AS
			BEGIN

				SET NOCOUNT ON;

				SELECT @nextid = [nextid]
					FROM dbo.[tbsys_systemobjects]
					WHERE [viewname] = @viewname;

				UPDATE dbo.[tbsys_systemobjects] SET [nextid] = [nextid] + 1
					WHERE [viewname] = @viewname;

			END'
	END


	-- Modification history table
	EXEC sp_executesql N'IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[tbsys_scriptedchanges]'') AND type in (N''U''))
		DROP TABLE [dbo].[tbsys_scriptedchanges]'

	IF OBJECT_ID('tbsys_scriptedchanges', N'U') IS NULL	
	BEGIN
		EXECUTE sp_executeSQL N'CREATE TABLE tbsys_scriptedchanges ([id] uniqueidentifier, [sequence] integer, [file] nvarchar(MAX), [uploaddate] datetime, [runtype] integer, [lastrundate] datetime, [runonce] bit, [runinversion] nvarchar(10), [description] nvarchar(MAX));'
	END

	-- Generate apply scripts procedure
	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRApplyScripts]') AND xtype ='P')
		DROP PROCEDURE [dbo].spASRApplyScripts

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRUploadScript]') AND xtype ='P')
		DROP PROCEDURE [dbo].spASRUploadScript

	EXEC sp_executesql N'CREATE PROCEDURE dbo.[spASRApplyScripts] (@runtype integer)
	AS
	BEGIN
		
		SET NOCOUNT ON;

		DECLARE @NVarCommand nvarchar(MAX);
		DECLARE @changes table(id uniqueidentifier, [file] nvarchar(MAX), [sequence] integer);
		
		-- Collate hotfixes
		INSERT @changes
			SELECT [id], [file], [sequence]
				FROM dbo.[tbsys_scriptedchanges]
				WHERE (runtype = @runtype) AND ([runonce] = 0 OR ([runonce] = 1 AND [lastrundate] IS NULL))
				ORDER BY [sequence];

		-- Build hotixes and apply
		SET @NVarCommand = '''';
		SELECT @NVarCommand = @NVarCommand + [file]
			FROM @changes
			ORDER BY [sequence];
		EXECUTE sp_executeSQL @NVarCommand;

		-- Mark the hotfixes as complete
		UPDATE [tbsys_scriptedchanges]
			SET [lastrundate] = GETDATE()
			FROM @changes c WHERE c.id = [tbsys_scriptedchanges].id;

	END'

	EXEC sp_executesql N'CREATE PROCEDURE dbo.spASRUploadScript
	(@runtype integer, @script nvarchar(MAX), @runonce bit, @runinversion nvarchar(10), @sequence integer, @description nvarchar(MAX))
	AS
	BEGIN

		INSERT tbsys_scriptedchanges ([sequence], [file], [uploaddate], [runtype], [runonce], [runinversion], [description])
			VALUES (@sequence, @script, GETDATE(), @runtype, @runonce, @runinversion, @description)
		
	END'


/* ------------------------------------------------------------- */
PRINT 'Step 5 - Upgrade image data structures to varbinary(max)'

	-- User defined tables
	SET @NVarCommand = ''
	SELECT @NVarCommand = @NVarCommand + 'ALTER TABLE dbo.[' + o.Name + '] ALTER COLUMN [' + c.ColumnName + '] varbinary(MAX);' 
		FROM ASRSysColumns c 
		INNER JOIN ASRSysTables t ON c.tableID = t.TableID
		INNER JOIN sys.sysobjects o ON t.tablename = o.name AND o.xtype = 'U' 
		INNER JOIN sys.syscolumns oc ON c.columnname = oc.name AND oc.id = o.id AND oc.type = 34
		WHERE (c.datatype = -4 AND c.OLEType >= 2) OR (c.datatype = -3 AND c.OLEType >= 2);
	EXECUTE sp_executesql @NVarCommand;

	-- System tables
	EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysPictures] ALTER COLUMN [Picture] varbinary(MAX);';
	EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysPermissionCategories] ALTER COLUMN [Picture] varbinary(MAX);';
	EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysWorkflowInstanceValues] ALTER COLUMN [FileUpload_File] varbinary(MAX);';
	EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysWorkflowInstanceValues] ALTER COLUMN [TempFileUpload_File] varbinary(MAX);';


/* ------------------------------------------------------------- */
PRINT 'Step 6 - Create views on metadata tables'

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spASRConvertTableToView]') AND xtype ='P')
		DROP PROCEDURE [dbo].[spASRConvertTableToView]

	SET @NVarCommand = 'CREATE PROCEDURE spASRConvertTableToView
		(@oldname nvarchar(255), @newname nvarchar(255), @IDName nvarchar(255), @ObjectType integer)
	AS
	BEGIN

		DECLARE @permissions TABLE([owner] nvarchar(255), [object] nvarchar(255), [grantee] nvarchar(255), [grantor] nvarchar(255)
			, [protecttype] nvarchar(255), [action] nvarchar(10), [column] nvarchar(MAX))

		DECLARE @NVarCommand nvarchar(MAX),
				@columnnames nvarchar(MAX);

		SET @columnnames = '''';

		IF OBJECT_ID(@newname, N''U'') IS NULL	
		BEGIN

			-- Rename existing table
			SET @NVarCommand = ''EXECUTE sp_rename '''''' + @oldname + '''''', '''''' + @newname + '''''';'';
			EXECUTE sp_executesql @NVarCommand

			-- Drop existing view
			IF EXISTS(SELECT * FROM sys.sysobjects WHERE name = @oldname AND type = ''V'')
			BEGIN
				SET @NVarCommand = ''DROP VIEW dbo.'' + @oldname
				EXECUTE sp_executesql @NVarCommand
			END

			-- Build list of columns for the view (exclude some). Needed because select * does not allow indexing
			SELECT @columnnames = @columnnames + ''base.['' + syscolumns.name + ''], ''  
				FROM sysobjects 
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.xtype=''U''
					AND sysobjects.id =  OBJECT_ID(@newname)
					AND NOT (syscolumns.name = ''lastupdated'' OR syscolumns.name = ''lastupdatedby'')
			ORDER BY sysobjects.name,syscolumns.colid

			-- Generate the view
			SET @NVarCommand = ''CREATE VIEW dbo.['' + @oldname + '']
					WITH SCHEMABINDING
					AS SELECT '' + LOWER(@columnnames) + '' obj.[locked], obj.[lastupdated], obj.[lastupdatedby]
						FROM dbo.['' + @newname + ''] base
						INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.'' + @IDName + '' AND obj.objecttype = '' + convert(nvarchar(2),@ObjectType) + ''
						INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
						WHERE obj.effectivedate <= dt.[date]''
			EXECUTE sp_executesql @NVarCommand

			-- Generate index
			SET @NVarCommand = ''CREATE UNIQUE CLUSTERED INDEX [idx_'' + @IDName + ''] ON [dbo].['' + @oldname + ''](['' + @IDName + ''] ASC)''
			EXECUTE sp_executesql @NVarCommand


			-- Drop existing triggers on the base table
			IF EXISTS(SELECT id FROM sys.sysobjects o WHERE o.xtype = ''TR'' AND name = ''INS_'' + @newname)
			BEGIN
				SET @NVarCommand = ''DROP TRIGGER [INS_'' + @newname +'' ];''
				EXECUTE sp_executesql @NVarCommand;
			END


			SET @columnnames = ''''
			SELECT @columnnames = @columnnames + ''['' + syscolumns.name + ''], ''  
				FROM sysobjects 
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.xtype=''U''
					AND sysobjects.id =  OBJECT_ID(@newname)
					AND NOT (syscolumns.name = ''lastupdated'' OR syscolumns.name = ''lastupdatedby'')
			ORDER BY sysobjects.name,syscolumns.colid

			-- Generate triggers on the the scripted view
			SET @NVarCommand = ''CREATE TRIGGER INS_'' + @oldname +'' ON [dbo].['' + @oldname + '']
				INSTEAD OF INSERT
				AS
				BEGIN

					SET NOCOUNT ON;

					-- Update objects table
					IF NOT EXISTS(SELECT [guid]
						FROM dbo.[tbsys_scriptedobjects] o
						INNER JOIN inserted i ON i.'' + @IDName + '' = o.targetid AND o.objecttype = '' + convert(nvarchar(2),@ObjectType) + '')
					BEGIN
						INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
							SELECT NEWID(), '' + convert(nvarchar(2),@ObjectType) + '', ['' + @IDName + ''], dbo.[udfsys_getownerid](), ''''01/01/1900'''',1,0, GETDATE()
								FROM inserted;
					END

					-- Update base table								
					INSERT dbo.['' + @newname + ''] ('' + SUBSTRING(@columnnames,0,LEN(@columnnames)) + '') 
						SELECT '' + SUBSTRING(@columnnames,0,LEN(@columnnames)) + '' FROM inserted;

				END'';
				EXECUTE sp_executesql @NVarCommand;

			-- Generate triggers on the scripted view
			SET @NVarCommand = ''CREATE TRIGGER [dbo].[DEL_'' + @oldname + ''] ON [dbo].['' + @oldname + '']
				INSTEAD OF DELETE
				AS
				BEGIN
					SET NOCOUNT ON;

					DELETE FROM ['' + @newname + ''] WHERE '' + @IDName + '' IN (SELECT '' + @IDName + '' FROM deleted);
				END''
			EXECUTE sp_executesql @NVarCommand;


			-- Grant permissions on this view
			SET @NVarCommand = ''GRANT SELECT, INSERT, UPDATE, DELETE ON '' + @oldname + '' TO [ASRSysAdmins];'';
			EXECUTE sp_executesql @NVarCommand;

			SET @NVarCommand = ''GRANT SELECT ON '' + @oldname + '' TO [ASRSysGroup];'';
			EXECUTE sp_executesql @NVarCommand;

		END
	END'
	EXECUTE sp_executesql @NVarCommand;

	-- Change the metadata table to new structure
	EXEC dbo.spASRConvertTableToView 'ASRSysTables', 'tbsys_tables', 'tableid', 1;
	EXEC dbo.spASRConvertTableToView 'ASRSysColumns', 'tbsys_columns', 'columnid', 2;
	EXEC dbo.spASRConvertTableToView 'ASRSysViews', 'tbsys_views', 'viewid', 3;
	EXEC dbo.spASRConvertTableToView 'ASRSysWorkflows', 'tbsys_workflows', 'id', 10;
	--EXEC dbo.spASRConvertTableToView 'ASRSysScreens', 'tbsys_screens', 'screenid', 14;

/* ------------------------------------------------------------- */
PRINT 'Step 7 - Drop existing system triggers'

	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'DROP TRIGGER ' +  o.name + ';' + CHAR(13)
		FROM sys.sysobjects o
		INNER JOIN ASRSysTables t ON t.TableName = OBJECT_NAME(o.parent_obj)
		WHERE o.xtype = 'TR' AND (name = 'INS_' + t.TableName OR name = 'UPD_' + t.TableName OR name = 'DEL_' + t.TableName)
	EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 8 - Add abstraction layer to user defined tables'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRConvertDataTablesToViews]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRConvertDataTablesToViews];

	SET @NVarCommand = 'CREATE PROCEDURE dbo.spASRConvertDataTablesToViews(@oldname nvarchar(255))
		AS
		BEGIN

			DECLARE	@newname nvarchar(MAX),
					@sqlCommand nvarchar(MAX),
					@columnnames nvarchar(MAX),
					@sqlApplyPermissions nvarchar(MAX),
					@sqlRevokePermissions nvarchar(MAX);

			DECLARE @permissions TABLE([owner] nvarchar(255), [object] nvarchar(255), [grantee] nvarchar(255), [grantor] nvarchar(255)
					, [protecttype] nvarchar(255), [action] nvarchar(10), [column] nvarchar(MAX))

			SET @newname = ''tbuser_'' + @oldname;
			SET @columnnames = '''';
			SET @sqlApplyPermissions = '''';
			SET @sqlRevokePermissions = '''';

			IF EXISTS(SELECT name FROM sys.sysobjects WHERE name = @oldname AND xtype = ''U'')
			BEGIN
			
				-- Rename the original object
				EXECUTE sp_rename @oldname, @newname;

				-- Build list of columns for the view (exclude some). Needed because select * does not allow indexing
				SELECT @columnnames = @columnnames + ''['' + syscolumns.name + ''], ''  
					FROM sysobjects 
					INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
					WHERE sysobjects.xtype=''U''
						AND sysobjects.id =  OBJECT_ID(@newname) AND NOT syscolumns.name = ''ID''
				ORDER BY sysobjects.name,syscolumns.colid;

				-- Create the view on this object
				SET @sqlCommand = ''CREATE VIEW dbo.['' + @oldname + '']
										WITH SCHEMABINDING
										AS SELECT '' + @columnnames + ''[ID] FROM dbo.['' + @newname + ''];'';
				EXECUTE sp_executesql @sqlCommand;

				-- Read the security for the base table
				INSERT @permissions
					EXEC sp_helprotect @name = @newname, @grantorname = ''dbo'';
		 
				-- Apply the permissions onto the view
				SELECT @sqlApplyPermissions = @sqlApplyPermissions + p.protecttype + '' '' + p.[action] + '' ON '' +  @oldname +
					CASE p.[column]
						WHEN ''.'' THEN ''''
						WHEN ''(All+New)'' THEN ''''
						ELSE ''('' + p.[column] + '')''
					END
					+ '' TO ['' + p.[grantee] + + ''];'' + CHAR(13) FROM @permissions p;
				EXECUTE sp_executesql @sqlApplyPermissions;

				-- Revoke existing permissions on the base table
				SELECT @sqlRevokePermissions = @sqlRevokePermissions + ''REVOKE SELECT, UPDATE, DELETE, INSERT ON '' + p.[object] + '' TO ['' + p.[grantee] + ''];'' + CHAR(13)
					FROM @permissions p
				EXECUTE sp_executesql @sqlRevokePermissions;

			END
		END';
	EXECUTE sp_executesql @NVarCommand;

	-- Move the user defined tables
	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'EXECUTE dbo.spASRConvertDataTablesToViews ''' + TableName + ''';'
		FROM ASRSysTables;
	EXECUTE sp_executesql @NVarCommand;

	--SET @NVarCommand = '';
	--SELECT @NVarCommand = @NVarCommand + 'EXECUTE dbo.spASRUpdateTableStructures ''' + TableName + ''';'
	--	FROM tbsys_tables;
	--EXECUTE sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 9 - Drop all HR Pro defined object (schema binding)'

-- Can't drop until we've backed up the security model!!!!!

	-- Table Views
	--SELECT @NVarCommand = @NVarCommand + 'DROP VIEW dbo.[' + o.name + '];'
	--	FROM dbo.sysobjects o
	--	INNER JOIN tbsys_tables t ON t.tablename = o.name
	--	WHERE o.xtype= 'V'
	--EXECUTE sp_executesql @NVarCommand;

	-- Views
	--SELECT @NVarCommand = @NVarCommand + 'DROP VIEW dbo.[' + o.name + '];'
	--	FROM dbo.sysobjects o
	--	INNER JOIN tbsys_views v ON v.viewname = o.name
	--	WHERE o.xtype= 'V'

	--PRINT @NVarCommand;
	--EXECUTE sp_executesql @NVarCommand;

	-- Calculations
	SELECT @NVarCommand = @NVarCommand + 'DROP FUNCTION dbo.[' + name + '];'
		FROM dbo.sysobjects
		WHERE name LIKE 'udfcalc_%'
			AND xtype in (N'FN', N'IF', N'TF')
	EXECUTE sp_executesql @NVarCommand;



/* ------------------------------------------------------------- */
PRINT 'Step 10 - Convert audit table to view'

	-- Rename the base audit log
	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[tbsys_audittrail]') AND xtype = 'U')
	BEGIN
		EXECUTE sp_executesql N'EXECUTE sp_rename [ASRSysAuditTrail], [tbsys_audittrail];';
		EXEC spsys_setsystemsetting 'integration', 'auditlog', 0;
	END

	SET @sSPCode = '';
	IF EXISTS(SELECT dbo.[udfsys_getmodulesetting]('MODULE_AUDIT','Param_AuditTable'))
	BEGIN
		SELECT @sSPCode = 'IF EXISTS(SELECT * FROM dbo.sysobjects WHERE id = object_id(N''' + [tablename] + ''', ''V''))
			DROP VIEW dbo.[' + [tablename] + ']' FROM dbo.[tbsys_tables]
			WHERE tableid = dbo.[udfsys_getmodulesetting]('MODULE_AUDIT','Param_AuditTable');
		EXECUTE sp_executesql @sSPCode;
		
		SELECT @sSPCode = 'CREATE VIEW dbo.[' + [tablename] + '] AS SELECT * FROM dbo.[tbuser_' + [tablename] + ']' FROM dbo.[tbsys_tables]
			WHERE tableid = dbo.[udfsys_getmodulesetting]('MODULE_AUDIT','Param_AuditTable');
	--	EXECUTE sp_executesql @sSPCode;

	END

	-- Remove old audit view
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysAuditTrail]') AND xtype = 'V')
		DROP VIEW [dbo].[ASRSysAuditTrail];
		
	-- Alter structure of the base table
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbsys_audittrail', 'U') AND name = 'tableid')
		EXECUTE sp_executesql N'ALTER TABLE dbo.[tbsys_audittrail] ADD [tableid] integer NULL'

	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbsys_audittrail', 'U') AND name = 'tablename')
		EXECUTE sp_executesql N'ALTER TABLE dbo.[tbsys_audittrail] DROP COLUMN [tablename]'

	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbsys_audittrail', 'U') AND name = 'columnname')
		EXECUTE sp_executesql N'ALTER TABLE dbo.[tbsys_audittrail] DROP COLUMN [columnname]'

	-- Create audit view
	EXECUTE sp_executesql N'CREATE VIEW [ASRSysAuditTrail]
		WITH SCHEMABINDING
		AS SELECT
			a.[id], a.[UserName], a.[DateTimeStamp], a.[RecordID], a.[RecordDesc], a.[OldValue], a.[NewValue],
			t.[Tablename], c.[Columnname], a.[CMGExportDate],	a.[CMGCommitDate], a.[ColumnID], a.[Deleted]
		FROM [dbo].[tbsys_audittrail] a
		INNER JOIN dbo.[tbsys_tables] t ON t.[tableid] = a.[tableid]
		INNER JOIN dbo.[tbsys_columns] c ON c.[columnid] = a.[columnid] AND t.[tableid] = c.[tableid]'

	-- Remove old triggers 
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[INS_ASRSysAuditTrail]') AND xtype = 'TR')
		DROP TRIGGER [dbo].[INS_ASRSysAuditTrail];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[DEL_ASRSysAuditTrail]') AND xtype = 'TR')
		DROP TRIGGER [dbo].[DEL_ASRSysAuditTrail];

	-- Generate triggers on the audit logs		
	EXECUTE sp_executesql N'CREATE TRIGGER [dbo].[DEL_ASRSysAuditTrail] ON [dbo].[ASRSysAuditTrail]
		INSTEAD OF DELETE
		AS
		BEGIN
			SET NOCOUNT ON;
			DELETE FROM dbo.[tbsys_audittrail] WHERE ID IN (SELECT ID FROM deleted);
		END'


/* ------------------------------------------------------------- */
PRINT 'Step 11 - Add new calculation procedures'

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_statutoryredundancypay]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_statutoryredundancypay];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfstat_MaternityExpectedReturn]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfstat_MaternityExpectedReturn];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfstat_ParentalLeaveEntitlement]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfstat_ParentalLeaveEntitlement];
	
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_convertcharactertonumeric]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_convertcharactertonumeric];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_convertcurrency]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_convertcurrency];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_divide]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].udfsys_divide;

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_fieldchangedbetweentwodates]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_fieldchangedbetweentwodates];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_fieldlastchangedate]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_fieldlastchangedate];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_firstnamefromforenames]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_firstnamefromforenames];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_getfieldfromdatabaserecord]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getfieldfromdatabaserecord];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_getfunctionparametertype]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getfunctionparametertype];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_getuniquecode]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_getuniquecode];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_initialsfromforenames]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_initialsfromforenames];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isbetween]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isbetween];
	
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isfieldempty]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isfieldempty];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isfieldpopulated]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isfieldpopulated];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isnivalid]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isnivalid];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isovernightprocess]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isovernightprocess];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_isvalidpayrollcharacterset]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_isvalidpayrollcharacterset];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_justdate]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_justdate];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_maternityexpectedreturndate]')	AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_maternityexpectedreturndate];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_nicedate]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_nicedate];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_nicetime]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_nicetime];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_parentalleaveentitlement]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_parentalleaveentitlement];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_parentalleavetaken]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_parentalleavetaken];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_propercase]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_propercase];
	
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_remainingmonthssincewholeyears]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_remainingmonthssincewholeyears];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_roundtostartofnearestmonth]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_roundtostartofnearestmonth];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_servicelength]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_servicelength];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_uniquecode]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_uniquecode];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_username]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_username];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_weekdaysbetweentwodates]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_weekdaysbetweentwodates];
		
	IF EXISTS (SELECT *	FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[udfsys_wholemonthsbetweentwodates]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_wholemonthsbetweentwodates];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_wholeyearsbetweentwodates]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_wholeyearsbetweentwodates];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_workingdaysbetweentwodates]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_workingdaysbetweentwodates];
	

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfstat_MaternityExpectedReturn] (
			@EWCDate datetime,
			@LeaveStart datetime,
			@BabyBirthDate datetime,
			@Ordinary varchar(MAX)
			)
	RETURNS datetime			
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @pdblResult datetime;

		IF LOWER(@Ordinary) = ''ordinary''
			IF DATEDIFF(d,''04/06/2003'', @EWCDate) >= 0
				SET @pdblResult = DATEADD(ww,26,@LeaveStart);
			ELSE
				IF DATEDIFF(d,''04/30/2000'', @EWCDate) >= 0
					SET @pdblResult = DATEADD(ww,18,@LeaveStart);
				ELSE
					SET @pdblResult = DATEADD(ww,14,@LeaveStart);
		ELSE
			IF DATEDIFF(d,''04/06/2003'', @EWCDate) >= 0
				SET @pdblResult = DATEADD(ww,52,@LeaveStart);
			ELSE
				--29 weeks from baby birth date (but return on the monday before!)
				SET @pdblResult = DATEADD(d,203 - datepart(dw,DATEADD(d,-2,@BabyBirthDate)),@BabyBirthDate);

		RETURN @pdblResult;

	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfstat_ParentalLeaveEntitlement] (
		@DateOfBirth	datetime,
		@AdoptedDate	datetime,
		@Disabled		bit,
		@Region			varchar(MAX))
	RETURNS float
	AS
	BEGIN

		DECLARE @pdblResult			float,
			@Today					datetime,
			@ChildAge				integer,
			@Adopted				bit,
			@YearsOfResponsibility	integer,
			@StartDate				datetime,
			@Standard				integer,
			@Extended				integer;

		SET @Standard = 65;
		SET @Extended = 90;
		IF @Region = ''Rep of Ireland''
		BEGIN
			SET @Standard = 70;
			SET @Extended = 70;
		END

		-- Check if we should used the Date of Birth or the Date of Adoption column...
		SET @Adopted = 0;
		SET @StartDate = @DateOfBirth;
		IF NOT @AdoptedDate IS NULL
		BEGIN
			SET @Adopted = 1;
			SET @StartDate = @AdoptedDate;
		END

		-- Set variables based on this date...
		--( years of responsibility = years since born or adopted)
		SET @Today = GETDATE();
		SELECT @ChildAge = [dbo].[udfsys_wholeyearsbetweentwodates](@DateOfBirth, @Today);
		SELECT @YearsOfResponsibility = [dbo].[udfsys_wholeyearsbetweentwodates](@StartDate, @Today);

		SELECT @pdblResult = CASE
			WHEN @Disabled = 0 AND @Adopted = 0 AND @ChildAge < 5
				THEN @Standard
			WHEN @Disabled = 0 AND @Adopted = 1 AND @ChildAge < 18
				AND @YearsOfResponsibility < 5 THEN	@Standard
			WHEN @Disabled = 1 AND @Adopted = 0 AND @ChildAge < 18 
				AND DATEDIFF(d,''12/15/1994'',@DateOfBirth) >= 0 THEN @Extended
			WHEN @Disabled = 1 AND @Adopted = 1 AND @ChildAge < 18 
				AND DATEDIFF(d,''12/15/1994'',@AdoptedDate) >= 0 THEN @Extended
			ELSE 0
			END;

		RETURN ISNULL(@pdblResult,0);

	END'
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_weekdaysbetweentwodates](
		@datefrom AS datetime,
		@dateto AS datetime)
	RETURNS integer
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result integer;

		SELECT @result = CASE 
			WHEN DATEDIFF (day, @datefrom, @dateto) <= 0 THEN 0
			ELSE DATEDIFF(day, @datefrom, @dateto + 1) 
				- (2 * (DATEDIFF(day, @datefrom - (DATEPART(dw, @datefrom) -1),
					@dateto	- (DATEPART(dw, @dateto) - 1)) / 7))
				- CASE WHEN DATEPART(dw, @datefrom) = 1 THEN 1 ELSE 0 END
				- CASE WHEN DATEPART(dw, @dateto) = 7 THEN 1 ELSE 0	END
				END;
				
		RETURN @result;
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_wholemonthsbetweentwodates](
		@date1 	datetime,
		@date2 	datetime)
	RETURNS integer
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result integer;
	
		-- Clean dates (trim time part)
		SET @date1 = DATEADD(D, 0, DATEDIFF(D, 0, @date1));
		SET @date2 = DATEADD(D, 0, DATEDIFF(D, 0, @date2));
	
		IF @date1 < @date2
		BEGIN
	
			-- Get the total number of months
			SET @result = DATEDIFF(mm, @date1, @date2);
	      
			-- See if the day field of pvParam2 < pvParam1 day field and if so - 1
			IF DAY(@date2) < DAY(@date1)
			BEGIN
				SET @result = @result -1;
			END
		END
		
		RETURN @result
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_wholeyearsbetweentwodates] (
	     @date1  datetime,
	     @date2  datetime)
	RETURNS integer 
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result integer;
		
	    -- Get the number of whole years
	    SET @result = YEAR(@date2) - YEAR(@date1);
	
	    -- See if the date passed in months are greater than todays month
	    IF MONTH(@date1) > MONTH(@date2)
	    BEGIN
			SET @result = @result - 1;
	    END
	    
	    -- See if the months are equal and if they are test the day value
	    IF MONTH(@date1) = MONTH(@date2)
	    BEGIN
	        IF DAY(@date1) > DAY(@date2)
	            BEGIN
					SET @result = @result - 1;
	            END
	        END
	        
	    RETURN @result;
	
	END';
	EXECUTE sp_executeSQL @sSPCode;


	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_convertcharactertonumeric]
		(@psToConvert nvarchar(MAX))
	RETURNS numeric(38,8)
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @result numeric(38,8);

		SET @result = 0;

		IF ISNUMERIC(@psToConvert) > 0
			SET @result = CONVERT(NUMERIC(38,8), @psToConvert);

		RETURN @result;

	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_divide](@value numeric(38,8), @divideby numeric(38,8))
	RETURNS numeric(38,8)
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @result numeric(38,8);
		
		IF @divideby = 0 SET @result = 0;
		ELSE set @result = @value / @divideby;

		RETURN @result;

	END'
	EXECUTE sp_executeSQL @sSPCode;




	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_firstnamefromforenames] (
		@forenames nvarchar(max))
	RETURNS nvarchar(max)
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result nvarchar(max);
	
		IF (LEN(@forenames) = 0 ) OR (@forenames IS null)
		BEGIN
			SET @result = '''';
		END
		ELSE
		BEGIN
			IF CHARINDEX('' '', @forenames) > 0
				SET @result = LEFT(@forenames, CHARINDEX('' '', @forenames));
			ELSE
				SET @result = @forenames;
		END
		
		RETURN @result;
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_fieldchangedbetweentwodates](
		@colrefID	varchar(32),
		@fromdate	datetime,
		@todate		datetime,
		@recordID	integer
	)
	RETURNS bit
	AS
	BEGIN

		DECLARE @result		bit,
				@tableid	integer,
				@columnid	integer;
		
		SET @tableid = SUBSTRING(@colrefID, 1, 8);
		SET @columnid = SUBSTRING(@colrefID, 10, 8);
		SET @fromdate = DATEADD(dd, 0, DATEDIFF(dd, 0, @fromdate));
		SET @todate = DATEADD(dd, 0, DATEDIFF(dd, 0, @todate));

		SELECT @result = CASE WHEN
				EXISTS(SELECT [DateTimeStamp] FROM [tbsys_audittrail]
					WHERE [ColumnID] = @columnid AND [TableID] = @tableID
					AND @recordID = [RecordID] 
					AND [DateTimeStamp] >= @fromdate AND DateTimeStamp < @todate + 1)
				THEN 1 ELSE 0 END;

		RETURN @result;

	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_fieldlastchangedate](
		@colrefID	varchar(32),
		@recordID	integer
	)
	RETURNS datetime
	AS
	BEGIN

		DECLARE @result		datetime,
				@tableid	integer,
				@columnid	integer;
		
		SET @tableid = SUBSTRING(@colrefID, 1, 8);
		SET @columnid = SUBSTRING(@colrefID, 10, 8);

		SELECT TOP 1 @result = [DateTimeStamp] FROM [tbsys_audittrail]
			WHERE [ColumnID] = @columnid AND [TableID] = @tableID
				AND @recordID = [RecordID]
			ORDER BY [DateTimeStamp] DESC ;

		RETURN @result;

	END'
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getfieldfromdatabaserecord](
			@searchcolumn AS nvarchar(255),
			@searchexpression AS nvarchar(MAX),
			@returnfield AS nvarchar(255))
		RETURNS nvarchar(MAX)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result nvarchar(MAX);
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getuniquecode](
			@prefix AS nvarchar(255),
			@rootvalue as integer,
			@rowoffset AS integer)
		RETURNS [nvarchar](255)
		AS
		BEGIN
		
			DECLARE @result nvarchar(255);

			--SELECT @result = SUBSTRING([maxcodesuffix], LEN(LEFT([maxcodesuffix] + ''1''
			--	, PATINDEX(''%[0-9]%'', [maxcodesuffix] + ''1'') )), 100)
			--	+ @rowoffset
			--	FROM ASRSysUniqueCodes WHERE [codeprefix] = @prefix;

			SELECT @result = convert(nvarchar(255),@rowoffset);
			
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_getfunctionparametertype]
			(@functionid integer, @parameterindex integer)
		RETURNS integer
		AS
		BEGIN
		
			DECLARE @result integer;
		
			SELECT @result = [parametertype] FROM ASRSysFunctionParameters
				WHERE @functionid = [functionID] AND @parameterindex = [parameterIndex];
		
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_initialsfromforenames] 
		(
			@forenames	varchar(8000),
			@padwithspace bit
		)
		RETURNS nvarchar(10)
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result nvarchar(10);
			DECLARE @icounter integer;

			SET @result = '''';
			SET @icounter = 1;
		
			IF LEN(@forenames) > 0 
			BEGIN
				SET @result = UPPER(left(@forenames,1));
		
				WHILE @icounter < LEN(@forenames)
				BEGIN
					IF SUBSTRING(@forenames, @icounter, 1) = '' ''
					BEGIN
						IF @padwithspace = 1
							SET @result = @result + '' '' + UPPER(SUBSTRING(@forenames, @icounter+1, 1));
						ELSE
							SET @result = @result + UPPER(SUBSTRING(@forenames, @icounter+1, 1));
					END
			
					SET @icounter = @icounter +1;
				END
		
				SET @result = @result + '' ''
			
			END
		
			RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isbetween](
			@column1 AS nvarchar(MAX),
			@column2 AS nvarchar(MAX),
			@column3 AS nvarchar(MAX))
		RETURNS bit
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit;
			
			SET @result = 1;
			SELECT @result = CASE WHEN @column1 BETWEEN @column2 AND @column3 THEN 1 ELSE 0 END
				
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isfieldpopulated](
			@inputcolumn as nvarchar(MAX))
		RETURNS bit
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit;
			
			SET @result = 0;
			SELECT @result = (
				CASE 
					WHEN @inputcolumn IS NULL THEN 0 
					ELSE
						CASE
		--					WHEN LEN(convert(nvarchar(1),@inputcolumn)) = 0 THEN 0
							WHEN DATALENGTH(@inputcolumn) = 0 THEN 0
							ELSE 1
						END
					END);
		
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isfieldempty](
		@inputcolumn as nvarchar(MAX))
	RETURNS bit
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result bit;
		
		SET @result = 0;			
		SELECT @result = (
			CASE 
				WHEN @inputcolumn IS NULL THEN 1
				ELSE
					CASE
						WHEN DATALENGTH(@inputcolumn) = 0 THEN 1
						ELSE 0
					END
				END);
	
		RETURN @result;
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isovernightprocess] ()
	RETURNS bit 
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result bit;
		
		SET @result = 0;
	    RETURN @result;
	
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_nicedate](
		@inputdate as datetime)
	RETURNS nvarchar(max)
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result varchar(MAX);
		
		SET @result = '''';
		SELECT @result = DATENAME(dw, @inputdate) + '', '' 
			+ DATENAME(mm, @inputdate) + '' '' 
			+ LTRIM(STR(DATEPART(dd, @inputdate))) 
			+ '' '' + LTRIM(STR(DATEPART(yy, @inputdate)));

		RETURN @result;
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_nicetime](
		@inputdate as varchar(20))
	RETURNS nvarchar(255)
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result varchar(255);
	
		SELECT @Result = 
			CASE
			WHEN LEN(LTRIM(RTRIM(@inputdate))) = 0 then ''''
			ELSE 
				CASE 
					WHEN ISDATE(@inputdate) = 0 THEN ''***''
					ELSE (CONVERT(varchar(2),((DATEPART(hour,CONVERT(datetime, @inputdate)) + 11) % 12) + 1)
						+ '':'' + RIGHT(''00'' + DATENAME(minute, CONVERT(datetime, @inputdate)),2)
						+ CASE 
							WHEN DATEPART(hour, CONVERT(datetime, @inputdate)) > 11 THEN '' pm''
							ELSE '' am''
						END) 
				END 
		END		
		RETURN @result;
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_propercase](
		@text as nvarchar(max))
	RETURNS nvarchar(max)
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @reset bit;
		DECLARE @result varchar(8000);
		DECLARE @i integer;
		DECLARE @c char(1);

		SET @i = 1;
		SET @result = '''';
		SET @reset = 1;
	      
		WHILE (@i <= len(@text))
			SELECT @c= substring(@text,@i,1)
				, @result = @result + CASE WHEN @reset=1 THEN UPPER(@c) 
										   ELSE LOWER(@c) END
				, @reset = CASE WHEN @c LIKE ''[a-zA-Z]'' THEN 0
								ELSE 1
								END
				, @i = @i + 1;
	
		RETURN @result;
		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_remainingmonthssincewholeyears]
		(@pdtDate 	datetime, @dtToday datetime)
	RETURNS integer
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @iResult integer;

		SET @pdtDate = convert(datetime, convert(varchar(20), @pdtDate, 101));

		-- Get the number of whole months
		SET @iResult = month(@dtToday) - month(@pdtDate);
	 
		-- Test the day value
		IF DAY(@pdtDate) > DAY(@dtToday)
			SET @iResult = @iResult - 1;

		IF @iResult < 0
			SET @iResult = @iResult + 12;

		RETURN @iResult;

	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_roundtostartofnearestmonth]
		(@pdtDate 	datetime)
	RETURNS datetime
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @dtDateNextMonth	datetime,
				@dtDateThisMonth 	datetime,
				@dtResult			datetime;

		SET @pdtDate = convert(datetime, convert(varchar(20), @pdtDate, 101));

		-- Create a date with one month added to the date and move it to the first day of that month
		SET @dtDateNextMonth = DATEADD(mm, 1, @pdtDate);
		SET @dtDateNextMonth = DATEADD(dd, -1 * (DAY(@dtDateNextMonth) - 1), @dtDateNextMonth);

		-- Create a date which is the first of the month passed in
		SET @dtDateThisMonth = DATEADD(dd, -1 * (DAY(@pdtDate) - 1), @pdtDate);
	    
		-- See which is the greatest gap between the two start month dates and the passed in date
		IF (@pdtDate - (@dtDateThisMonth) + 1) < ((@dtDateNextMonth) - (@pdtDate))
			SET @dtResult = @dtDateThisMonth
		ELSE
			SET @dtResult = @dtDateNextMonth;

		RETURN @dtResult;

	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_servicelength] (
		     @startdate  datetime,
		     @leavingdate  datetime,
		     @period nvarchar(2))
		RETURNS integer 
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result integer;
			DECLARE @amount integer;
		
			-- If start date is in the future ignore
			IF @startdate > GETDATE()
				RETURN 0;
			
			-- Trim the leaving date
			IF @leavingdate IS NULL OR @leavingdate > GETDATE()
				SET @leavingdate = GETDATE();
		
			SET @amount = [dbo].[udfsys_wholeyearsbetweentwodates](@startdate, @leavingdate);
		
			-- Years
			IF @period = ''Y'' SET @result = @amount
			
			--Months
			ELSE IF @period = ''M''
				SET @result = [dbo].[udfsys_wholemonthsbetweentwodates]
					(@startdate, @leavingdate) - (@amount * 12);
			
		    RETURN @result;
		
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_statutoryredundancypay] 	(
		@pdtStartDate 		datetime,
		@pdtLeaveDate 		datetime,
		@pdtDOB				datetime,
		@pdblWeeklyRate 	float,
		@pdblStatLimit 		float)
	RETURNS float	
	WITH SCHEMABINDING
	AS
	BEGIN
		DECLARE @pdblRedundancyPay	float,
				@dtMinAgeBirthday	datetime,
				@dtServiceFrom		datetime,
				@iServiceYears 		integer,
				@iAgeY				integer,
				@iAgeM 				integer,
				@dblRate1 			float,
				@dblRate2 			float,
				@dblRate3 			float,
				@dtTempDate 		datetime,
				@iTempAgeY			integer,
				@iTemp				integer,
				@dblTemp2 			float,
				@iAfterOct2006		bit,
				@iMinAge			integer;
	
		SET @pdblRedundancyPay = 0;
		SET @iAfterOct2006 = CASE WHEN DATEDIFF(dd,@pdtLeaveDate,''10/01/2006'') <= 0 THEN 1 ELSE 0 END;
	
		IF @iAfterOct2006 = 1
			SET @iMinAge = 15;
		ELSE
			SET @iMinAge = 18;
	
		-- First three parameters are compulsory, so return 0 and exit if they are not set
		IF (@pdtStartDate IS NULL) OR (@pdtLeaveDate IS NULL) OR (@pdtDOB IS NULL) RETURN 0;
	
		SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))
		SET @pdtLeaveDate = convert(datetime, convert(varchar(20), @pdtLeaveDate, 101))
		SET @pdtDOB = convert(datetime, convert(varchar(20), @pdtDOB, 101))

		-- Calculate start date
	   	SET @dtServiceFrom = @pdtStartDate;
		if @iAfterOct2006 = 0
		BEGIN
			SET @dtMinAgeBirthday = DATEADD(yy, @iMinAge, @pdtDOB);
			IF @dtMinAgeBirthday >= @pdtStartDate
				SET @dtServiceFrom = @dtMinAgeBirthday;
		END

		-- Calculate number of applicable complete yrs the employee has been employed
		SELECT @iServiceYears = dbo.udfsys_wholeyearsbetweentwodates(@dtServiceFrom, @pdtLeaveDate);
		
		-- Exit if its less than 2 years
		IF @iServiceYears < 2 RETURN 0;
	
		-- Calculate the employees years and months to the leave date
		SELECT @iAgeY = dbo.udfsys_wholeyearsbetweentwodates(@pdtDOB, @pdtLeaveDate);

		SET @dtTempDate = DATEADD(yy, @iAgeY, @pdtDOB);
		SELECT @iAgeM = dbo.udfsys_wholemonthsbetweentwodates(@dtTempDate, @pdtLeaveDate);
	
		-- Only count up to 20 years for redundancy
		SELECT @iServiceYears =	CASE WHEN @iServiceYears < 20 THEN @iServiceYears ELSE 20 END;
	
		-- Fill in the rates depending on service and age
		SET @iTempAgeY = @iAgeY;
		SET @dblRate1 = 0;
		SET @dblRate2 = 0;
		SET @dblRate3 = 0;
	
		IF @iTempAgeY >= 41
		BEGIN
			SET @iTemp = @iTempAgeY - 41;
			SELECT @dblRate1 = CASE WHEN @iServiceYears < @iTemp THEN @iServiceYears ELSE @iTemp END;
			SET @iTempAgeY = 41;
			SET @iServiceYears = @iServiceYears - @dblRate1;
		END
	
		IF @iTempAgeY >= 22
		BEGIN
			SET @iTemp = @iTempAgeY - 22;
			SELECT @dblRate2 = CASE WHEN @iServiceYears < @iTemp THEN @iServiceYears ELSE @iTemp END;
			SET @iTempAgeY = 22;
			SET @iServiceYears = @iServiceYears - @dblRate2;
		END
	
		IF @iTempAgeY >= @iMinAge
		BEGIN
			SET @iTemp = @iTempAgeY - @iMinAge;
			SELECT @dblRate3 = CASE WHEN @iServiceYears < @iTemp THEN @iServiceYears ELSE @iTemp END;
		END
	
		-- Calculate the redundancy pay
		SELECT @dblTemp2 = CASE WHEN @pdblStatLimit < @pdblWeeklyRate THEN @pdblStatLimit ELSE @pdblWeeklyRate END;
	
		SET @pdblRedundancyPay = ((@dblRate1 * 1.5) + (@dblRate2) + (@dblRate3 * 0.5)) * @dblTemp2;
	
		IF @iAfterOct2006 = 0 AND @iAgeY = 64 
			SET @pdblRedundancyPay = @pdblRedundancyPay * (12 - @iAgeM) / 12;

		RETURN @pdblRedundancyPay;

	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_username]
		(@userid as integer)
	RETURNS varchar(255)
	WITH SCHEMABINDING
	AS
	BEGIN	
		RETURN SYSTEM_USER;		
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_workingdaysbetweentwodates](
		@date1 	datetime,
		@date2 	datetime)
	RETURNS integer
	WITH SCHEMABINDING
	AS
	BEGIN		
		RETURN 0;			
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_justdate](
		@date 	datetime)
	RETURNS datetime
	WITH SCHEMABINDING
	AS
	BEGIN
		RETURN DATEADD(D, 0, DATEDIFF(D, 0, @date));
	END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isnivalid](
			@input AS nvarchar(MAX))
		RETURNS bit
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit;
			
			DECLARE @ValidPrefixes varchar(MAX);
			DECLARE @ValidSuffixes varchar(MAX);
			DECLARE @Prefix varchar(MAX);
			DECLARE @Suffix varchar(MAX);
			DECLARE @Numerics varchar(MAX);

			SET @result = 1;
			IF ISNULL(@input,'''') = '''' RETURN 1

			SET @ValidPrefixes = 
				''/AA/AB/AE/AH/AK/AL/AM/AP/AR/AS/AT/AW/AX/AY/AZ'' +
				''/BA/BB/BE/BH/BK/BL/BM/BT'' +
				''/CA/CB/CE/CH/CK/CL/CR'' +
				''/EA/EB/EE/EH/EK/EL/EM/EP/ER/ES/ET/EW/EX/EY/EZ'' +
				''/GY'' +
				''/HA/HB/HE/HH/HK/HL/HM/HP/HR/HS/HT/HW/HX/HY/HZ'' +
				''/JA/JB/JC/JE/JG/JH/JJ/JK/JL/JM/JN/JP/JR/JS/JT/JW/JX/JY/JZ'' +
				''/KA/KB/KE/KH/KK/KL/KM/KP/KR/KS/KT/KW/KX/KY/KZ'' +
				''/LA/LB/LE/LH/LK/LL/LM/LP/LR/LS/LT/LW/LX/LY/LZ'' +
				''/MA/MW/MX'' +
				''/NA/NB/NE/NH/NL/NM/NP/NR/NS/NW/NX/NY/NZ'' +
				''/OA/OB/OE/OH/OK/OL/OM/OP/OR/OS/OX'' +
				''/PA/PB/PC/PE/PG/PH/PJ/PK/PL/PM/PN/PP/PR/PS/PT/PW/PX/PY'' +
				''/RA/RB/RE/RH/RK/RM/RP/RR/RS/RT/RW/RX/RY/RZ'' +
				''/SA/SB/SC/SE/SG/SH/SJ/SK/SL/SM/SN/SP/SR/SS/ST/SW/SX/SY/SZ'' +
				''/TA/TB/TE/TH/TK/TL/TM/TP/TR/TS/TT/TW/TX/TY/TZ'' +
				''/WA/WB/WE/WK/WL/WM/WP'' +
				''/YA/YB/YE/YH/YK/YL/YM/YP/YR/YS/YT/YW/YX/YY/YZ'' +
				''/ZA/ZB/ZE/ZH/ZK/ZL/ZM/ZP/ZR/ZS/ZT/ZW/ZX/ZY/'';

			SET @ValidSuffixes = ''/ /A/B/C/D/'';

			SET @Prefix = ''/''+left(@input+''  '',2)+''/''
			SET @Suffix = ''/''+substring(@input+'' '',9,1)+''/''
			SET @Numerics = SUBSTRING(@input,3,6)

			IF charindex(@Prefix,@ValidPrefixes) = 0 OR charindex(@Suffix,@ValidSuffixes) = 0 OR ISNUMERIC(@Numerics) = 0
				SET @result = 0;
				
			RETURN @result;
			
		END';
	EXECUTE sp_executeSQL @sSPCode;


	SET @sSPCode = 'CREATE FUNCTION [dbo].[udfsys_isvalidpayrollcharacterset](
			@input AS varchar(MAX),
			@charset varchar(1))
		RETURNS bit
		WITH SCHEMABINDING
		AS
		BEGIN
		
			DECLARE @result bit;

			--Charset A - typically Address
			--Charset C - typically Forename
			--Charset D - typically Surname

			DECLARE @ValidCharacters varchar(MAX),
					@Index int;

			IF      @Charset = ''A'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''0123456789,&/(). =!"%&*;<>+:?''
			ELSE IF @Charset = ''B'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ''
			ELSE IF @Charset = ''C'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''''
			ELSE IF @Charset = ''D'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''0123456789,&/(). ''
			ELSE IF @Charset = ''G'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''0123456789,&/(). =!"%&*;<>+:?''
			ELSE IF @Charset = ''H'' SET @ValidCharacters = ''abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''''. ''
			
			SET @result = 1;
			SET @Index = 1;
			WHILE (@Index <= datalength(@input) AND @result = 1)
			BEGIN
				IF charindex(substring(@input,@Index,1),@ValidCharacters) = 0
					SET @result = 0;
				SET @Index = @Index + 1;
			END	

			RETURN @result;

		END'
	EXECUTE sp_executeSQL @sSPCode;



/* ------------------------------------------------------------- */
PRINT 'Step 12 - Populate code generation tables'

	EXEC sp_executesql N'IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[tbstat_componentcode]'') AND type in (N''U''))
		DROP TABLE [dbo].[tbstat_componentcode]'

	EXEC sp_executesql N'IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[tbstat_componentdependancy]'') AND type in (N''U''))
		DROP TABLE [dbo].[tbstat_componentdependancy]'

	EXEC sp_executesql N'CREATE TABLE [dbo].[tbstat_componentcode](
			[id] [int] NOT NULL,
			[objectid] [uniqueidentifier] NOT NULL,
			[code] [nvarchar](max) NULL,
			[precode] [nvarchar](MAX) NULL,
			[aftercode] [nvarchar](50) NULL,
			[datatype] [int] NULL,
			[name] [nvarchar](255) NOT NULL,
			[isoperator] [bit] NULL,
			[operatortype] [tinyint] NULL,
			[rownumberrequired] [bit] NULL,
			[calculatepostaudit] [bit] NULL
		) ON [PRIMARY]';

	EXEC sp_executesql N'CREATE TABLE [dbo].[tbstat_componentdependancy](
			[id] [integer] NOT NULL,
			[type] [integer] NOT NULL,
			[modulekey] [nvarchar](50) NOT NULL,
			[parameterkey] [nvarchar](50) NOT NULL,
			[code] nvarchar(MAX) NOT NULL
		) ON [PRIMARY]';

	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''4ec6c760-2157-492d-9161-24aa7c8a7b35'', N''AND'', NULL, N''And'', NULL, 1, 178, 5)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''8ef94e11-6693-422d-8099-bedee430083a'', N''+'', NULL, N''Concatenated with'', NULL, 1, 0, 17)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [precode], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''57bc755b-61b5-41a0-92a0-321aab134b9c'', N'','', NULL, N''Is Contained Within'', N''(CHARINDEX('', N'')>0)'', 1, 177, 14)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [precode], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''a34f7387-91a1-40d6-b42f-f8032609cfd6'', N'','', NULL, N''Divided by'', N''dbo.[udfsys_divide]('', N'')'', 1, 0, 4)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''d4521a9e-2974-49ef-849c-0d132aca93a0'', N''='', NULL, N''Is equal to'', NULL, 1, 177, 7)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''14b67bc6-ab84-4bf5-b20c-40c16e94a193'', N''>'', NULL, N''Is greater than'', NULL, 1, 177, 10)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''f54c9ae5-4790-403b-bb66-a026d67df26e'', N''>='', NULL, N''Is greater than OR equal to'', NULL, 1, 177, 12)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''11f18863-ecbd-4930-ab00-544a4fba5162'', N''<'', NULL, N''Is less than'', NULL, 1, 177, 9)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''14dd3e78-331a-47f6-81d2-5ce5df8c6935'', N''<='', NULL, N''Is less than OR equal to'', NULL, 1, 177, 11)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''3543f5d9-eef6-48c7-8aa4-934fe4202700'', N''<>'', NULL, N''Is NOT equal to'', NULL, 1, 177, 8)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''435776a4-6803-4a08-972b-c40480313ce8'', N''-'', NULL, N''Minus'', NULL, 1, 0, 2)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''b35d6bba-d45e-4ec4-bcd0-1d3d3e2d78fc'', N''%'', NULL, N''Modulus'', NULL, 1, 0, 16)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''68a326f9-ca7f-496f-b6e1-0d0f488ac7f6'', N''OR'', NULL, N''Or'', NULL, 1, 178, 6)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''6e51716a-4ac3-49dc-97a5-2bc417e38c2f'', N''+'', NULL, N''Plus'', NULL, 1, 0, 1)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''1acde45c-39a1-4a50-8526-aed3b8e6392b'', N''*'', NULL, N''Times by'', NULL, 1, 0, 3)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''a0aefbd0-b295-4598-9432-d4f653eca1ac'', N''*'', NULL, N''To the power of'', NULL, 1, 0, 15)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''e6bd0161-786d-42a8-bdff-8400963e3e89'', N''[dbo].[udf_ASRFn_AbsenceBetweenTwoDates] ({0}, {1}, {2}, {3}, GETDATE())'', 2, N''Absence between Two Dates'', NULL, 0, 0, 47)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (47, 1, ''MODULE_PERSONNEL'', ''Param_TablePersonnel'', '''')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''a4776d94-8917-4f5b-ad36-f4104b04e3e0'', N''[dbo].[udf_ASRFn_AbsenceDuration]({0}, {1}, {2}, {3}, {4})'', 2, N''Absence Duration'', NULL, 0, 0, 30)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (30, 1, ''MODULE_PERSONNEL'', ''Param_TablePersonnel'', '''')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''edfa5940-f5ba-47b5-bd93-8f19c35490b3'', N''DATEADD(DD, {1}, DATEADD(D, 0, DATEDIFF(D, 0, {0})))'', 4, N''Add Days to Date'', NULL, 0, 0, 44)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''51a4dc3e-41a4-4b1a-8df8-8d9a1baed196'', N''DATEADD(MM, {1}, DATEADD(D, 0, DATEDIFF(D, 0, {0})))'', 4, N''Add Months to Date'', NULL, 0, 0, 23)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''bc6a9215-696d-492c-8acb-95c99f440530'', N''DATEADD(YY, {1}, DATEADD(D, 0, DATEDIFF(D, 0, {0})))'', 4, N''Add Years to Date'', NULL, 0, 0, 24)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''078108bf-77b2-42a3-b426-42126337f397'', N''[dbo].[udf_ASRFn_BradfordFactor]({0}, {1}, {2}, {3})'', 2, N''Bradford Factor'', NULL, 0, 0, 73)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (73, 1, ''MODULE_PERSONNEL'', ''Param_TablePersonnel'', '''')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''eb449e75-e061-4502-973b-5e3a3e39c2d2'', N''dbo.[udfsys_convertcharactertonumeric]({0})'', 2, N''Convert Character to Numeric'', NULL, 0, 0, 25)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''56b64c0d-84d9-4b15-9c9e-b1fdb42ea4d1'', N''dbo.[udfsys_convertcurrency]({0},{1},{2})'', 2, N''Convert Currency'', NULL, 0, 0, 51)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''88430aa0-f580-4157-8b2f-c73841cea211'', N''LTRIM(STR({0}, 20, convert(integer,{1})))'', 1, N''Convert Numeric to Character'', NULL, 0, 0, 3)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''98e87fe4-bb86-4382-bf53-40fa1275d677'', N''LOWER({0})'', 1, N''Convert to Lowercase'', NULL, 0, 0, 8)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''9e055f03-efe9-4c47-a528-85cd3c57c12a'', N''[dbo].[udfsys_propercase]({0})'', 1, N''Convert to Proper Case'', NULL, 0, 0, 12)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''59a5f6dd-8284-45a2-a68e-01e9f6d2e13e'', N''UPPER({0})'', 1, N''Convert to Uppercase'', NULL, 0, 0, 2)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''302dbbe5-d900-4547-8090-5de3dd3a4970'', N''SYSTEM_USER'', 1, N''Current User'', NULL, 0, 0, 17)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''8a4abce8-984e-4d4f-b1ca-aaef09e1c08d'', N''DATEPART(day, {0})'', 2, N''Day of Date'', NULL, 0, 0, 34)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''b41669c9-59d7-449f-be4f-6d4c6b809db9'', N''DATEPART(weekday, {0})'', 2, N''Day of the Week'', NULL, 0, 0, 28)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''24884a1c-fc85-4bba-8752-cb594c4607f2'', N''(DATEDIFF(dd,{0}, {1})+1)'', 2, N''Days between Two Dates'', NULL, 0, 0, 45)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''25033092-aa37-406d-ba0e-7b59b81c9b69'', N'''', 3, N''Does Record Exist'', NULL, 0, 0, 74)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''a774b4f7-5792-41c5-99fb-301af38f0e68'', N''LEFT({0}, {1})'', 1, N''Extract Characters from the Left'', NULL, 0, 0, 6)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''0d948a6a-e6db-440f-b5fc-25ac323425ae'', N''RIGHT({0}, {1})'', 1, N''Extract Characters from the Right'', NULL, 0, 0, 13)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5c4e830d-6b52-481d-b94e-e6d65912cde2'', N''SUBSTRING({0}, {1}, {2})'', 1, N''Extract Part of a Character String'', NULL, 0, 0, 14)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id], [calculatepostaudit]) VALUES (N''f61ea313-4866-4a29-a19f-e2d4fe3db23d'', N''dbo.udfsys_fieldchangedbetweentwodates({0}, {1}, {2}, {3})'', 3, N''Field Changed between Two Dates'', NULL, 0, 0, 53, 1)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (53, 2, '''', '''', ''@prm_ID'')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id], [calculatepostaudit]) VALUES (N''532861e4-23ac-474b-ae04-1a85724e7988'', N''dbo.udfsys_fieldlastchangedate({0}, {1})'', 4, N''Field Last Change Date'', NULL, 0, 0, 52, 1)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (52, 2, '''', '''', ''@prm_ID'')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''4be2a715-f36b-4507-8090-9b1159de3aab'', N''DATEADD(dd, 1 - DATEPART(dd,{0}), {0})'', 2, N''First Day of Month'', NULL, 0, 0, 55)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''e3f98ac8-bfbf-4a98-8dd3-89f2830c1c95'', N''DATEADD(dd, 1 - DATEPART(dy, {0}), {0})'', 2, N''First Day of Year'', NULL, 0, 0, 57)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''263f4cc8-7c8d-4c5d-bdea-9e4ced21f078'', N''[dbo].[udfsys_firstnamefromforenames]({0})'', 1, N''First Name from Forenames'', NULL, 0, 0, 21)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''f14ebf8d-98e6-4e36-a1e1-35efd0023c55'', N''CASE WHEN ({0}) = 1 THEN {1} ELSE {2} END'', 0, N''If... Then... Else...'', NULL, 0, 0, 4)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5feedcc3-e731-46b0-b7fe-2027e1e9ded4'', N''[dbo].[udfsys_initialsfromforenames]({0},0)'', 1, N''Initials from Forenames'', NULL, 0, 0, 20)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''7d539e37-6d9f-44b3-a694-7db9638a2502'', N''dbo.[udfsys_isbetween]({0},{1},{2})'', 0, N''Is Between'', NULL, 0, 0, 38)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''a9997816-add0-467f-999d-79ef30c2b713'', N''[dbo].[udfsys_isfieldempty]({0})'', 3, N''Is Field Empty'', NULL, 0, 0, 16)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''8caf9f74-dee4-4618-8d59-e292847f202a'', N''[dbo].[udfsys_isfieldpopulated]({0})'', 3, N''Is Field Populated'', NULL, 0, 0, 61)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''63d90dd1-1fb0-42a7-8135-83cb25293d7b'', N''[dbo].[udfsys_isovernightprocess]() '', 3, N''Is Overnight Process'', NULL, 0, 0, 50)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''94127e4f-8046-4516-83a0-2062dd0ea2e6'', N'''', 3, N''Is Personnel That Current User Reports To'', NULL, 0, 0, 72)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''17d67659-4e60-40ee-bb72-763f4f85a645'', N'''', 3, N''Is Personnel That Reports To Current User'', NULL, 0, 0, 68)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''0a0d63a7-d926-4b8c-9f4e-2c3ae3d650ab'', N'''', 3, N''Is Post That Current User Reports To'', NULL, 0, 0, 70)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''cb6680b4-1940-435d-8144-bae2af8f37a1'', N'''', 3, N''Is Post That Reports To Current User'', NULL, 0, 0, 66)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''3e9cae1a-0948-481d-8d0b-9c13ca5d9373'', N''DATEADD(dd, -1, DATEADD(mm, 1, DATEADD(dd, 1 - DATEPART(dd, {0}), {0})))'', 4, N''Last Day of Month'', NULL, 0, 0, 56)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''bc2ce0fc-6e2c-43a2-86f9-0ed45cba129a'', N''DATEADD(dd, -1, DATEADD(yy, 1, DATEADD(dd, 1 - DATEPART(dy, {0}), {0})))'', 4, N''Last Day of Year'', NULL, 0, 0, 58)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''18babc7b-b84e-4ca9-9e10-c630bb004891'', N''LEN({0})'', 2, N''Length of Character Field'', NULL, 0, 0, 7)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''bba7fff7-bd75-4953-abd1-2f70418bbb80'', N''[dbo].[udfsys_maternityexpectedreturndate](@prm_ID)'', 4, N''Maternity Expected Return Date'', NULL, 0, 0, 64)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5acc9ebe-af46-438e-9ebd-2741b42e26e0'', N''CASE WHEN {0} > {1} THEN {0} ELSE {1} END'', 1, N''Maximum'', NULL, 0, 0, 9)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''e03d8884-8835-425c-b268-1eec196917eb'', N''CASE WHEN {0} < {1} THEN {0} ELSE {1} END'', 1, N''Minimum'', NULL, 0, 0, 10)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''7d1376aa-5bdc-4844-9e5d-b3499b807639'', N''DATEPART(MM, {0})'', 2, N''Month of Date'', NULL, 0, 0, 33)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''fbff11aa-2aa9-43c9-b75f-5f2333ff880e'', N''DATENAME(weekday, {0})'', 1, N''Name of Day'', NULL, 0, 0, 60)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''d97bb954-d303-4eeb-90ce-1466287de905'', N''DATENAME(month, {0})'', 1, N''Name of Month'', NULL, 0, 0, 59)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''ccbcd03a-0c7e-47c7-b4bf-d6e8bd7963e8'', N''[dbo].[udfsys_nicedate]({0})'', 1, N''Nice Date'', NULL, 0, 0, 35)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''42a88b07-200f-4785-9c1f-e4b5a97a9001'', N''[dbo].[udfsys_nicetime]({0})'', 1, N''Nice Time'', NULL, 0, 0, 36)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''e59b8c9c-31d1-494f-b9c3-ca0a6a6aef1e'', N''(convert(float, len(replace(left({0}, 14), SPACE(1), SPACE(0)))) / 2)'', 2, N''Number of Working Days per Week'', NULL, 0, 0, 29)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''84044568-fea7-48d5-ae8a-f8178b7ed927'', N''[dbo].[udfsys_parentalleaveentitlement](@prm_ID)'', 2, N''Parental Leave Entitlement'', NULL, 0, 0, 62)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5278a126-c44e-41c5-9e7a-1c890c297d3f'', N''[dbo].[udfsys_parentalleavetaken](@prm_ID)'', 2, N''Parental Leave Taken'', NULL, 0, 0, 63)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''ed3be9d9-28f1-4345-a8c8-ca9f0c18a3a2'', N''({0})'', 0, N''Parentheses'', NULL, 0, 0, 27)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''06e67c1b-c376-4fc9-a260-e9a12022791f'', N''[dbo].[udfsys_remainingmonthssincewholeyears]({0}, GETDATE())'', 2, N''Remaining Months since Whole Years'', NULL, 0, 0, 19)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''b86c77e6-e393-499e-9114-95a201a316d4'', N''LTRIM(RTRIM({0}))'', 1, N''Remove Leading and Trailing Spaces'', NULL, 0, 0, 5)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5c2244b5-ee8b-4f80-bc9e-defb9ba10b36'', N''[dbo].[udfsys_roundtostartofnearestmonth]({0})'', 4, N''Round Date to Start of Nearest Month'', NULL, 0, 0, 37)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''07e1acb6-6943-4a92-956f-5df24aa2f3d2'', N''FLOOR({0})'', 2, N''Round Down to Nearest Whole Number'', NULL, 0, 0, 31)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''6c8c5ca0-ae52-46fc-9289-06e989c32d6d'', N''ROUND({0} / {1}, 0) * {1}'', 2, N''Round to Nearest Number'', NULL, 0, 0, 49)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''49161588-d050-4f0b-a0cd-3d9d6393f5f3'', N''CEILING({0})'', 2, N''Round Up to Nearest Whole Number'', NULL, 0, 0, 48)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''022a1f4c-b15b-411a-a49f-08ec4c3497e4'', N''CHARINDEX ({1}, {0}, 0) '', 1, N''Search for Character String'', NULL, 0, 0, 11)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''cfa37a8b-4d7b-4abc-ae80-7866219e4469'', N''[dbo].[udfsys_servicelength]({0}, {1}, ''''M'''')'', 2, N''Service Months'', NULL, 0, 0, 40)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''81847039-a90d-476c-88a5-c5e447d77701'', N''[dbo].[udfsys_servicelength]({0}, {1}, ''''Y'''')'', 2, N''Service Years'', NULL, 0, 0, 39)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''2bf404b7-970e-4fdb-9977-00d516a6cc84'', N''[dbo].[udfsys_statutoryredundancypay]({0}, {1}, {2}, {3}, {4})'', 2, N''Statutory Redundancy Pay'', NULL, 0, 0, 41)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''1cce61bf-ee36-4779-83b9-233885440437'', N''(DATEADD(D, 0, DATEDIFF(D, 0, GETDATE())))'', 4, N''System Date'', NULL, 0, 0, 1)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''1b77e32f-756b-4e97-94d2-f0b053b0baca'', N''CONVERT(varchar,GETDATE(),8)'', 1, N''System Time'', NULL, 0, 0, 15)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id], [rownumberrequired]) VALUES (N''a8974869-0964-40e9-bbbf-4ac6157bf07f'', N''[dbo].[udfsys_getuniquecode] ({0}, {1}, {2})'', 0, N''Unique Code'', NULL, 0, 0, 43, 1)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (43, 2, '''', '''', ''@rownumber'')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''09e7dfb0-3bc2-4db5-a596-9639eb3e77b5'', N''[dbo].[udfsys_weekdaysbetweentwodates] ({0}, {1})'', 2, N''Weekdays between Two Dates'', NULL, 0, 0, 22)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''fbccef52-27be-4ee4-8afa-d8228da2e952'', N''[dbo].[udfsys_wholemonthsbetweentwodates] ({0}, {1})'', 2, N''Whole Months between Two Dates'', NULL, 0, 0, 26)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''1b5082ad-36bb-4bf8-b859-22a1de8f8d2e'', N''[dbo].[udfsys_wholeyearsbetweentwodates] ({0}, {1})'', 2, N''Whole Years between Two Dates'', NULL, 0, 0, 54)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''97880cd2-c73d-4c7e-a4c4-971824b850e6'', N''[dbo].[udfsys_wholeyearsbetweentwodates] ({0}, GETDATE())'', 2, N''Whole Years until Current Date'', NULL, 0, 0, 18)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''f11ffb85-31bc-4b12-9b3d-e4464c868ca4'', N''[dbo].[udf_ASRFn_WorkingDaysBetweenTwoDates] ({0}, {1}, {2})'', 2, N''Working Days between Two Dates'', NULL, 0, 0, 46)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentdependancy] ([id], [type], [modulekey], [parameterkey], [code]) VALUES (46, 1, ''MODULE_PERSONNEL'', ''Param_TablePersonnel'', '''')';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5c9a6256-ac11-456d-92fe-a5e2f5ba4c11'', N''DATEPART(YYYY, {0})'', 2, N''Year of Date'', NULL, 0, 0, 32)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5b636d9f-7589-46d4-bd6a-0e23aef81a51'', N''NOT'', 0, N''Not'', NULL, 1, 180, 13)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5da8bb7e-f632-4ed0-b236-e042b88f3a1b'', N''[dbo].[udfsys_getfieldfromdatabaserecord] ({0}, {1}, {2})'', 0, N''Get field from database record'', NULL, 0, 0, 42)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''a40b59a0-3b3c-4348-9e6b-dd56a8dbab86'', N''[dbo].[udfsys_isnivalid]({0})'', 3, N''Is Valid NI Number'', NULL, 0, 0, 75)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''5b2c16e0-489a-4061-a0eb-eb94d5f2ee6f'', N''[dbo].[udfsys_isvalidpayrollcharacterset]({0}, {1})'', 3, N''Is Valid Payroll Character Set'', NULL, 0, 0, 76)';
	EXEC sp_executesql N'INSERT [dbo].[tbstat_componentcode] ([objectid], [code], [datatype], [name], [aftercode], [isoperator], [operatortype], [id]) VALUES (N''84b11964-4b8b-46e4-b340-f3d5598a82fe'', N''REPLACE({0}, {1}, {2})'', 1, N''Replace Characters In A String'', NULL, 0, 0, 77)';



/* ------------------------------------------------------------- */
PRINT 'Step 13 - Administration module stored procedures'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_getcomponentcode]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_getcomponentcode]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_getcomponentcodedependancies]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_getcomponentcodedependancies]

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spadmin_getcomponentcode]
	AS
	BEGIN
		SELECT [id], [code], [name], ISNULL([datatype],0) AS [returntype]
			, [precode], [aftercode], [isoperator], [operatortype], [aftercode] 
			, ISNULL([rownumberrequired], 0) AS [rownumberrequired]
			, ISNULL([CalculatePostAudit],0) AS [calculatepostaudit]
			FROM tbstat_componentcode WHERE [id] IS NOT NULL;
	END';

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spadmin_getcomponentcodedependancies]
		(@componentid integer)
	AS
	BEGIN
		SELECT c.type, c.modulekey, c.parameterkey, m.parametervalue AS value,
				CASE m.parametertype
   					WHEN ''PType_ColumnID'' THEN 2
					WHEN ''PType_EmailID'' THEN 18
					WHEN ''PType_Encrypted'' THEN 0
					WHEN ''PType_Other'' THEN 0
					WHEN ''PType_ScreenID'' THEN 14
					WHEN ''PType_TableID'' THEN 1
				END AS [settingtype],
				c.[code] AS [code]
			FROM tbstat_componentdependancy c 
			LEFT JOIN ASRSysModulesetup m on m.modulekey = c.modulekey AND m.parameterkey = c.parameterkey
			WHERE c.id = @componentid;
	END';


/* ------------------------------------------------------------- */
PRINT 'Step 15 - Remove redundant procedures'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRAudit]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRAudit];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRFn_RemainingMonthsSinceWholeYears]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRFn_RemainingMonthsSinceWholeYears];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRFn_RoundDateToStartOfNearestMonth]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRFn_RoundDateToStartOfNearestMonth];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRFn_WeekdaysFromStartAndEndDates]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRFn_WeekdaysFromStartAndEndDates];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSysFnMaternityExpectedReturn]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spASRSysFnMaternityExpectedReturn]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSysFnParentalLeaveEntitlement]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spASRSysFnParentalLeaveEntitlement]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSysFnParentalLeaveTaken]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spASRSysFnParentalLeaveTaken]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRMaternityExpectedReturn]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spASRMaternityExpectedReturn]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRParentalLeaveEntitlement]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spASRParentalLeaveEntitlement]



/* ------------------------------------------------------------- */
PRINT 'Step 16 - Server settings'

	EXEC sp_dboption @DBName, 'recursive triggers', 'FALSE';



/* ------------------------------------------------------------- */
PRINT 'Step 17 - Trigger functionality'

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_intransactiontrigger]') AND type in (N'U'))
	DROP TABLE [dbo].[tbsys_intransactiontrigger]

	EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_intransactiontrigger](
		[spid] [integer] NOT NULL,
		[tablefromid] [integer] NOT NULL,
		[actiontype] [tinyint] NULL)'


/* ------------------------------------------------------------- */
PRINT 'Step 18 - System Functions'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_generateuniquecodes]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spadmin_generateuniquecodes]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_generategetfields]') AND xtype = 'P')
		DROP PROCEDURE dbo.[spadmin_generategetfields]

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[tbsys_uniquecodes]') AND xtype = 'U')
		EXECUTE sp_executesql N'EXECUTE sp_rename [ASRSysUniqueCodes], [tbsys_uniquecodes];';

	EXECUTE sp_executesql N'CREATE PROCEDURE dbo.[spadmin_generateuniquecodes]
		AS
		BEGIN
			DECLARE @sSQL nvarchar(MAX);

			SET @sSQL = '''';

			SELECT @sSQL = ISNULL(@sSQL, '''') +
					+ '' UNION SELECT '''''' + c2.ValueCharacter + '''''' AS [codeprefix], ''
					+ '' MAX(['' + cl.columnname + '']) AS [maxcodesuffix], 1 AS [isbound] FROM dbo.[tbuser_'' +  t.tablename + '']''
				FROM ASRSysExprComponents c
					INNER JOIN ASRSysExpressions e ON c.ComponentID = e.ParentComponentID
					INNER JOIN ASRSysExprComponents c2 ON e.ExprID = c2.ExprID AND c2.ValueCharacter IS NOT NULL
					INNER JOIN ASRSysExpressions p ON p.ExprID = c.ExprID
					INNER JOIN ASRSysColumns cl ON cl.calcexprid = ISNULL(dbo.udf_ASRRootExpressionID(p.ParentComponentID), p.exprid)
					INNER JOIN ASRSysTables t ON cl.tableid = t.tableid
				WHERE c.FunctionID = 43;

				SET @sSQL = ''CREATE VIEW dbo.[ASRSysUniqueCodes] AS SELECT ''
					+ SUBSTRING(@ssql,15,DATALENGTH(@ssql));

				EXEC sp_executesql N''IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''''[dbo].[ASRSysUniqueCodes]'''') AND type in (N''''V''''))
					DROP VIEW [dbo].[ASRSysUniqueCodes]'';

				EXEC sp_executesql @sSQL;

		END'

	EXECUTE sp_executesql N'CREATE PROCEDURE dbo.[spadmin_generategetfields]
		AS
		BEGIN

			EXEC sp_executesql N''IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''''[dbo].[ASRSysGetFieldFromDatabaseRecord]'''') AND type in (N''''V''''))
				DROP VIEW [dbo].[ASRSysGetFieldFromDatabaseRecord]'';

			EXEC sp_executesql N''CREATE VIEW dbo.[ASRSysGetFieldFromDatabaseRecord] AS SELECT '''' AS [searchcolumnid]
				,'''' AS [returnfieldid]
				, '''' AS [lookupkey], '''' AS [returnvalue]'';
		
		END'

	EXECUTE sp_executesql N'spadmin_generategetfields';
	EXECUTE sp_executesql N'spadmin_generateuniquecodes';
	
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
		IF @sObject LIKE 'sp_ASRExpr_%' OR @sObject LIKE 'sp_ASRDfltExpr_%' OR @sObject LIKE 'spASREmail_%' OR @sObject LIKE 'spASRUpdateOLEField_%'
	        SET @NVarCommand = 'REVOKE EXECUTE ON [' + @sObject + '] TO [ASRSysGroup]'
		ELSE		    
			SET @NVarCommand = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
    END
    ELSE
    BEGIN
		SET @NVarCommand = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
    END

    EXECUTE sp_executeSQL @NVarCommand


    FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '4.3';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '4.3.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '4.3.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '4.2.0';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v4.3')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.3 Of HR Pro'
