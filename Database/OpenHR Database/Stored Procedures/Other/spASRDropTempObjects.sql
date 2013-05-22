CREATE PROCEDURE [dbo].[spASRDropTempObjects]
AS
BEGIN

	DECLARE	@sObjectName varchar(255),
			@sUsername varchar(255),
			@sXType varchar(50);
				
	DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
	SELECT [dbo].[sysobjects].[name], [sys].[schemas].[name], [dbo].[sysobjects].[xtype]
	FROM [dbo].[sysobjects] 
			INNER JOIN [sys].[schemas]
			ON [dbo].[sysobjects].[uid] = [sys].[schemas].[schema_id]
	WHERE LOWER([sys].[schemas].[name]) != 'dbo' AND LOWER([sys].[schemas].[name]) != 'messagebus'
			AND (OBJECTPROPERTY(id, N'IsUserTable') = 1
				OR OBJECTPROPERTY(id, N'IsProcedure') = 1
				OR OBJECTPROPERTY(id, N'IsInlineFunction') = 1
				OR OBJECTPROPERTY(id, N'IsScalarFunction') = 1
				OR OBJECTPROPERTY(id, N'IsTableFunction') = 1);

	OPEN tempObjects;
	FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType;
	WHILE (@@fetch_status <> -1)
	BEGIN		
		IF UPPER(@sXType) = 'U'
			-- user table
			BEGIN
				EXEC ('DROP TABLE [' + @sUsername + '].[' + @sObjectName + ']');
			END

		IF UPPER(@sXType) = 'P'
			-- procedure
			BEGIN
				EXEC ('DROP PROCEDURE [' + @sUsername + '].[' + @sObjectName + ']');
			END

		IF UPPER(@sXType) = 'TF'
			-- UDF
			BEGIN
				EXEC ('DROP FUNCTION [' + @sUsername + '].[' + @sObjectName + ']');
			END

		IF UPPER(@sXType) = 'FN'
			-- UDF
			BEGIN
				EXEC ('DROP FUNCTION [' + @sUsername + '].[' + @sObjectName + ']');
			END
		
		FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType;
		
	END
	CLOSE tempObjects;
	DEALLOCATE tempObjects;
	
	EXEC ('DELETE FROM [dbo].[ASRSysSQLObjects]');


	-- Clear out any temporary tables that may have got left behind from the createunique function
	DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
	SELECT [dbo].[sysobjects].[name]
	FROM [dbo].[sysobjects] 
	INNER JOIN [dbo].[sysusers]	ON [dbo].[sysobjects].[uid] = [dbo].[sysusers].[uid]
	LEFT JOIN ASRSysTables ON sysobjects.[name] = ASRSysTables.TableName
	WHERE LOWER([dbo].[sysusers].[name]) = 'dbo'
		AND OBJECTPROPERTY(sysobjects.id, N'IsUserTable') = 1
		AND ASRSysTables.TableName IS NULL
		AND [dbo].[sysobjects].[name] LIKE 'tmp%';

	OPEN tempObjects;
	FETCH NEXT FROM tempObjects INTO @sObjectName;
	WHILE (@@fetch_status <> -1)
	BEGIN		
		EXEC ('DROP TABLE [dbo].[' + @sObjectName + ']');
		FETCH NEXT FROM tempObjects INTO @sObjectName;
	END

	CLOSE tempObjects;
	DEALLOCATE tempObjects;

END
