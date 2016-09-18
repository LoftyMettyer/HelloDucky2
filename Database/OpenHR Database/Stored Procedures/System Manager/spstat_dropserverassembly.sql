CREATE PROCEDURE [dbo].[spstat_dropserverassembly] (@assemblyname nvarchar(255))
	AS
	BEGIN

	  IF EXISTS (SELECT name FROM sys.assemblies WHERE name = @assemblyname)
	  BEGIN

		  DECLARE @error int
		  SET @error = 0

		  -- Drop the assembly user defined aggregates, triggers, functions and procedures
		  DECLARE @moduleId sysname
		  DECLARE @moduleName sysname
		  DECLARE @moduleType char(2)
		  DECLARE @moduleClass tinyint

	  DECLARE assemblyModules CURSOR FAST_FORWARD FOR
		SELECT t.object_id, t.name, t.type, t.parent_class as class
		  FROM sys.triggers t
		  INNER JOIN sys.assembly_modules m ON t.object_id = m.object_id
		  INNER JOIN sys.assemblies a ON m.assembly_id = a.assembly_id
		  WHERE a.Name = @assemblyname
		UNION
		SELECT o.object_id, o.name, o.type, NULL as class
		  FROM sys.objects o
		  INNER JOIN sys.assembly_modules m ON o.object_id = m.object_id
		  INNER JOIN sys.assemblies a ON m.assembly_id = a.assembly_id
		  WHERE a.Name = @assemblyname
	  OPEN assemblyModules
	  FETCH NEXT FROM assemblyModules INTO @moduleId, @moduleName, @moduleType, @moduleClass
	  WHILE (@error = 0 AND @@FETCH_STATUS = 0)
	  BEGIN
		DECLARE @dropModuleString nvarchar(256)
		IF (@moduleType = 'AF') SET @dropModuleString = N'AGGREGATE'
		IF (@moduleType = 'TA') SET @dropModuleString = N'TRIGGER'
		IF (@moduleType = 'FT' OR @moduleType = 'FS') SET @dropModuleString = N'FUNCTION'
		IF (@moduleType = 'PC') SET @dropModuleString = N'PROCEDURE'
			SET @dropModuleString = N'DROP ' + @dropModuleString + ' [' + REPLACE(@moduleName, ']', ']]') + ']'
		IF (@moduleType = 'TA' AND @moduleClass = 0)
		BEGIN
		  SET @dropModuleString = @dropModuleString + N' ON DATABASE'
		END
		EXEC sp_executesql @dropModuleString
		FETCH NEXT FROM assemblyModules INTO @moduleId, @moduleName, @moduleType, @moduleClass
	  END
	  CLOSE assemblyModules
	  DEALLOCATE assemblyModules
	  
	  -- Drop the assembly user defined types
	  DECLARE @typeId int
	  DECLARE @typeName sysname
	  DECLARE assemblyTypes CURSOR FAST_FORWARD
		FOR SELECT t.user_type_id, t.name
		  FROM sys.assembly_types t
		  INNER JOIN sys.assemblies a ON t.assembly_id = a.assembly_id
		  WHERE a.Name = @assemblyname
	  OPEN assemblyTypes
	  FETCH NEXT FROM assemblyTypes INTO @typeId, @typeName
	  WHILE (@error = 0 AND @@FETCH_STATUS = 0)
	  BEGIN
		DECLARE @dropTypeString nvarchar(256)
		SET @dropTypeString = N'DROP TYPE [' + REPLACE(@typeName, ']', ']]') + ']'
		IF NOT EXISTS (SELECT name FROM sys.extended_properties WHERE major_id = @typeId AND name = 'AutoDeployed')
		BEGIN
		  DECLARE @quotedTypeName sysname
		  SET @quotedTypeName = REPLACE(@typeName, '''', '''''')
		  RAISERROR(N'The assembly user defined type ''%s'' cannot be preserved because it was not automatically deployed.', 16, 1,@quotedTypeName)
		  SET @error = @@ERROR
		END
		ELSE
		BEGIN
		  EXEC sp_executesql @dropTypeString
		  FETCH NEXT FROM assemblyTypes INTO @typeId, @typeName
		END
	  END
	  CLOSE assemblyTypes
	  DEALLOCATE assemblyTypes

	  -- Drop the assembly
	  IF (@error = 0)
	  
		SET @dropModuleString = 'IF EXISTS (SELECT name FROM sys.assemblies WHERE name = N''' + @assemblyname + ''')
				DROP ASSEMBLY [' +  @assemblyname + '] WITH NO DEPENDENTS;';
		EXEC sp_executesql @dropModuleString

	  END
	END