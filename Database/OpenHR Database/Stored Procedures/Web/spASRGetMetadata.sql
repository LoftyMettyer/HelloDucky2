﻿CREATE PROCEDURE [dbo].[spASRGetMetadata] (@Username varchar(255))
WITH ENCRYPTION
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @licenseKey			varchar(MAX);

	EXEC [dbo].[sp_ASRIntGetSystemSetting] 'Licence', 'Key', 'moduleCode', @licenseKey OUTPUT, 0, 0;

	SELECT TableID, TableName, TableType, DefaultOrderID, RecordDescExprID FROM dbo.ASRSysTables;

	SELECT ColumnID, TableID, ColumnName, DataType, ColumnType, Use1000Separator, Size, Decimals, LookupTableID, LookupColumnID FROM dbo.ASRSysColumns;

	SELECT ParentID, ChildID FROM dbo.ASRSysRelations;

	SELECT ModuleKey, ParameterKey, ISNULL(ParameterValue,'') AS ParameterValue, ParameterType FROM dbo.ASRSysModuleSetup;

	SELECT * FROM dbo.ASRSysUserSettings WHERE Username = @Username;

	SELECT functionID, functionName, returnType FROM dbo.ASRSysFunctions;

	SELECT * FROM dbo.ASRSysFunctionParameters ORDER BY functionID, parameterIndex;

	SELECT * FROM dbo.ASRSysOperators;

	SELECT * FROM dbo.ASRSysOperatorParameters ORDER BY OperatorID, parameterIndex;
	
	-- Selected system settings
	SELECT * FROM ASRSysSystemSettings;


END

