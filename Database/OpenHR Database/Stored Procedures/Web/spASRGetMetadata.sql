CREATE PROCEDURE [dbo].[spASRGetMetadata] (@Username varchar(255))
WITH ENCRYPTION
AS
BEGIN

	DECLARE @licenseKey			varchar(MAX);

	EXEC [dbo].[sp_ASRIntGetSystemSetting] 'Licence', 'Key', 'moduleCode', @licenseKey OUTPUT, 0, 0;


	SELECT TableID, TableName, TableType, DefaultOrderID, RecordDescExprID FROM dbo.ASRSysTables;

	SELECT ColumnID, TableID, ColumnName, DataType, ColumnType, Use1000Separator, Size, Decimals FROM dbo.ASRSysColumns;

	SELECT ParentID, ChildID FROM dbo.ASRSysRelations;

	SELECT ModuleKey, ParameterKey, ISNULL(ParameterValue,'') AS ParameterValue, ParameterType FROM dbo.ASRSysModuleSetup;

	SELECT * FROM dbo.ASRSysUserSettings WHERE Username = @Username;

	SELECT functionID, functionName, returnType FROM dbo.ASRSysFunctions;

	SELECT * FROM dbo.ASRSysFunctionParameters ORDER BY functionID, parameterIndex;

	SELECT * FROM dbo.ASRSysOperators;

	SELECT * FROM dbo.ASRSysOperatorParameters ORDER BY OperatorID, parameterIndex;
	
	-- Which modules are enabled?
	SELECT 'WORKFLOW' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,1024) AS [enabled]
	UNION
	SELECT 'PERSONNEL' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,1) AS [enabled]
	UNION
	SELECT 'ABSENCE' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,4) AS [enabled]
	UNION
	SELECT 'TRAINING' AS [name],  dbo.udfASRNetIsModuleLicensed(@licenseKey,8) AS [enabled]
	UNION
	SELECT  'VERSIONONE' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,2048) AS [enabled];


	-- Selected system settings
	SELECT * FROM ASRSysSystemSettings;


END

