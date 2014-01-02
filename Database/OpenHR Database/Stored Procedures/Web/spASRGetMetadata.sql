CREATE PROCEDURE [dbo].[spASRGetMetadata] (@Username varchar(255))
AS
BEGIN


	SELECT TableID, TableName, TableType, DefaultOrderID, RecordDescExprID FROM ASRSysTables;

	SELECT ColumnID, TableID, ColumnName, DataType, Use1000Separator, Size, Decimals FROM ASRSysColumns;

	SELECT ParentID, ChildID FROM ASRSysRelations;

	SELECT ModuleKey, ParameterKey, ISNULL(ParameterValue,'') AS ParameterValue, ParameterType FROM ASRSysModuleSetup;

	SELECT * FROM ASRSysUserSettings WHERE Username = @Username;

	SELECT functionID, functionName, returnType FROM ASRSysFunctions;

	SELECT * FROM ASRSysFunctionParameters ORDER BY functionID, parameterIndex;

	SELECT * FROM ASRSysOperators;

	SELECT * FROM ASRSysOperatorParameters ORDER BY OperatorID, parameterIndex;




END

