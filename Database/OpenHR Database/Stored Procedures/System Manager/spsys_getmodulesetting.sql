CREATE PROCEDURE [dbo].[spsys_getmodulesetting](
	@moduleKey AS varchar(50),
	@parameterKey AS varchar(50),
	@paramterType AS varchar(50),			
	@parameterValue AS nvarchar(MAX) OUTPUT)
AS
BEGIN
	SELECT @parameterValue = [parameterValue] FROM [asrsysModuleSetup] WHERE [ModuleKey] = @moduleKey 
		AND [ParameterKey] = @parameterKey AND [ParameterType] = @paramterType;
END