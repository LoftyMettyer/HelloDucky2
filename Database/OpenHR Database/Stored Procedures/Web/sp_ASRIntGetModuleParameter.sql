CREATE PROCEDURE [dbo].[sp_ASRIntGetModuleParameter]
(
	@psModuleKey 		varchar(255), 
	@psParameterKey 	varchar(255),
	@psParameter		varchar(1000) OUTPUT
)
AS
BEGIN
	SELECT @psParameter = parameterValue 
	FROM [dbo].[ASRSysModuleSetup]
	WHERE moduleKey = @psModuleKey 
		AND parameterKey = @psParameterKey;
END