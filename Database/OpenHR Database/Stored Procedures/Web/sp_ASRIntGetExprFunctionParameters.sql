CREATE PROCEDURE [dbo].[sp_ASRIntGetExprFunctionParameters] 
AS
BEGIN
	/* Return a recordset of the runtiume function partameter definitions. */
	DECLARE @fEnableUDFFunctions	bit,
			@sSQLVersion int

	SET @fEnableUDFFunctions = 0
	SELECT @sSQLVersion = dbo.udfASRSQLVersion()

	IF @sSQLVersion >= 8
	BEGIN  
		SET @fEnableUDFFunctions = 1
	END
	SELECT ASRSysFunctions.functionID, 
		ASRSysFunctionParameters.parameterName
	FROM ASRSysFunctions
	LEFT OUTER JOIN ASRSysFunctionParameters ON ASRSysFunctions.functionID = ASRSysFunctionParameters.functionID
	WHERE (ASRSysFunctions.runtime = 1)
		OR ((ASRSysFunctions.UDF = 1) AND (@fEnableUDFFunctions = 1))
	ORDER BY ASRSysFunctions.functionID, ASRSysFunctionParameters.parameterIndex
END
