CREATE PROCEDURE [dbo].[spASRIntGetParameterPassByType] (
	@piFunctionID		integer,
	@piParamIndex		integer,
	@piResult				integer		OUTPUT
) 
AS

	/* Return 1 if the given function parameter is passed by value
	Return 2 if the given parameter is passed by reference */
	DECLARE @iResult	integer
	SET @piResult = 1
	SELECT @iResult = parameterType
	FROM ASRSysFunctionParameters
	WHERE functionID = @piFunctionID
		AND parameterIndex = @piParamIndex;

	IF @iResult IS null SET @iResult = 0;
	IF @iResult >= 100 SET @piResult = 2;

GO