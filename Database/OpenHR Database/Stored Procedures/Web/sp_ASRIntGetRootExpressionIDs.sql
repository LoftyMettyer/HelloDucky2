CREATE PROCEDURE [dbo].[sp_ASRIntGetRootExpressionIDs] (
	@piCompID		integer,
	@piRootExprID	varchar(255)	OUTPUT)
AS
BEGIN
	DECLARE @iParentCompID	integer;

	SELECT @iParentCompID = ASRSysExpressions.parentComponentID, 
		@piRootExprID = ASRSysExpressions.exprID
	FROM [dbo].[ASRSysExpressions]
	JOIN ASRSysExprComponents ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID
	WHERE ASRSysExprComponents.componentID = @piCompID;

	IF (@iParentCompID > 0)
	BEGIN
		EXECUTE [dbo].[sp_ASRIntGetRootExpressionIDs] @iParentCompID, @piRootExprID OUTPUT;
	END

	IF @piRootExprID IS null SET @piRootExprID = 0;
END