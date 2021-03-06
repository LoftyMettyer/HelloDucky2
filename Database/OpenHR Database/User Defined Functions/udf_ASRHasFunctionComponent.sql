
CREATE FUNCTION [dbo].[udf_ASRHasFunctionComponent]
	(
		@piExpressionID integer,
		@piFunctionID integer
	)
	RETURNS integer
	AS
BEGIN

		DECLARE @iCount integer
		DECLARE @bFound bit

		SELECT @iCount = COUNT(*)
			FROM ASRSysExprComponents c 
			LEFT OUTER JOIN ASRSysExpressions e ON c.ComponentID = e.parentComponentID
			WHERE c.ExprID = @piExpressionID AND
				((c.Type = 1 AND dbo.udf_ASRHasFunctionComponent(c.FieldSelectionFilter,@piFunctionID) > 0) OR
				(c.Type = 2 AND c.FunctionID = @piFunctionID) OR
				(c.Type = 2 AND dbo.udf_ASRHasFunctionComponent(e.exprID,@piFunctionID) > 0) OR
				(c.Type = 3 AND dbo.udf_ASRHasFunctionComponent(c.CalculationID,@piFunctionID) > 0) OR
				(c.Type = 10 AND dbo.udf_ASRHasFunctionComponent(c.FilterID,@piFunctionID) > 0))

		IF @iCount > 0 SET @bFound = 1
		ELSE SET @bFound = 0
		
		RETURN @bFound

	END
GO
