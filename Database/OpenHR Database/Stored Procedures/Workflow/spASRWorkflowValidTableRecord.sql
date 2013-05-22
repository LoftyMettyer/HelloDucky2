CREATE PROCEDURE [dbo].[spASRWorkflowValidTableRecord]
	@piTableID	integer,
	@piRecordID	integer,
	@pfValid	bit			OUTPUT
AS
BEGIN
	DECLARE	@sSQL	nvarchar(MAX),
			@sParam	nvarchar(500);
		
	SET @pfValid = 0;

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udf_ASRWorkflowValidTableRecord]')
			AND OBJECTPROPERTY(id, N'IsScalarFunction') = 1)
	BEGIN
		SET @sSQL = 'SET @pfValid = [dbo].[udf_ASRWorkflowValidTableRecord](' 
			+ convert(nvarchar(100), @piTableID) 
			+ ', ' 
			+ convert(nvarchar(100), @piRecordID)
			+ ')';
		SET @sParam = N'@pfValid bit OUTPUT';
		EXEC sp_executesql @sSQL, @sParam, @pfValid OUTPUT;
	END
END