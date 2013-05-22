CREATE PROCEDURE [dbo].[spASRIntGetRecordSelection]
(
	@psType		varchar(255),
	@piTableID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @fSysSecMgr	bit;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT
	
	IF UPPER(@psType) = 'PICKLIST'
	BEGIN
		SELECT picklistid, 
			name, 
			username, 
			access 
		FROM [dbo].[ASRSysPicklistName]
		WHERE (tableid = @piTableID)
			AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
		ORDER BY [name];
	END

	IF UPPER(@psType) = 'FILTER'
	BEGIN
		SELECT exprid, 
			name, 
			username, 
			access 
		FROM [dbo].[ASRSysExpressions]
		WHERE tableid = @piTableID 
			AND type = 11 
			AND (returnType = 3 OR type = 10) 
			AND parentComponentID = 0 
			AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
		ORDER BY [name];
	END
	
	IF UPPER(@psType) = 'CALC'
	BEGIN
		IF @piTableID > 0
		BEGIN
			SELECT exprid, 
				name, 
				username, 
				access 
			FROM [dbo].[ASRSysExpressions]
			WHERE (tableid = @piTableID)
				AND  type = 10 
				AND (returnType = 0 OR type = 10) 
				AND parentComponentID = 0 
				AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
			ORDER BY [name];
		END
		ELSE
		BEGIN
			SELECT exprid, 
				name, 
				username, 
				access 
			FROM [dbo].[ASRSysExpressions] 
			WHERE  type = 18 
				AND (returnType = 4 OR type = 10) 
				AND parentComponentID = 0 
				AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
			ORDER BY [name];
		END
	END
END