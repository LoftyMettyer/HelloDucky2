CREATE PROCEDURE [dbo].[spstat_getaudittrail] (
	@piAuditType	int,
	@psOrder 		varchar(MAX))
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sSQL			varchar(MAX),
			@sExecString	nvarchar(MAX);

	IF @piAuditType = 1
	BEGIN

		SET @sSQL = 'SELECT userName AS [User], 
			dateTimeStamp AS [Date / Time], 
			tableName AS [Table], 
			columnName AS [Column], 
			oldValue AS [Old Value], 
			newValue AS [New Value], 
			recordDesc AS [Record Description],
			id
			FROM dbo.ASRSysAuditTrail ';

		IF LEN(@psOrder) > 0
			SET @sExecString = @sSQL + @psOrder;
		ELSE
			SET @sExecString = @sSQL;
		
	END
	ELSE IF @piAuditType = 2
	BEGIN

		SET @sSQL =  'SELECT userName AS [User], 
			dateTimeStamp AS [Date / Time],
			groupName AS [User Group],
			viewTableName AS [View / Table],
			columnName AS [Column], 
			action AS [Action],
			permission AS [Permission], 
			id
			FROM dbo.ASRSysAuditPermissions ';

		IF LEN(@psOrder) > 0
			SET @sExecString = @sSQL + @psOrder;
		ELSE
			SET @sExecString = @sSQL;

	END
	ELSE IF @piAuditType = 3
	BEGIN
		SET @sSQL = 'SELECT userName AS [User],
    			dateTimeStamp AS [Date / Time],
			groupName AS [User Group], 
			userLogin AS [User Login],
			[Action], 
			id
			FROM dbo.ASRSysAuditGroup ';

		IF LEN(@psOrder) > 0
			SET @sExecString = @sSQL + @psOrder;
		ELSE
			SET @sExecString = @sSQL;

	END
	ELSE IF @piAuditType = 4
	BEGIN
		SET @sSQL = 'SELECT DateTimeStamp AS [Date / Time],
    			UserGroup AS [User Group],
			UserName AS [User], 
			ComputerName AS [Computer Name],
			HRProModule AS [Module],
			Action AS [Action], 
			id
			FROM dbo.ASRSysAuditAccess ';

		IF LEN(@psOrder) > 0
			SET @sExecString = @sSQL + @psOrder;
		ELSE
			SET @sExecString = @sSQL;

	END

	-- Retreive selected data
	IF LEN(@sExecString) > 0 EXECUTE sp_executeSQL @sExecString;

END
