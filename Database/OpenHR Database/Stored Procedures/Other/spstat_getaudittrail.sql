CREATE PROCEDURE [dbo].[spstat_getaudittrail] (
		@piAuditType	int,
		@psOrder 		varchar(MAX),
		@psFilter		varchar(MAX),
		@piTop			int)
	AS
	BEGIN

		SET NOCOUNT ON;

		DECLARE @sSQL nvarchar(MAX);

		IF @piAuditType = 1
		BEGIN
			SET @sSQL = 'SELECT {TOP} 
				a.userName AS [User], 
				a.dateTimeStamp AS [Date / Time], 
				a.tableName AS [Table], 
				a.columnName AS [Column], 
				a.oldValue AS [Old Value], 
				a.newValue AS [New Value], 
				a.recordDesc AS [Record Description],
				a.id,
				CASE WHEN c.DataType = 2 OR c.DataType = 4 THEN 1 ELSE 0 END AS IsNumeric
				FROM dbo.ASRSysAuditTrail a
				LEFT JOIN ASRSysColumns c ON c.ColumnID = a.ColumnID';
		END
		ELSE IF @piAuditType = 2
			SET @sSQL =  'SELECT {TOP} 
				a.userName AS [User], 
				a.dateTimeStamp AS [Date / Time],
				a.groupName AS [User Group],
				a.viewTableName AS [View / Table],
				a.columnName AS [Column], 
				a.action AS [Action],
				a.permission AS [Permission], 
				a.id
				FROM dbo.ASRSysAuditPermissions a';
		ELSE IF @piAuditType = 3
			SET @sSQL = 'SELECT {TOP} 
				a.userName AS [User],
    			a.dateTimeStamp AS [Date / Time],
				a.groupName AS [User Group], 
				a.userLogin AS [User Login],
				a.[Action], 
				a.id
				FROM dbo.ASRSysAuditGroup a';
		ELSE IF @piAuditType = 4
			SET @sSQL = 'SELECT {TOP} 
				a.DateTimeStamp AS [Date / Time],
				a.UserGroup AS [User Group],
				a.UserName AS [User], 
				a.ComputerName AS [Computer Name],
				a.HRProModule AS [Module],
				a.Action AS [Action], 
				a.id
				FROM dbo.ASRSysAuditAccess a';
				
		IF LEN(@psFilter) > 0
			SET @sSQL = @sSQL + CHAR(10) + 'WHERE ' + @psFilter;

		IF LEN(@psOrder) > 0
			SET @sSQL = @sSQL + CHAR(10) + @psOrder;
				
		-- Retreive selected data
		IF LEN(@sSQL) > 0 
		BEGIN
			IF ISNULL(@piTop, 0) > 0
				SET @sSQL = REPLACE(@sSQL, '{TOP}', 'TOP ' + convert(varchar, @piTop));
				
			EXECUTE sp_executeSQL @sSQL;
		END

	END