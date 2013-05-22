CREATE PROCEDURE [dbo].[spstat_audittable] (
	@piTableID int,
	@piRecordID int,
	@psRecordDesc varchar(255),
	@psValue varchar(MAX))
AS
BEGIN	
	DECLARE @sTableName varchar(128);

	-- Get the table name for the given column.
	SELECT @sTableName = tableName 
		FROM dbo.ASRSysTables
		WHERE tableID = @piTableID;

	IF @sTableName IS NULL SET @sTableName = '<Unknown>';

	-- Insert a record into the Audit Trail table.
	INSERT INTO dbo.ASRSysAuditTrail 
		(userName, 
		dateTimeStamp, 
		tablename, 
		recordID, 
		recordDesc, 
		columnname, 
		oldValue, 
		newValue,
		ColumnID, 
		Deleted)
	VALUES 
		(CASE
			WHEN UPPER(LEFT(APP_NAME(), 15)) = 'OPENHR WORKFLOW' THEN 'OpenHR Workflow'
			ELSE user
		END, 
		getDate(), 
		@sTableName, 
		@piRecordID, 
		@psRecordDesc, 
		'', 
		'', 
		@psValue,
		0, 
		0);
END

