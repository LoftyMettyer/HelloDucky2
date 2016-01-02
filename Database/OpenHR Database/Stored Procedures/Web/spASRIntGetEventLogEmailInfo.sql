CREATE PROCEDURE [dbo].[spASRIntGetEventLogEmailInfo] (
	@psSelectedIDs	varchar(MAX),
	@psSubject		varchar(MAX) OUTPUT,
	@psOrderColumn	varchar(MAX),
	@psOrderOrder	varchar(MAX)
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sSelectString 	nvarchar(MAX),
			@sOrderSQL		varchar(MAX);
	
	/* Clean the input string parameters. */
	IF len(@psSelectedIDs) > 0 SET @psSelectedIDs = replace(@psSelectedIDs, '''', '''''');
	IF len(@psOrderColumn) > 0 SET @psOrderColumn = replace(@psOrderColumn, '''', '''''');
	IF len(@psOrderOrder) > 0 SET @psOrderOrder = replace(@psOrderOrder, '''', '''''');

	SELECT @psSubject = IsNull(SettingValue,'<<Unknown Customer>>') + ' - Event Log' 
	FROM ASRSysSystemSettings 
	WHERE Lower(Section) = 'licence' 
		AND Lower(SettingKey) = 'customer name';

	SET @sSelectString = '';

	/* create SELECT statment string */
	SET @sSelectString = 'SELECT 	A.ID, 
		A.Name, 
		A.DateTime,
		A.EndTime,
		IsNull(A.Duration,-1) AS Duration, 
		A.Username, 
		CASE A.Mode 
			WHEN 1 THEN ''Batch'' 
			ELSE ''Manual'' 
		END AS ''Mode'', 
		CASE A.Status 
			WHEN 0 THEN ''Pending''
		  WHEN 1 THEN ''Cancelled'' 
			WHEN 2 THEN ''Failed'' 
			WHEN 3 THEN ''Successful'' 
			WHEN 4 THEN ''Skipped'' 
			WHEN 5 THEN ''Error''
			ELSE ''Unknown'' 
		END AS Status, 
		CASE A.Type 
			WHEN 0 THEN ''Unknown''
			WHEN 1 THEN ''Cross Tab'' 
			WHEN 2 THEN ''Custom Report'' 
			WHEN 3 THEN ''Data Transfer'' 
			WHEN 4 THEN ''Export'' 
			WHEN 5 THEN ''Global Add'' 
			WHEN 6 THEN ''Global Delete'' 
			WHEN 7 THEN ''Global Update'' 
			WHEN 8 THEN ''Import'' 
			WHEN 9 THEN ''Mail Merge'' 
			WHEN 10 THEN ''Diary Delete'' 
			WHEN 11 THEN ''Diary Rebuild''
			WHEN 12 THEN ''Email Rebuild''
			WHEN 13 THEN ''Standard Report''
			WHEN 14 THEN ''Record Editing''
			WHEN 15 THEN ''System Error''
			WHEN 16 THEN ''Match Report''
			WHEN 17 THEN ''Calendar Report''
			WHEN 18 THEN ''Envelopes & Labels''
			WHEN 19 THEN ''Label Definition''
			WHEN 20 THEN ''Record Profile''
			WHEN 21	THEN ''Succession Planning''
			WHEN 22 THEN ''Career Progression''
			WHEN 25 THEN ''Workflow Rebuild''
			WHEN 35 THEN ''9-Box Grid Report''
			WHEN 38 THEN ''Talent Report''
			ELSE ''Unknown''  
		END AS Type,
		CASE 
			WHEN A.SuccessCount IS NULL THEN ''N/A''
			ELSE CONVERT(varchar, A.SuccessCount)
		END AS SuccessCount,
		CASE
			WHEN A.FailCount IS NULL THEN ''N/A''
			ELSE CONVERT(varchar, A.FailCount)
		END AS FailCount,
		A.BatchName AS BatchName,
		A.BatchJobID AS BatchJobID,
		A.BatchRunID AS BatchRunID,
		B.Notes, 
		B.ID AS ''DetailsID'' ,
		(SELECT count(ID) 
			FROM ASRSysEventLogDetails C 
			WHERE C.EventLogID = A.ID) as ''count''
		FROM ASRSysEventLog A
		LEFT OUTER JOIN ASRSysEventLogDetails B
			ON A.ID = B.EventLogID
		WHERE A.ID IN (' + @psSelectedIDs + ')';

	SET @sOrderSQL = '';
	
	IF @psOrderColumn = 'Type'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Type] 
				WHEN 1 THEN ''Cross Tab''
				WHEN 2 THEN ''Custom Report''
				WHEN 3 THEN ''Data Transfer''
				WHEN 4 THEN ''Export''
				WHEN 5 THEN ''Global Add''
				WHEN 6 THEN ''Global Delete''
				WHEN 7 THEN ''Global Update''
				WHEN 8 THEN ''Import''
				WHEN 9 THEN ''Mail Merge''
				WHEN 10 THEN ''Diary Delete''
				WHEN 11 THEN ''Diary Rebuild''
				WHEN 12 THEN ''Email Rebuild''
				WHEN 13 THEN ''Standard Report''
				WHEN 14 THEN ''Record Editing''
				WHEN 15 THEN ''System Error''
				WHEN 16 THEN ''Match Report''
				WHEN 17 THEN ''Calendar Report''
				WHEN 18 THEN ''Envelopes & Labels''
				WHEN 19 THEN ''Label Definition''
				WHEN 20 THEN ''Record Profile''
				WHEN 21 THEN ''Succession Planning''
				WHEN 22 THEN ''Career Progression''
				WHEN 25 THEN ''Workflow Rebuild''
				WHEN 35 THEN ''9-Box Grid Report''
				WHEN 38 THEN ''Talent Report''
				ELSE ''Unknown''
			END ';
	END

	IF @psOrderColumn = 'Mode'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Mode] 
				WHEN 1 THEN ''Batch''
				WHEN 0 THEN ''Manual''
				ELSE ''Unknown''
			END ';
	END
	
	IF @psOrderColumn = 'Status'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Status] 
				WHEN 0 THEN ''Pending''
				WHEN 1 THEN ''Cancelled''
				WHEN 2 THEN ''Failed''
				WHEN 3 THEN ''Successful''
				WHEN 4 THEN ''Skipped''
				WHEN 5 THEN ''Error''
				ELSE ''Unknown''
			END ';
	END
	
	IF len(@sOrderSQL) = 0
	BEGIN
		SET @sOrderSQL = @psOrderColumn;
	END
	
	SET @sOrderSQL = @sOrderSQL + ' ' + @psOrderOrder;

	IF LEN(LTRIM(RTRIM(@sOrderSQL))) > 0 
	BEGIN
		SET @sSelectString = @sSelectString + ' ORDER BY ' + @sOrderSQL;
	END

	EXEC sp_executeSQL @sSelectString;
END