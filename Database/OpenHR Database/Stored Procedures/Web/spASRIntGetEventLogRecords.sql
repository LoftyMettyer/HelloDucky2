CREATE PROCEDURE [dbo].[spASRIntGetEventLogRecords] (
	@pfError 						bit 				OUTPUT, 
	@psFilterUser					varchar(MAX),
	@piFilterType					integer,
	@piFilterStatus					integer,
	@piFilterMode					integer,
	@psOrderColumn					varchar(MAX),
	@psOrderOrder					varchar(MAX),
	@piRecordsRequired				integer,
	@pfFirstPage					bit					OUTPUT,
	@pfLastPage						bit					OUTPUT,
	@psAction						varchar(100),
	@piTotalRecCount				integer				OUTPUT,
	@piFirstRecPos					integer				OUTPUT,
	@piCurrentRecCount				integer
)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE	@sRealSource 			sysname,
			@sSelectSQL				varchar(MAX),
			@iTempCount 			integer,
			@sExecString			nvarchar(MAX),
			@sTempExecString		nvarchar(MAX),
			@sTempParamDefinition	nvarchar(500),
			@iCount					integer,
			@iGetCount				integer,
			@sFilterSQL				varchar(MAX),
			@sOrderSQL				varchar(MAX),
			@sReverseOrderSQL		varchar(MAX);
			
	/* Clean the input string parameters. */
	IF len(@psAction) > 0 SET @psAction = replace(@psAction, '''', '''''');
	IF len(@psFilterUser) > 0 SET @psFilterUser = replace(@psFilterUser, '''', '''''');
	IF len(@psOrderColumn) > 0 SET @psOrderColumn = replace(@psOrderColumn, '''', '''''');
	IF len(@psOrderOrder) > 0 SET @psOrderOrder = replace(@psOrderOrder, '''', '''''');

	/* Initialise variables. */
	SET @pfError = 0;
	SET @sExecString = '';
	SET @sRealSource = 'ASRSysEventLog';
	SET @psAction = UPPER(@psAction);

	IF (@psAction <> 'MOVEPREVIOUS') AND (@psAction <> 'MOVENEXT') AND (@psAction <> 'MOVELAST') 
		BEGIN
			SET @psAction = 'MOVEFIRST';
		END

	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 50;

	/* Construct the filter SQL from ther input parameters. */
	SET @sFilterSQL = '';
	
	SET @sFilterSQL = @sFilterSQL + ' Type NOT IN (23, 24) ';

	IF @psFilterUser <> '-1' 
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		SET @sFilterSQL = @sFilterSQL + ' LOWER(username) = ''' + lower(@psFilterUser) + '''';
	END
	IF @piFilterType <> -1
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		SET @sFilterSQL = @sFilterSQL + ' Type = ' + convert(varchar(MAX), @piFilterType) + ' ';
	END
	IF @piFilterStatus <> -1
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		SET @sFilterSQL = @sFilterSQL + ' Status = ' + convert(varchar(MAX), @piFilterStatus) + ' ';
	END
	IF @piFilterMode <> -1
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		--SET @sFilterSQL = @sFilterSQL + ' Mode = ' + convert(varchar(MAX), @piFilterMode) + ' ';
		SET @sFilterSQL = @sFilterSQL + 
			CASE @piFilterMode 
				WHEN 1 THEN '[Mode] = 1 AND ([ReportPack] = 0 OR [ReportPack] IS NULL)'
				WHEN 2 THEN '[ReportPack] = 1'
				WHEN 0 THEN '[Mode] = 0 AND ([ReportPack] = 0 OR [ReportPack] IS NULL)'
			END 
	END
	
	/* Construct the order SQL from ther input parameters. */
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
					WHEN 39 THEN ''Organisation Report''
					ELSE ''Unknown''
				END ';
	END
	ELSE
	BEGIN
		IF @psOrderColumn = 'Mode'
		BEGIN
			SET @sOrderSQL =	
				' CASE ' + @piFilterMode + '
						WHEN 1 THEN ''Batch''
						WHEN 0 THEN ''Manual''
						WHEN 2 THEN ''Pack''
					END ';
		END
		ELSE 
		BEGIN
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
			ELSE
			BEGIN
				SET @sOrderSQL = @psOrderColumn;
			END
		END
	END
	
	SET @sReverseOrderSQL = @sOrderSQL;
	if @psOrderOrder = 'DESC'
	BEGIN
		SET @sReverseOrderSQL = @sReverseOrderSQL + ' ASC ';
	END
	ELSE
	BEGIN
		SET @sReverseOrderSQL = @sReverseOrderSQL + ' DESC ';
	END

	SET @sOrderSQL = @sOrderSQL + ' ' + @psOrderOrder + ' ';


	SET @sSelectSQL = '[DateTime],
					[EndTime],
					IsNull([Duration],-1) AS ''Duration'', 
		 			CASE [Type] 
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
						WHEN 39 THEN ''Organisation Report''
						ELSE ''Unknown''  
					END + char(9) + 
				 	[Name] + char(9) + 
		 			CASE Status 
						WHEN 0 THEN ''Pending''
					  WHEN 1 THEN ''Cancelled'' 
						WHEN 2 THEN ''Failed'' 
						WHEN 3 THEN ''Successful'' 
						WHEN 4 THEN ''Skipped'' 
						WHEN 5 THEN ''Error''
						ELSE ''Unknown'' 
					END + char(9) +
					CASE 
						WHEN [Mode] = 1 AND ([ReportPack] = 0 OR [ReportPack] IS NULL) THEN ''Batch''
						WHEN [Mode] = 0 AND ([ReportPack] = 0 OR [ReportPack] IS NULL) THEN ''Manual''
						ELSE ''Pack''
				 	END + char(9) + 
					[Username] + char(9) + 
					IsNull(convert(varchar, [BatchJobID]), ''0'') + char(9) +
					IsNull(convert(varchar, [BatchRunID]), ''0'') + char(9) +
					IsNull([BatchName],'''') + char(9) +
					IsNull(convert(varchar, [SuccessCount]),''0'') + char(9) +
					IsNull(convert(varchar, [FailCount]), ''0'') AS EventInfo ';

		
	
	/****************************************************************************************************************************************/
	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.ID) FROM ' + @sRealSource;

	IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;

	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;
	SET @piTotalRecCount = @iCount;
	/****************************************************************************************************************************************/
	
	IF len(@sSelectSQL) > 0 
		BEGIN
			SET @sSelectSQL = @sRealSource + '.ID, ' + @sSelectSQL;
			SET @sExecString = 'SELECT ' ;

			IF @psAction = 'MOVEFIRST'
				BEGIN
					SET @sExecString = @sExecString + 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
					
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource ;

					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END
					
					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END

					/* Set the position variables */
					SET @piFirstRecPos = 1;
					SET @pfFirstPage = 1;
					SET @pfLastPage = 
					CASE 
						WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
						ELSE 0
					END;
				END
		
			IF (@psAction = 'MOVELAST')
				BEGIN
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource;

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource;
					
					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END

					/* Add the reverse order code */
					IF len(@sReverseOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL;
						END

					SET @sExecString = @sExecString + ')'

					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END

					/* Set the position variables */
					SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1;
					IF @piFirstRecPos < 1 SET @piFirstRecPos = 1;
					SET @pfFirstPage = 	CASE 
									WHEN @piFirstRecPos = 1 THEN 1
									ELSE 0
								END;
					SET @pfLastPage = 1;

				END

			IF (@psAction = 'MOVENEXT') 
				BEGIN
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource;

					IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
						BEGIN
							SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
						END
					ELSE
						BEGIN
							SET @iGetCount = @piRecordsRequired;
						END

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID ' + 
						' FROM ' + @sRealSource;

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired - 1) + ' ' + @sRealSource + '.ID ' + 
						' FROM ' + @sRealSource;

					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END
					
					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END
					
					SET @sExecString = @sExecString + ')';

					/* Add the reverse order code */
					IF len(@sReverseOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL;
						END

					SET @sExecString = @sExecString + ')';

					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END

					/* Set the position variables */
					SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount;
					SET @pfFirstPage = 0
					SET @pfLastPage = 	CASE 
									WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
									ELSE 0
								END;
				END

			IF @psAction = 'MOVEPREVIOUS'
				BEGIN	
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource;

					IF @piFirstRecPos <= @piRecordsRequired
						BEGIN
							SET @iGetCount = @piFirstRecPos - 1;
						END
					ELSE
						BEGIN
							SET @iGetCount = @piRecordsRequired;
						END
		
					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource;

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource;
				
					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END
					
					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END
					
					SET @sExecString = @sExecString + ')';

					/* Add the reverse order code */
					IF len(@sReverseOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL + ')';
						END
					
					SET @sExecString = @sExecString

					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END
		
					/* Set the position variables */
					SET @piFirstRecPos = @piFirstRecPos - @iGetCount;
					IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1;
					SET @pfFirstPage = CASE WHEN @piFirstRecPos = 1 
															THEN 1
															ELSE 0
														 END;
					SET @pfLastPage = CASE WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount 
															THEN 1
															ELSE 0
														END;
				END

		END

	EXECUTE sp_executeSQL @sExecString;
END