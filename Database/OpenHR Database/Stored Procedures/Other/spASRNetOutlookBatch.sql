CREATE PROCEDURE [dbo].[spASRNetOutlookBatch]
(
	@Content			varchar(MAX)	OUTPUT,
	@AllDayEvent		bit OUTPUT,
	@StartDate			datetime		OUTPUT,
	@EndDate			datetime		OUTPUT,
	@StartTime			varchar(100)	OUTPUT,
	@EndTime			varchar(100)	OUTPUT,
	@Subject			varchar(MAX)	OUTPUT,
	@Folder				varchar(MAX)	OUTPUT,
	@LinkID				integer,
	@RecordID			integer,
	@FolderID			integer,
	@StartDateColumnID	integer,
	@EndDateColumnID	integer,
	@FixedStartTime		varchar(100),
	@FixedEndTime		varchar(100),
	@StartTimeColumnID	integer,
	@EndTimeColumnID	integer,
	@TimeRange			integer,
	@Title				varchar(MAX),
	@SubjectExprID		integer,
	@RecordDescExprID	integer,
	@DateFormat			varchar(100),
	@FolderPath			varchar(MAX),
	@FolderType			integer,
	@FolderExprID		integer)
AS
BEGIN

	DECLARE @sSQL nvarchar(MAX);
	DECLARE @sParamDefinition nvarchar(500);
	DECLARE @CharValue varchar(MAX);
	DECLARE @Heading varchar(MAX);
	DECLARE @TableName varchar(MAX);
	DECLARE @ColumnName varchar(MAX);
	DECLARE @DataType integer;
		
	SELECT @sSQL = 'SELECT @StartDate=['+ColumnName+'] FROM ['+TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
	FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
	WHERE ColumnID = @StartDateColumnID;
	SET @sParamDefinition = N'@StartDate datetime OUTPUT';
	EXEC sp_executesql @sSQL,  @sParamDefinition, @StartDate OUTPUT;

	SET @EndDate = Null
	IF @EndDateColumnID > 0
	BEGIN
		SELECT @sSQL = 'SELECT @EndDate=['+ColumnName+'] FROM ['+TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
		FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
		WHERE ColumnID = @EndDateColumnID
		SET @sParamDefinition = N'@EndDate datetime OUTPUT'
		EXEC sp_executesql @sSQL,  @sParamDefinition, @EndDate OUTPUT
		IF rtrim(@EndDate) = '' SET @EndDate = null
	END

	IF @TimeRange = 0
	BEGIN
		SET @AllDayEvent = 1
		SET @StartTime = ''
		SET @EndTime = ''
	END
	IF @TimeRange = 1
	BEGIN
		SET @AllDayEvent = 0
		SET @StartTime = @FixedStartTime
		SET @EndTime = @FixedEndTime
	END
	IF @TimeRange = 2
	BEGIN
		SET @AllDayEvent = 0

		SELECT @sSQL = 'SELECT @StartTime=['+ColumnName+'] FROM ['+TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
		FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
		WHERE ColumnID = @StartTimeColumnID
		SET @sParamDefinition = N'@StartTime varchar(100) OUTPUT'
		EXEC sp_executesql @sSQL,  @sParamDefinition, @StartTime OUTPUT

		SELECT @sSQL = 'SELECT @EndTime=['+ColumnName+'] FROM ['+TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
		FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
		WHERE ColumnID = @EndTimeColumnID
		SET @sParamDefinition = N'@EndTime varchar(100) OUTPUT'
		EXEC sp_executesql @sSQL,  @sParamDefinition, @EndTime OUTPUT

		IF UPPER(@StartTime) = 'AM'
			SELECT @StartTime = SettingValue FROM ASRSysSystemSettings
			WHERE [Section] = 'outlook' and [Settingkey] = 'amstarttime'
		IF UPPER(@StartTime) = 'PM'
			SELECT @StartTime = SettingValue FROM ASRSysSystemSettings
			WHERE [Section] = 'outlook' and [Settingkey] = 'pmstarttime'
		IF UPPER(@EndTime) = 'AM'
			SELECT @EndTime = SettingValue FROM ASRSysSystemSettings
			WHERE [Section] = 'outlook' and [Settingkey] = 'amendtime'
		IF UPPER(@EndTime) = 'PM'
			SELECT @EndTime = SettingValue FROM ASRSysSystemSettings
			WHERE [Section] = 'outlook' and [Settingkey] = 'pmendtime'
	END


	SET @Subject = ''
	IF @SubjectExprID > 0
	BEGIN
		SET @sSQL = 'DECLARE @hResult int
			IF EXISTS(SELECT * FROM sysobjects WHERE type = ''P'' AND name = ''sp_ASRExpr_'+convert(nvarchar(100),@SubjectExprID)+''')
		             BEGIN
		                EXEC @hResult = sp_ASRExpr_'+convert(nvarchar(100),@SubjectExprID)+' @Subject OUTPUT, '+convert(nvarchar(100),@RecordID)+'
		                IF @hResult <> 0 SET @Subject = ''''
		                SET @Subject = CONVERT(varchar(255), @Subject)
			     END
			     ELSE SET @Subject = '''''
		SET @sParamDefinition = N'@Subject varchar(MAX) OUTPUT'
		EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
	END
	ELSE
	BEGIN
		IF @RecordDescExprID > 0
		BEGIN
			SET @sSQL = 'DECLARE @hResult int
				IF EXISTS(SELECT * FROM sysobjects WHERE type = ''P'' AND name = ''sp_ASRExpr_'+convert(nvarchar(100),@RecordDescExprID)+''')
			             BEGIN
			                EXEC @hResult = sp_ASRExpr_'+convert(nvarchar(100),@RecordDescExprID)+' @Subject OUTPUT, '+convert(nvarchar(100),@RecordID)+'
			                IF @hResult <> 0 SET @Subject = ''''
			                SET @Subject = CONVERT(varchar(255), @Subject)
				     END
				     ELSE SET @Subject = '''''
			SET @sParamDefinition = N'@Subject varchar(MAX) OUTPUT'
			EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
			IF @Subject <> ''
				SET @Subject = ': '+@Subject
		END
		SET @Subject = @Title+@Subject
	END


	SET @Folder = @FolderPath
	IF @FolderType > 0
	BEGIN
		SET @sSQL = 'DECLARE @hResult int
			IF EXISTS(SELECT * FROM sysobjects WHERE type = ''P'' AND name = ''sp_ASRExpr_'+convert(nvarchar(100),@FolderExprID)+''')
		             BEGIN
		                EXEC @hResult = sp_ASRExpr_'+convert(nvarchar(100),@FolderExprID)+' @Folder OUTPUT, '+convert(nvarchar(100),@RecordID)+'
		                IF @hResult <> 0 SET @Folder = ''''
			     END
			     ELSE SET @Folder = '''''
		SET @sParamDefinition = N'@Folder varchar(MAX) OUTPUT'
		EXEC sp_executesql @sSQL,  @sParamDefinition, @Folder OUTPUT
	END


	DECLARE CursorColumns CURSOR FOR 
	SELECT isnull(ASRSysOutlookLinksColumns.Heading,''),
		ASRSysTables.TableName,
		ASRSysColumns.ColumnName,
		ASRSysColumns.DataType
	FROM ASRSysOutlookLinksColumns
	JOIN ASRSysColumns
		ON ASRSysColumns.ColumnID = ASRSysOutlookLinksColumns.ColumnID
	JOIN ASRSysTables
		ON ASRSysColumns.TableID = ASRSysTables.TableID
	WHERE LinkID = @LinkID
	ORDER BY [Sequence] DESC

	SET @Content = char(13) + @Content

	OPEN CursorColumns
	FETCH NEXT FROM CursorColumns
	INTO	@Heading, @TableName, @ColumnName, @DataType

	WHILE @@FETCH_STATUS = 0
	BEGIN

		IF @Heading <> '' SET @Heading = @Heading+': '

		IF @DataType = 12
			SELECT @sSQL = 'SELECT @CharValue='''+@Heading+'''+isnull(['+@ColumnName+'],'''') FROM ['+@TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
		IF @DataType = 11
			SELECT @sSQL = 'SELECT @CharValue='''+@Heading+'''+case when ['+@ColumnName+'] is null then ''<Empty>'' else convert(varchar(255),['+@ColumnName+'],'+@DateFormat+') end FROM ['+@TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
		IF @DataType = -7
			SELECT @sSQL = 'SELECT @CharValue='''+@Heading+'''+case when ['+@ColumnName+'] = 1 then ''Y'' else ''N'' end FROM ['+@TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)
		IF @DataType <> 11 AND @DataType <> 12 AND @DataType <> -7
			SELECT @sSQL = 'SELECT @CharValue='''+@Heading+'''+convert(varchar(255),isnull(['+@ColumnName+'],'''')) FROM ['+@TableName+'] WHERE ID = '+convert(nvarchar(100),@RecordID)

		SET @sParamDefinition = N'@CharValue varchar(MAX) OUTPUT'
		EXEC sp_executesql @sSQL,  @sParamDefinition, @CharValue OUTPUT

		IF @CharValue IS Null SET @CharValue = ''
		SET @Content = @CharValue + char(13) + @Content

		FETCH NEXT FROM CursorColumns
		INTO	@Heading, @TableName, @ColumnName, @DataType
	END

	CLOSE CursorColumns
	DEALLOCATE CursorColumns

END