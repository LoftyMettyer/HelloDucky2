CREATE PROCEDURE [dbo].[sp_ASRIntGetEmpIDFromWLID] (
	@piEmpRecordID	integer		OUTPUT,
	@piWLRecordID	integer
)
AS
BEGIN
	DECLARE @iUserGroupID		integer,
		@sUserGroupName			sysname,
		@iChildViewID 			integer,
		@sTempExecString		nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@iEmpTableID			integer,
		@iWLTableID				integer,
		@sWLRealSource			varchar(255),
		@sWLTableName			sysname,
		@sActualUserName		sysname;

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* NB. To reach this point we have already checked that the user has 'read' permission on the Waiting List table. */
	SELECT @iEmpTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable'
	IF @iEmpTableID IS NULL SET @iEmpTableID = 0;

	SELECT @iWLTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_WaitListTable'
	IF @iWLTableID IS NULL SET @iWLTableID = 0;

	SELECT @sWLTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iWLTableID;
	
	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iWLTableID
		AND role = @sUserGroupName;
		
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sWLRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sWLTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_')
		SET @sWLRealSource = left(@sWLRealSource, 255);
	END

	SET @sTempExecString = 'SELECT @iID = ID_' + convert(nvarchar(100), @iEmpTableID) +
		' FROM ' + @sWLRealSource +
		' WHERE id = ' + convert(nvarchar(100), @piWLRecordID);
	SET @sTempParamDefinition = N'@iID integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @piEmpRecordID OUTPUT;

	IF @piEmpRecordID IS null SET @piEmpRecordID = 0;
END