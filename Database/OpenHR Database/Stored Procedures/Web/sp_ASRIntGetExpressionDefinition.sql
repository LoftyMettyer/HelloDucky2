CREATE PROCEDURE [dbo].[sp_ASRIntGetExpressionDefinition] (
	@piExprID		integer,
	@psAction		varchar(100),
	@psErrMsg		varchar(MAX)	OUTPUT,
	@piTimestamp	integer			OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return the defintions of each component and expression in the given expression. */
	DECLARE @sExprIDs		varchar(MAX),
		@sComponentIDs		varchar(MAX),
		@sTempExprIDs		varchar(MAX),
		@sTempComponentIDs	varchar(MAX),
		@sCurrentUser		sysname,
		@iCount				integer,
		@sOwner				varchar(255),
		@sAccess			varchar(MAX),
		@iBaseTableID		integer,
		@sBaseTableID		varchar(100),
		@fSysSecMgr			bit,
		@sExecString		nvarchar(MAX);
	
	SET @psErrMsg = '';
	SET @sCurrentUser = SYSTEM_USER;

	/* Check the expressions exists. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysExpressions
	WHERE exprID = @piExprID;

	IF @iCount = 0
	BEGIN
		SET @psErrMsg = 'expression has been deleted by another user.';
		RETURN;
	END

	SELECT @sOwner = userName,
		@sAccess = access,
		@iBaseTableID = tableID,
		@piTimestamp = convert(integer, timestamp)
	FROM ASRSysExpressions
	WHERE exprID = @piExprID;

	IF @sAccess <> 'RW'
	BEGIN
		exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;
	
		IF @fSysSecMgr = 1 SET @sAccess = 'RW';
	END
	
	IF @iBaseTableID IS null 
	BEGIN
		SET @sBaseTableID = '0';
	END
	ELSE
	BEGIN
		SET @sBaseTableID = convert(varchar(100), @iBaseTableID);
	END

	/* Check the current user can view the expression. */
	IF (@sAccess = 'HD') AND (@sOwner <> @sCurrentUser) 
	BEGIN
		SET @psErrMsg = 'expression has been made hidden by another user.';
		RETURN;
	END

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@sOwner <> @sCurrentUser) 
	BEGIN
		SET @psErrMsg = 'expression has been made read only by another user.';
		RETURN;
	END

	SET @sExprIDs = convert(varchar(MAX), @piExprID);
	SET @sComponentIDs = '0';

	/* Get a list of the components and sub-expressions in the given expression. */
	exec sp_ASRIntGetSubExpressionsAndComponents @piExprID, @sTempExprIDs OUTPUT, @sTempComponentIDs OUTPUT;

	IF len(@sTempExprIDs) > 0 SET @sExprIDs = @sExprIDs + ',' + @sTempExprIDs;
	IF len(@sTempComponentIDs) > 0 SET @sComponentIDs = @sComponentIDs + ',' + @sTempComponentIDs;

	SET @sExecString = 'SELECT
		''C'' as [type],
		ASRSysExprComponents.componentID AS [id],
		convert(varchar(100), ASRSysExprComponents.componentID)+ char(9) +
		convert(varchar(100), ASRSysExprComponents.exprID)+ char(9) +
		convert(varchar(100), ASRSysExprComponents.type)+ char(9) +
		CASE WHEN ASRSysExprComponents.fieldColumnID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldColumnID) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldPassBy IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldPassBy) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionTableID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionTableID) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionRecord IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionRecord) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionLine IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionLine) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionOrderID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionOrderID) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionFilter IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionFilter) END + char(9) +
		CASE WHEN ASRSysExprComponents.functionID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.functionID) END + char(9) +
		CASE WHEN ASRSysExprComponents.calculationID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.calculationID) END + char(9) +
		CASE WHEN ASRSysExprComponents.operatorID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.operatorID) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueType IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.valueType) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueCharacter IS null THEN '''' ELSE ASRSysExprComponents.valueCharacter END + char(9) +
		CASE WHEN ASRSysExprComponents.valueNumeric IS null THEN '''' ELSE convert(varchar(100), convert(numeric(38, 2), ASRSysExprComponents.valueNumeric)) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueLogic IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.valueLogic) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueDate IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.valueDate, 101) END + char(9) +
		CASE WHEN ASRSysExprComponents.promptDescription IS null THEN '''' ELSE ASRSysExprComponents.promptDescription END + char(9) +
		CASE WHEN ASRSysExprComponents.promptMask IS null THEN '''' ELSE ASRSysExprComponents.promptMask END + char(9) +
		CASE WHEN ASRSysExprComponents.promptSize IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.promptSize) END + char(9) +
		CASE WHEN ASRSysExprComponents.promptDecimals IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.promptDecimals) END + char(9) +
		CASE WHEN ASRSysExprComponents.functionReturnType IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.functionReturnType) END + char(9) +
		CASE WHEN ASRSysExprComponents.lookupTableID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.lookupTableID) END + char(9) +
		CASE WHEN ASRSysExprComponents.lookupColumnID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.lookupColumnID) END + char(9) +
		CASE WHEN ASRSysExprComponents.filterID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.filterID) END + char(9) +
		CASE WHEN ASRSysExprComponents.expandedNode IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.expandedNode) END + char(9) + 
		CASE WHEN ASRSysExprComponents.promptDateType IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.promptDateType) END + char(9) + 
		CASE 
			WHEN ASRSysExprComponents.type = 1 THEN fldtabs.tablename + 
				CASE 
					WHEN (ASRSysExprComponents.fieldPassBy = 2) OR (ASRSysExprComponents.fieldSelectionRecord <> 5) then '' : '' + fldcols.columnname
					ELSE ''''
				END +
				CASE 
					WHEN ASRSysExprComponents.fieldPassBy = 2 then ''''
					ELSE
						CASE 
							WHEN fldrelations.parentID IS null THEN ''''
							ELSE
								CASE 
									WHEN ASRSysExprComponents.fieldSelectionRecord = 1 THEN '' (first record''
									WHEN ASRSysExprComponents.fieldSelectionRecord = 2 THEN '' (last record''
									WHEN ASRSysExprComponents.fieldSelectionRecord = 3 THEN '' (line '' + convert(varchar(100), ASRSysExprComponents.fieldSelectionLine)
									WHEN ASRSysExprComponents.fieldSelectionRecord = 4 THEN '' (total''
									WHEN ASRSysExprComponents.fieldSelectionRecord = 5 THEN '' (record count''
									ELSE '' (''
								END +
								CASE 
									WHEN fldorders.name IS null THEN ''''
									ELSE '', order by '''''' + fldorders.name + ''''''''
								END  +
								CASE 
									WHEN fldfilters.name IS null then ''''
									ELSE '', filter by '''''' + fldfilters.name + ''''''''
								END + 
								'')''
						END
				END
			WHEN ASRSysExprComponents.type = 2 THEN ASRSysFunctions.functionName
			WHEN ASRSysExprComponents.type = 3 THEN calcexprs.name
			WHEN ASRSysExprComponents.type = 5 THEN ASRSysOperators.name
			WHEN ASRSysExprComponents.type = 10 THEN filtexprs.name
			ELSE ''''
		END + char(9) +
		CASE WHEN fldcols.tableID IS null THEN '''' ELSE convert(varchar(100), fldcols.tableID) END + char(9) + 
		CASE WHEN fldorders.name IS null THEN '''' ELSE fldorders.name END + char(9) + 
		CASE WHEN fldfilters.name IS null THEN '''' ELSE fldfilters.name END
		AS [definition]
	FROM ASRSysExprComponents
	LEFT OUTER JOIN ASRSysExpressions calcexprs ON ASRSysExprComponents.calculationID = calcexprs.exprID
	LEFT OUTER JOIN ASRSysExpressions filtexprs ON ASRSysExprcomponents.filterID = filtexprs.exprID
	LEFT OUTER JOIN ASRSysColumns fldcols ON ASRSysExprComponents.FieldColumnID = fldcols.columnID
	LEFT OUTER JOIN ASRSysTables fldtabs ON fldcols.tableID = fldtabs.tableID
	LEFT OUTER JOIN ASRSysFunctions ON ASRSysExprComponents.functionID = asrsysfunctions.functionID 
	LEFT OUTER JOIN ASRSysOperators ON ASRSysExprComponents.operatorID = asrsysoperators.operatorID 
	LEFT OUTER JOIN ASRSysRelations fldrelations ON (ASRSysExprComponents.fieldTableID = fldrelations.childID and fldrelations.parentID = ' + @sBaseTableID + ')
	LEFT OUTER JOIN ASRSysOrders fldorders ON ASRSysExprComponents.fieldSelectionOrderID = fldorders.orderID
	LEFT OUTER JOIN ASRSysExpressions fldfilters ON ASRSysExprComponents.fieldSelectionFilter = fldfilters.exprID	
	WHERE ASRSysExprComponents.componentID IN (' + @sComponentIDs + ')
	UNION
	SELECT 	
		''E'' as [type],
		ASRSysExpressions.exprID AS [id],
		convert(varchar(100), ASRSysExpressions.exprID)+ char(9) +
		ASRSysExpressions.name + char(9) +
		convert(varchar(100), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(100), ASRSysExpressions.returnType) + char(9) +
		convert(varchar(100), ASRSysExpressions.returnSize) + char(9) +
		convert(varchar(100), ASRSysExpressions.returnDecimals) + char(9) +
		convert(varchar(100), ASRSysExpressions.type) + char(9) +
		convert(varchar(100), ASRSysExpressions.parentComponentID) + char(9) +
		ASRSysExpressions.userName + char(9) +
		ASRSysExpressions.access + char(9) +
		CASE WHEN ASRSysExpressions.description IS null THEN '''' ELSE ASRSysExpressions.description END + char(9) +
		convert(varchar(100), convert(integer, ASRSysExpressions.timestamp)) + char(9) + 
		convert(varchar(100), isnull(ASRSysExpressions.viewInColour, 0)) + char(9) +
		convert(varchar(100), isnull(ASRSysExpressions.expandedNode, 0)) AS [definition]
	FROM ASRSysExpressions
	WHERE ASRSysExpressions.exprID IN (' + @sExprIDs + ')
	ORDER BY [id]';
	
	EXECUTE sp_EXecuteSQL @sExecString;
END