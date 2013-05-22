CREATE PROCEDURE [dbo].[sp_ASRGetControlDetails] 
	(@piScreenID int)
AS
BEGIN
	SELECT cont.*, 
		col.[columnName], col.[columnType], col.[datatype], col.[defaultValue],
		col.[size], col.[decimals], col.[lookupTableID], 
		col.[lookupColumnID], col.[lookupFilterColumnID], col.[lookupFilterOperator], col.[lookupFilterValueID], 
		col.[spinnerMinimum], col.[spinnerMaximum], col.[spinnerIncrement], 
		col.[mandatory], col.[uniquecheck], col.[uniquechecktype], col.[convertcase], 
		col.[mask], col.[blankIfZero], col.[multiline], col.[alignment] AS colAlignment, 
		col.[calcExprID], col.[gotFocusExprID], col.[lostFocusExprID], col.[dfltValueExprID], col.[calcTrigger], 
		ISNULL(col.readOnly,0) AS [readOnly], 
		ISNULL(cont.readonly,0) AS [ScreenReadOnly],
		col.[statusBarMessage], col.[errorMessage], col.[linkTableID], col.[linkViewID],
		col.[linkOrderID], col.[Afdenabled], tab.[TableName],col.[Trimming], col.[Use1000Separator],
		col.[QAddressEnabled], col.[OLEType], col.[MaxOLESizeEnabled], col.[MaxOLESize], col.[AutoUpdateLookupValues],
		0 AS [locked]
	FROM [dbo].[ASRSysControls] cont
		LEFT OUTER JOIN [dbo].[ASRSysTables] tab ON cont.[tableID] = tab.[tableID]
		LEFT OUTER JOIN [dbo].[ASRSysColumns] col ON col.[tableID] = cont.[tableID] AND col.[columnID] = cont.[columnID]
	WHERE cont.[ScreenID] = @piScreenID
	ORDER BY cont.[PageNo], 
		cont.[ControlLevel] DESC, 
		cont.[tabIndex];
END
