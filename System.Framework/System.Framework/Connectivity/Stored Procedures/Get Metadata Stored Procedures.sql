
	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spadmin_gettables]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spadmin_gettables];

	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spadmin_getcolumns]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spadmin_getcolumns];

	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spadmin_getscreens]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spadmin_getscreens];

	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spadmin_getexpressions]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spadmin_getexpressions];

	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spadmin_getcomponents]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spadmin_getcomponents];

	GO
	
	CREATE PROCEDURE dbo.spadmin_gettables
	AS
	BEGIN

		SELECT [TableID] AS [id]
		  ,[TableType]
		  ,[lastUpdated]
		  ,[DefaultOrderID]
		  ,[RecordDescExprID]
		  ,[DefaultEmailID]
		  ,[TableName] AS [name]
		  ,[ManualSummaryColumnBreaks]
		  ,[AuditInsert]
		  ,[AuditDelete]
		  ,[isremoteview]
		  ,0 AS [state]
		FROM dbo.[tbsys_tables];

	END

	GO

	CREATE PROCEDURE dbo.spadmin_getcolumns
	AS
	BEGIN
	
		SELECT [columnID] AS [id]
		  ,[tableID] AS [tableid]
		  ,[columnType]
		  ,[datatype]
		  ,[defaultValue]
		  ,[size]
		  ,[decimals]
		  ,[lookupTableID]
		  ,[lookupColumnID]
		  ,[controltype]
		  ,[spinnerMinimum]
		  ,[spinnerMaximum]
		  ,[spinnerIncrement]
		  ,[audit]
		  ,[duplicate]
		  ,[mandatory]
		  ,[uniquecheck]
		  ,[convertcase] AS [case]
		  ,[mask]
		  ,[alphaonly]
		  ,[blankIfZero]
		  ,[multiline]
		  ,[alignment]
		  ,[calcExprID] AS [calcid]
		  ,[gotFocusExprID]
		  ,[lostFocusExprID]
		  ,[calcTrigger]
		  ,[readOnly] AS [isreadonly]
		  ,[statusBarMessage]
		  ,[errorMessage]
		  ,[linkTableID]
		  ,[Afdenabled]
		  ,[Afdindividual]
		  ,[Afdforename]
		  ,[Afdsurname]
		  ,[Afdinitial]
		  ,[Afdtelephone]
		  ,[Afdaddress]
		  ,[Afdproperty]
		  ,[Afdstreet]
		  ,[Afdlocality]
		  ,[Afdtown]
		  ,[Afdcounty]
		  ,[dfltValueExprID] AS [defaultcalcid]
		  ,[linkOrderID]
		  ,[OleOnServer]
		  ,[childUniqueCheck]
		  ,[LinkViewID]
		  ,[DefaultDisplayWidth]
		  ,[ColumnName] AS [name]
		  ,[UniqueCheckType]
		  ,[Trimming]
		  ,[Use1000Separator]
		  ,[LookupFilterColumnID]
		  ,[LookupFilterValueID]
		  ,[QAddressEnabled]
		  ,[QAIndividual]
		  ,[QAAddress]
		  ,[QAProperty]
		  ,[QAStreet]
		  ,[QALocality]
		  ,[QATown]
		  ,[QACounty]
		  ,[LookupFilterOperator]
		  ,[Embedded]
		  ,[OLEType]
		  ,[MaxOLESizeEnabled]
		  ,[MaxOLESize]
		  ,[AutoUpdateLookupValues]
		  ,[CalculateIfEmpty]
		  ,'' AS [description]
		  ,0 AS [state]
	FROM dbo.[ASRSysColumns];
END

GO

CREATE PROCEDURE dbo.spadmin_getscreens
AS
BEGIN
	SELECT [ScreenID] AS [id]
		  ,[Name]
		  ,[TableID]
		  ,[OrderID]
		  ,[Height]
		  ,[Width]
		  ,[PictureID]
		  ,[FontName]
		  ,[FontSize]
		  ,[FontBold]
		  ,[FontItalic]
		  ,[FontStrikeThru]
		  ,[FontUnderline]
		  ,[GridX]
		  ,[GridY]
		  ,[AlignToGrid]
		  ,[DfltForeColour]
		  ,[DfltFontName]
		  ,[DfltFontSize]
		  ,[DfltFontBold]
		  ,[DfltFontItalic]
		  ,[QuickEntry]
		  ,[SSIntranet]
		  ,'' AS [description]
		  ,0 AS [state]
	  FROM [dbo].[ASRSysScreens];
	  
END

GO

CREATE PROCEDURE dbo.spadmin_getexpressions
AS
BEGIN

	SELECT [ExprID] AS [id]
		  ,[Name]
		  ,[ReturnType]
		  ,[ReturnSize] AS [size]
		  ,[ReturnDecimals] AS [decimals]
		  ,[Type]
		  ,[ParentComponentID]
		  ,[Description]
		  ,[Timestamp]
		  ,[TableID]
		  ,[Username]
		  ,[Access]
		  ,[ExpandedNode]
		  ,[ViewInColour]
		  ,[UtilityID]
		  ,0 AS [state]
	FROM [dbo].[ASRSysExpressions]
	WHERE ISNULL([tableID],0) > 0 AND ISNULL([ParentComponentID],0) = 0

END
GO


