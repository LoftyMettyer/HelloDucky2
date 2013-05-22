CREATE PROCEDURE [dbo].[sp_ASRIntGetReportChilds] (
	@piReportID		integer
)
AS
BEGIN
	/* Return the child table information based on the passed report ID */
	SELECT  
		CONVERT(varchar(255), C.ChildTable) + char(9) 
			+ T.TableName + char(9) 
			+ CONVERT(varchar(10), CASE 
				WHEN (X.access <> 'HD') OR (X.userName = system_user) THEN isnull(X.ExprID, 0)
				ELSE 0
			END) + char(9)
			+ CASE 
				WHEN (X.access <> 'HD') OR (X.userName = system_user) THEN isnull(X.Name, '')
				ELSE ''
			END + char(9)
			+ CONVERT(varchar(255), isnull(O.OrderID, 0)) + char(9) 
			+ isnull(O.Name, ' ' ) + char(9)
			+ CASE 
				WHEN C.ChildMaxRecords = 0 THEN 'All Records'
				ELSE CONVERT(varchar(100), C.ChildMaxRecords) 
			END AS [gridstring],
		C.ChildTable AS [TableID],
		T.TableName AS [Table],
		CASE 
			WHEN (X.access <> 'HD') OR (X.userName = system_user) THEN isnull(X.ExprID, 0)
			ELSE 0
		END AS [FilterID],
		CASE 
			WHEN (X.access <> 'HD') OR (X.userName = system_user) THEN isnull(X.Name, '')
			ELSE ''
		END AS [Filter],
		isnull(O.OrderID, 0) AS [OrderID],
	  O.Name AS [Order],
	  C.ChildMaxRecords AS [Records], 
		CASE 
			WHEN (X.access = 'HD') AND (X.userName = system_user) THEN 'Y'
			ELSE 'N'
		END AS [FilterHidden],
		CASE 
			WHEN isnull(O.OrderID, 0) <> isnull(C.ChildOrder,0) THEN 'Y'
			ELSE 'N'
		END AS [OrderDeleted],
		CASE 
			WHEN isnull(X.ExprID, 0) <> isnull(C.ChildFilter,0) THEN 'Y'
			ELSE 'N'
		END AS [FilterDeleted],
		CASE 
			WHEN (X.access = 'HD') AND (X.userName <> system_user) THEN 'Y'
			ELSE 'N'
		END AS [FilterHiddenByOther]
	FROM [dbo].[ASRSysCustomReportsChildDetails] C 
	INNER JOIN [dbo].[ASRSysTables] T ON C.ChildTable = T.TableID 
		LEFT OUTER JOIN [dbo].[ASRSysExpressions] X ON C.ChildFilter = X.ExprID 
		LEFT OUTER JOIN [dbo].[ASRSysOrders] O ON C.ChildOrder = O.OrderID
	WHERE C.CustomReportID = @piReportID
	ORDER BY T.TableName;
END