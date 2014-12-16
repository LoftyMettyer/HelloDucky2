CREATE PROCEDURE [dbo].[spASRIntGetCustomReportDetails] (@piCustomReportID integer)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT d.*, ISNULL(c.Use1000separator,0) AS Use1000separator
			, ISNULL(c.columnname,'') AS [columnname]
			, ISNULL(t.tableid,0) AS [tableid]
			, ISNULL(t.tablename,'') AS [tablename]
			, CASE c.datatype WHEN 11 THEN 1 ELSE 0 END AS [IsDateColumn]
			, CASE c.datatype WHEN -7 THEN 1 ELSE 0 END AS [IsBooleanColumn]
			, c.datatype AS [DataType]
		FROM ASRSysCustomReportsDetails d
		LEFT JOIN ASRSysColumns c ON c.columnid = d.ColExprID And d.Type = 'C'
		LEFT JOIN ASRSysTables t ON c.tableid = t.tableid
	WHERE CustomReportID = @piCustomReportID ORDER BY [Sequence];

END