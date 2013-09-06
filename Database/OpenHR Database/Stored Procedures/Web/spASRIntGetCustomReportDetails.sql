CREATE PROCEDURE spASRIntGetCustomReportDetails (@piCustomReportID integer)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT d.*, ISNULL(c.Use1000separator,0) AS Use1000separator
		FROM ASRSysCustomReportsDetails d
		LEFT JOIN ASRSysColumns c ON c.columnid = d.ColExprID And d.Type = 'C'
	WHERE CustomReportID = @piCustomReportID ORDER BY [Sequence];

END