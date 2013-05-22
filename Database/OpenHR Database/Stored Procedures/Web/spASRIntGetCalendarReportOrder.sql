CREATE PROCEDURE [dbo].[spASRIntGetCalendarReportOrder]
	(
	@piCalendarReportID 	integer, 
	@psErrorMsg				varchar(MAX)	OUTPUT
	)
AS
BEGIN
	DECLARE	@iCount	integer;

	SET @psErrorMsg = '';
	
	/* Check the calendar report exists. */
	SELECT @iCount = COUNT(*)
	FROM [dbo].[ASRSysCalendarReports]
	WHERE ID = @piCalendarReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'calendar report has been deleted by another user.';
		RETURN;
	END

	/* Check the calendar report has sort order details. */
	SELECT @iCount = COUNT(*)
	FROM [dbo].[ASRSysCalendarReportOrder]
	WHERE calendarReportID = @piCalendarReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'calendar report contains no sort order details.';
		RETURN;
	END
	
	SELECT 
		CONVERT(varchar,ASRSysCalendarReportOrder.ColumnID) + char(9) +
		(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ASRSysCalendarReportOrder.TableID) + 
		'.' + 
		(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportOrder.ColumnID) + char(9) +
		OrderType AS [OrderString]
	FROM [dbo].[ASRSysCalendarReportOrder]
	WHERE calendarReportID = @piCalendarReportID
	ORDER BY OrderSequence;
END