CREATE PROCEDURE spASRGetOwnersForAllUtilities
AS
BEGIN

	SET NOCOUNT ON

	DECLARE @UserNames TABLE (UserName varchar(50))

	INSERT @UserNames
		SELECT Username FROM ASRSysBatchJobName
		UNION
		SELECT Username FROM ASRSysCalendarReports
		UNION
		SELECT Username FROM ASRSysCrossTab
		UNION
		SELECT Username FROM ASRSysCustomReportsName
		UNION
		SELECT Username FROM ASRSysDataTransferName
		UNION
		SELECT Username FROM ASRSysExportName
		UNION
		SELECT Username FROM ASRSysExpressions
		UNION
		SELECT Username FROM ASRSysGlobalFunctions
		UNION
		SELECT Username FROM ASRSysImportName
		UNION
		SELECT Username FROM ASRSysLabelTypes
		UNION
		SELECT Username FROM ASRSysMailMergeName
		UNION
		SELECT Username FROM ASRSysMatchReportName
		UNION
		SELECT Username FROM ASRSysPickListName
		UNION
		SELECT Username FROM ASRSysRecordProfileName

	SELECT DISTINCT UserName FROM @UserNames
	WHERE UserName <> '' AND UserName IS NOT NULL AND UserName <> 'sa'
	ORDER BY UserName

END
