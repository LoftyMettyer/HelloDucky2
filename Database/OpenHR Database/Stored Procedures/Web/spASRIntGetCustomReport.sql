CREATE PROCEDURE dbo.spASRIntGetCustomReport
	(@ReportID integer)
AS
BEGIN

	SET NOCOUNT ON;
	
	-- Base report info
	SELECT c.[ID], c.[Name], c.[Description], c.[BaseTable], c.[AllRecords], c.[Picklist], c.[Filter]
		 , c.[Parent1Table], c.[Parent1Filter], c.[Parent2Table], c.[Parent2Filter], c.[Summary], c.[PrintFilterHeader]
		 , c.[UserName], c.[Timestamp], c.[Parent1AllRecords]
 		 , ISNULL(c.[Parent1Picklist],0) AS [Parent1Picklist]
		 , c.[Parent2AllRecords]
		 , ISNULL(c.[Parent2Picklist],0) AS [Parent2Picklist]
		 , c.[OutputPreview], c.[OutputFormat], c.[OutputScreen], c.[OutputPrinter], c.[OutputPrinterName], c.[OutputSave]
		 , c.[OutputSaveExisting], c.[OutputEmail], c.[OutputEmailAddr], c.[OutputEmailSubject], c.[OutputFilename]
		 , ISNULL(c.[OutputEmailAttachAs],0) AS [OutputEmailAttachAs]
		 , c.[IgnoreZeros]
		, t.tablename AS TableName
		, ISNULL(e.Name, '') AS EmailGroupName
		FROM dbo.ASRSYSCustomReportsName c
			INNER JOIN ASRSysTables t ON t.tableid = c.BaseTable
			LEFT JOIN ASRSysEmailGroupName e ON c.OutputEmailAddr = e.EmailGroupID
		WHERE c.ID = @ReportID;

	-- Child Report info
	SELECT C.ChildTable, C.ChildFilter, C.ChildMaxRecords, T.TableName, C.ChildOrder
		FROM ASRSYSCustomReportsChildDetails C
		      INNER JOIN ASRSysTables T ON T.TableID = C.ChildTable 
		WHERE C.CustomReportID = @ReportID;

END
