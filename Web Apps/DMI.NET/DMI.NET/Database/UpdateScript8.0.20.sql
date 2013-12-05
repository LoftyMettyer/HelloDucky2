
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetCustomReportDetails]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetCustomReportDetails];
GO


CREATE PROCEDURE spASRIntGetCustomReportDetails (@piCustomReportID integer)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT d.*, ISNULL(c.Use1000separator,0) AS Use1000separator
			, ISNULL(c.columnname,'') AS [columnname]
			, ISNULL(t.tableid,0) AS [tableid]
			, ISNULL(t.tablename,'') AS [tablename]
			, CASE c.datatype WHEN 11 THEN 1 ELSE 0 END AS [IsDateColumn]
			, CASE c.datatype WHEN -7 THEN 1 ELSE 0 END AS [IsBooleanColumn]
		FROM ASRSysCustomReportsDetails d
		LEFT JOIN ASRSysColumns c ON c.columnid = d.ColExprID And d.Type = 'C'
		LEFT JOIN ASRSysTables t ON c.tableid = t.tableid
	WHERE CustomReportID = @piCustomReportID ORDER BY [Sequence];

END

GO

DECLARE @sSQL nvarchar(MAX),
		@sGroup sysname,
		@sObject sysname,
		@sObjectType char(2);

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
		 INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
		OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
		OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
		AND (sysusers.name = 'dbo')

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
		IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
		BEGIN
				SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END
		ELSE
		BEGIN
				SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END

		FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects


EXEC spsys_setsystemsetting 'database', 'version', '8.0';
EXEC spsys_setsystemsetting 'intranet', 'version', '8.0.20';
EXEC spsys_setsystemsetting 'ssintranet', 'version', '8.0.20';
