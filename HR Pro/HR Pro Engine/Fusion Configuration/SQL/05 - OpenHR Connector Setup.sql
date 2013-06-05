
--DROP TABLE [fusion].[temptable]
--GO


--CREATE TABLE [fusion].[temptable](
--	[Message] [varchar](max) NULL,
--	[CreatedDateTime] [datetime] NULL)
	
 




----------------------------------------------------------------------------
-- OpenHR specifics
----------------------------------------------------------------------------

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spSendFusionMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[spSendFusionMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spGetDataForMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[spGetDataForMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSetDataForMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[pSetDataForMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[makeXMLSafe]') AND xtype = 'FN')
		DROP FUNCTION [fusion].[makeXMLSafe]

GO

CREATE PROCEDURE fusion.spSendFusionMessage(@TableID integer, @RecordID integer)
AS
BEGIN

	SET NOCOUNT ON;

	EXEC fusion.[pSendMessageCheckContext] @MessageType='StaffChange', @LocalId=@RecordID


END


GO

CREATE FUNCTION fusion.makeXMLSafe(@input varchar(MAX))
	RETURNS VARCHAR(MAX)
	BEGIN
	RETURN 
		Replace(Replace(Replace(Replace(Replace(@input,'&','&amp;'),'<', '&lt;'),'>', '&gt;'),'"', '&quot;'), '''', '&#39;')
	END

GO


--exec fusion.[pSetDataForMessage] 'staff', 12, 'sss'


CREATE PROCEDURE [fusion].[pSetDataForMessage](@messagetype varchar(255), @id integer OUTPUT, @xml varchar(MAX))
AS
BEGIN

	SET NOCOUNT ON;

	--INSERT [fusion].[temptable] (message) VALUES (@xml)

DECLARE @xmlCode xml;

--DECLARE @ID integer = 12

DECLARE @ParmDefinition nvarchar(500);
DECLARE @ssql nvarchar(MAX) = '0 AS ID',
		@sInsert nvarchar(MAX) = '0 AS ID',
		@sUpdate nvarchar(MAX) = '',
		@sColumns nvarchar(MAX),
		@sTableName nvarchar(MAX),
		@messagename nvarchar(MAX),
		@executeCode nvarchar(MAX) = ''

SET @messagename = 'staffChange'

--SELECT TOP 1 @xmlCode = convert(xml, message) FROM [fusion].[temptable] order by createddatetime desc
SET @xmlCode = convert(xml, @xml) 

-- Temp table
SET @ssql = 'DECLARE @mytable TABLE (ID integer'
SELECT @ssql = @ssql + ', ' + nodekey + ' nvarchar(MAX)'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename
SET @ssql = @ssql + ');'
SET @executeCode = @executeCode + @ssql + CHAR(13);


-- Insert
SET @sInsert = 'INSERT @mytable (ID '
SELECT @sInsert = @sInsert + ', ' + nodekey
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename
SET @sInsert = @sInsert + ')'

SET @ssql = '';
SELECT @ssql = @ssql + ',c.value(''nsWithXNS:' + nodekey + '[1]'', ''nvarchar(MAX)'') AS ' + nodekey + CHAR(13) 
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename

SET @ssql = 'WITH XMLNAMESPACES (''http://advancedcomputersoftware.com/xml/fusion/socialCare'' AS nsWithXNS)' + CHAR(13) +
	@sInsert +
	'SELECT 0' + @ssql + 'FROM @xmlCode.nodes(''nsWithXNS:staff'') AS mytable(c)'

SET @executeCode = @executeCode + @ssql + CHAR(13);

SELECT TOP 1 @sTableName = t.tablename
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename

SET @sInsert= '';
SELECT @sInsert = @sInsert + CASE WHEN LEN(@sInsert) > 0 THEN ', ' ELSE '' END + ' [' + c.columnname + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sInsert= 'INSERT ' + @sTableName + ' ( ' + @sInsert + ') SELECT ';

SET @sColumns = '';
SELECT @sColumns = @sColumns + CASE WHEN LEN(@sColumns) > 0 THEN ', ' ELSE '' END + '[' + e.NodeKey + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sInsert = @sInsert + @sColumns + ' FROM @mytable';

SELECT @sUpdate = @sUpdate + CASE WHEN LEN(@sUpdate) > 0 THEN ', ' ELSE '' END
	+ @sTableName + '.[' + c.columnname + '] = message.[' + e.nodekey + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sUpdate = 'UPDATE ' + @sTableName + ' SET ' + @sUpdate + ' FROM @mytable message WHERE ' + @sTableName + '.ID = @ID;'

SET @executeCode = @executeCode 
	+ 'IF (@ID > 0)' + CHAR(13)
	+ @sUpdate + CHAR(13) 
	+ ' ELSE ' + CHAR(13)
	+ ' BEGIN ' + CHAR(13)
	+ @sInsert  + CHAR(13)
	+ ' SELECT @ID = MAX(ID) FROM ' + @sTableName
	+ ' END' + CHAR(13)
	+ ' SELECT @ID;';

SET @ParmDefinition = N'@xmlCode xml, @ID integer OUTPUT';
EXEC sp_executeSQL @executeCode, @ParmDefinition, @xmlcode = @xmlcode, @id = @id




END

go

CREATE PROCEDURE [fusion].[spGetDataForMessage](@messagetype varchar(255), @ID integer, @ID_Parent1 integer, @ID_Parent2 integer, @ID_Parent3 integer)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @xmlmessageID INT = 1

	DECLARE @ssql nvarchar(MAX) = '',
			@ssql2 varchar(MAX) = '',
			@linesepcode nvarchar(255);

	DECLARE @xmlwholemessage	varchar(MAX),
			@xmllastmessage		varchar(MAX),
			@fusiontypeID		integer,
			@effectivefrom		datetime = '',
			@postguid			varchar(255),
			@selectcode			varchar(MAX);

	DECLARE @xmlcode TABLE (xmlmessageid smallint, xmlcode nvarchar(MAX))

	SET @linesepcode = ' + CHAR(13)+CHAR(10) + ''				'' + ';

	-- Get the creation date for this record
	--SELECT @effectivefrom = ISNULL(effectivefrom, GETDATE())
	--	FROM fusion.IdTranslation
	--	WHERE translationname = @messagetype AND localid = @ID;
	SELECT @effectivefrom = GETDATE();
	SET @ssql = '';

	---- Last XML message
	--SELECT TOP 1 @xmllastmessage = ISNULL(mt.LastGeneratedXml,'') 
	--	FROM fusion.messagetracking mt
	--		INNER JOIN fusion.IdTranslation tr ON tr.LocalId = @ID AND tr.BusRef = mt.BusRef
	--	WHERE tr.TranslationName = @messagetype
	--	ORDER BY mt.LastProcessedDate DESC;

	SET @xmllastmessage = '';

	--select columnid, ColumnName from ASRSysColumns where tableID = 1 order by columnname
	IF @messagetype = 'staffchange'
	BEGIN
		SELECT @ssql = @ssql + CASE LEN(@ssql) WHEN 0 THEN '' ELSE ' + ' END +
			CASE 
				WHEN NULLIF(x.value, '') IS NOT NULL
					THEN @linesepcode + '''<' + x.xmlnodekey + '>' + x.value + '</' + x.xmlnodekey + '>''' 

				WHEN c.datatype = 2
					THEN 'CASE ISNULL([' + c.ColumnName + '],0) WHEN 0 THEN '''' ELSE '
						+ @linesepcode + '''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.minoccurs = 0 AND x.nilable = 0 AND c.datatype = 11
					THEN 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN '''' ELSE '
						+ @linesepcode + '''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.minoccurs = 0 AND x.nilable = 0
					THEN 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN '''' ELSE '
						+ @linesepcode + '''<' + x.xmlnodekey + '>'' + fusion.makeXMLSafe([' + c.ColumnName + ']) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.nilable = 0 AND c.datatype = 11
					THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' +  x.xmlnodekey + '/>'' ELSE ''<' + x.xmlnodekey 
						+ '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.nilable = 0
					THEN + @linesepcode + '''<' + x.xmlnodekey + '>'' + ISNULL(fusion.makeXMLSafe([' + c.ColumnName + ']),'''') + ''</' + x.xmlnodekey + '>''' 

				WHEN x.nilable = 1 AND c.datatype = 11
					THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' + x.xmlnodekey 
						+ ' xsi:nil="true"/>'' ELSE ''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '],120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.nilable = 1
					THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' + x.xmlnodekey 
						+ ' xsi:nil="true"/>'' ELSE ''<' + x.xmlnodekey + '>'' + fusion.makeXMLSafe([' + c.ColumnName + ']) + ''</' + x.xmlnodekey + '>'' END' 
		 		
				ELSE '''UNKNOWN FIELD TYPE'''  END  
				-- + ' AS [column_' + convert(varchar(3), x.position) + ']'

			FROM [fusion].[MessageDefinition] x
				INNER JOIN ASRSysColumns c ON c.columnID = x.columnid
				INNER JOIN ASRSysTables t ON t.TableID = x.tableid

		select @ssql = N'SELECT @xml = ' + @ssql + ' FROM [personnel_records] WHERE ID = ' + convert(varchar(10),@ID)

print @ssql
		execute sp_executeSQL @ssql, N'@xml nvarchar(MAX) out', @xml = @xmlwholemessage output;

		SELECT N'<?xml version="1.0" encoding="utf-8"?>
		<staffChange version="1" staffRef="{0}" 
			xsi:schemaLocation="http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/staffChange.xsd"
			xmlns="http://advancedcomputersoftware.com/xml/fusion/socialCare"
			xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
			<data auditUserName="' + CURRENT_USER + '" recordStatus="Active" effectiveFrom="' + convert(varchar(10),@effectivefrom, 120) + '">
				<staff>'
				+ @xmlwholemessage +
				'</staff>
			</data>
			</staffChange>' AS xmlcode
			, @xmllastmessage AS xmllastmessage
			, '' AS [Column_0], '' AS [Column_1], '' AS [Column_2], '' AS [Column_3], '' AS [Column_4]
			, '' AS [Column_5], '' AS [Column_6], '' AS [Column_7], '' AS [Column_8]
			, '' AS [Column_9], '' AS [Column_10], '' AS [Column_11], '' AS [Column_12], '' AS [Column_13]
			, '' AS [Column_14], '' AS [Column_15], '' AS [Column_16], '' AS [Column_17], '' AS [Column_18]
			, '' AS [Column_19], '' AS [Column_20], '' AS [Column_21], '' AS [Column_22], '' AS [Column_23]
			, '' AS [Column_24], '' AS [Column_25], '' AS [Column_26], '' AS [Column_27], '' AS [Column_28], '' AS [Column_29]					
			
	END

	ELSE IF @messagetype = 'staffpostchange'
	BEGIN

		-- Get the creation date for this record
		SELECT @postguid = busref
			FROM fusion.IdTranslation
			WHERE translationname = 'staffpostchange' AND LocalId = @ID;

		IF @postguid IS NULL
		BEGIN
			SET @postguid = NEWID()
			INSERT fusion.IdTranslation (TranslationName, LocalId, BusRef) VALUES ('staffpostchange', @ID, @postguid)
		END


		SET @xmlwholemessage = '<?xml version="1.0" encoding="utf-8"?>
			<staffPostChange version="1" staffRef="{0}" staffPostRef="' + @postguid + '" 
				xsi:schemaLocation="http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/staffPostChange.xsd"
				xmlns="http://advancedcomputersoftware.com/xml/fusion/socialCare"
				xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
			<data auditUserName="' + CURRENT_USER + '" recordStatus="Active" effectiveFrom="{7}">
			<staffPost>
			  <name>{1}</name>
			  <department>{2}</department>
			  <site>{3}</site>
			  <contractedHoursPerWeek>{5}</contractedHoursPerWeek>
			  <maximumHoursPerWeek>{6}</maximumHoursPerWeek>
			</staffPost>
			</data>
			</staffPostChange>';	

		SELECT @xmlwholemessage AS xmlcode
			, @xmllastmessage AS xmllastmessage
			, convert(varchar(50),t.busref) AS [Column_0]
			, fusion.makeXMLSafe(ISNULL(a.[Duty_Type],'')) AS [Column_1]
			, fusion.makeXMLSafe(ISNULL(a.[Location],'')) AS [Column_2]
			, fusion.makeXMLSafe(ISNULL(a.[Division],'')) AS [Column_3]
			, '' AS [Column_4]
			, convert(varchar(20),ISNULL(a.[Actual_Hours],0)) AS [Column_5]
			, convert(varchar(20),ISNULL(a.[Post_Hours],0)) AS [Column_6]
			, ISNULL(convert(varchar(10),[Appointment_Start_Date], 120),'') AS [Column_7]
			, ISNULL(convert(varchar(10),[Appointment_End_Date], 120),'') AS [Column_8]		
			, '' AS [Column_9]
			, '' AS [Column_10]
			, '' AS [Column_11]
			, '' AS [Column_12]
			, '' AS [Column_13]
			, '' AS [Column_14]
			, '' AS [Column_15]
			, '' AS [Column_16]
			, '' AS [Column_17]
			, '' AS [Column_18]
			, '' AS [Column_19]
			, '' AS [Column_20]
			, '' AS [Column_21]
			, '' AS [Column_22]
			, '' AS [Column_23]
			, '' AS [Column_24]
			, '' AS [Column_25]
			, '' AS [Column_26]
			, '' AS [Column_27]
			, '' AS [Column_28]
			, '' AS [Column_29]			
		FROM Appointments a
			LEFT JOIN personnel_records p ON p.ID = a.ID_1
			LEFT JOIN fusion.IdTranslation t ON p.id = t.localid AND t.translationname = 'StaffChange'
		WHERE a.Id = @ID;
	
	END

END

go
