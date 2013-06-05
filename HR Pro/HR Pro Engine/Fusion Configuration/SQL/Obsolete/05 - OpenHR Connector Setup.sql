
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

	DECLARE @messageName varchar(255);

	DECLARE MessageCursor CURSOR LOCAL FAST_FORWARD FOR SELECT DISTINCT m.Name
		FROM ASRSysColumns c
		INNER JOIN fusion.element e ON e.ColumnID = c.ColumnID
		INNER JOIN fusion.messageelements me ON me.ElementID = e.ID		
		INNER JOIN fusion.message m ON me.MessageID = m.ID
		WHERE c.TableID = @tableID;
	
	OPEN MessageCursor;
	FETCH NEXT FROM MessageCursor INTO @messageName;
	WHILE @@FETCH_STATUS = 0 
	BEGIN 
		EXEC fusion.[pSendMessageCheckContext] @MessageType=@messageName, @LocalId=@RecordID
	    FETCH NEXT FROM MessageCursor INTO @messageName;
	END
	CLOSE MessageCursor;
	DEALLOCATE MessageCursor;

END


GO

CREATE FUNCTION fusion.makeXMLSafe(@input varchar(MAX))
	RETURNS VARCHAR(MAX)
	BEGIN
	RETURN 
		Replace(Replace(Replace(Replace(Replace(@input,'&','&amp;'),'<', '&lt;'),'>', '&gt;'),'"', '&quot;'), '''', '&#39;')
	END

GO

CREATE PROCEDURE [fusion].[pSetDataForMessage](@messagetype varchar(255), @id integer OUTPUT, @xml varchar(MAX), @parentguid varchar(255))
AS
BEGIN

	SET NOCOUNT ON;

DECLARE @xmlCode xml;

DECLARE @ParmDefinition nvarchar(500);
DECLARE @ssql nvarchar(MAX) = '0 AS ID',
		@sInsert nvarchar(MAX) = '0 AS ID',
		@sUpdate nvarchar(MAX) = '',
		@sColumns nvarchar(MAX),
		@sTableName nvarchar(255),
		@messagename nvarchar(255),
		@datanodeKey nvarchar(255),
		@foreignKeyName nvarchar(255),
		@foreignkeyvalue nvarchar(255),
		@executeCode nvarchar(MAX) = ''

SET @messagename = @messagetype

SET @xmlCode = convert(xml, @xml) 

SELECT @datanodeKey = DataNodeKey FROM fusion.message WHERE Name = @messagetype;

SELECT @foreignKeyName = 'ID_' + convert(varchar(4), c.TableID) FROM fusion.message m
	INNER JOIN fusion.[MessageRelations] mr ON mr.messageID = m.ID
	INNER JOIN fusion.[category] c ON c.ID = mr.categoryID
	WHERE mr.IsPrimaryKey = 0 AND m.name = @messagetype;

IF LEN(@foreignKeyName) > 0
BEGIN
	SELECT @foreignkeyvalue = LocalID FROM fusion.idtranslation WHERE busRef = @parentguid
	SET @foreignkeyvalue = ISNULL(@foreignkeyvalue,0)
END


-- Temp table
SET @ssql = 'DECLARE @mytable TABLE (ID integer'
SELECT @ssql = @ssql + ', [' + nodekey + '] nvarchar(MAX)'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename
SET @ssql = @ssql + ');'
SET @executeCode = @executeCode + @ssql + CHAR(13);

-- Insert
SET @sInsert = 'INSERT @mytable (ID '
SELECT @sInsert = @sInsert + ', [' + nodekey + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename
SET @sInsert = @sInsert + ')'

SET @ssql = '';
SELECT @ssql = @ssql + ',c.value(''nsWithXNS:' + nodekey + '[1]'', ''nvarchar(MAX)'') AS [' + nodekey + ']' + CHAR(13) 
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename

SET @ssql = 'WITH XMLNAMESPACES (''http://advancedcomputersoftware.com/xml/fusion/socialCare'' AS nsWithXNS)' + CHAR(13) +
	@sInsert +
	'SELECT 0' + @ssql + 'FROM @xmlCode.nodes(''nsWithXNS:' + @datanodeKey + ''') AS mytable(c)'

SET @executeCode = @executeCode + @ssql + CHAR(13);

SELECT TOP 1 @sTableName = t.tablename
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename

SELECT @sInsert = CASE WHEN LEN(@foreignKeyName) > 0 THEN @foreignKeyName ELSE '' END
SELECT @sInsert = @sInsert + CASE WHEN LEN(@sInsert) > 0 THEN ', ' ELSE '' END + ' [' + c.columnname + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sInsert= 'INSERT ' + @sTableName + ' ( ' + @sInsert + ') SELECT ';

SELECT @sColumns = CASE WHEN LEN(@foreignKeyName) > 0 THEN @foreignkeyvalue ELSE '' END
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

IF LEN(@executeCode) > 0
	EXEC sp_executeSQL @executeCode, @ParmDefinition, @xmlcode = @xmlcode, @id = @id
ELSE
	SELECT 0

END

go

CREATE PROCEDURE [fusion].[spGetDataForMessage](@messagetype varchar(255), @ID integer, @ID_Parent1 integer, @ID_Parent2 integer, @ID_Parent3 integer)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @xmlmessageID int = 1

	DECLARE @ssql nvarchar(MAX) = '',
			@ssql2 varchar(MAX) = '',
			@linesepcode nvarchar(255);

	DECLARE @xmlMessageBody	varchar(MAX),
			@xmllastmessage		varchar(MAX),
			@xmlMessageCode		varchar(MAX),
			@fusiontypeID		integer,
			@effectivefrom		datetime = '',
			@postguid			varchar(255),
			@tablename			varchar(255),
			@selectcode			varchar(MAX);

	DECLARE @messageID		int,
			@xmlns			varchar(255),
			@schemaLocation	varchar(255),
			@xmlxsi			varchar(255),
			@dataNodeKey	varchar(255),		
			@primaryKey		varchar(255),
			@foreignKey		varchar(MAX),
			@parentKey		varchar(255),
			@parentID		varchar(10),
			@version		int;
	
	-- Details for this message
	SELECT @messageID =	ID
			, @xmlns = [xmlns]
			, @schemaLocation = [xmlschemalocation]
			, @version = [version]
			, @xmlxsi = [xmlxsi]
			, @dataNodeKey = [DataNodeKey]
		FROM fusion.[message] WHERE name = @messagetype

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

	-- Get table name
	SELECT @tablename = t.tablename, @primaryKey = mr.NodeKey FROM fusion.MessageRelations mr
		INNER JOIN fusion.Category c ON c.ID = mr.CategoryID
		INNER JOIN ASRSysTables t ON t.TableID = c.TableID
		WHERE messageID = @messageID AND mr.IsPrimaryKey = 1

	-- Get relationship data
	SET @foreignKey = ''
	SELECT @foreignKey = NodeKey
		FROM fusion.MessageRelations mr
		INNER JOIN fusion.message m ON mr.MessageID = m.ID
		WHERE mr.IsPrimaryKey = 0 AND m.Name = @messagetype;

	SET @parentKey = '';
	SELECT @parentKey = 'ID_' + convert(varchar(10), c.TableID)
		FROM fusion.MessageRelations mr
		INNER JOIN fusion.message m ON mr.MessageID = m.ID
		INNER JOIN fusion.category c ON c.ID = mr.CategoryID
		WHERE mr.IsPrimaryKey = 0 AND m.Name = @messagetype;


	-- Build message body
	SET @ssql = '';
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
			WHERE xmlMessageID = @messagetype;


	IF LEN(@ssql) > 0 
	BEGIN
	
		SELECT @ssql = N'SELECT ' + CASE WHEN LEN(@parentKey) > 0 THEN ' @parentid = ' + @parentKey + ', ' ELSE '' END
			+ '@xml = ' + @ssql + ' FROM [' + @tablename + ']  WHERE ID = ' + convert(varchar(10),@ID)

		EXECUTE sp_executeSQL @ssql, N'@xml nvarchar(MAX) OUTPUT, @parentID int OUTPUT'
			, @xml = @xmlMessageBody OUTPUT, @parentID = @parentID OUTPUT;

		SELECT N'<?xml version="1.0" encoding="utf-8"?>
		<' + @messagetype + ' version="' + convert(varchar(2),@version) + '" ' + @primarykey + '="{0}" '
			+ CASE WHEN LEN(@foreignKey) > 0 THEN @foreignKey + '="{1}"' ELSE '' END +
			' xsi:schemaLocation="' + @schemaLocation + @messagetype + '.xsd"
			xmlns="' + @xmlns + '"
			xmlns:xsi="' + @xmlxsi + '">
			<data auditUserName="' + CURRENT_USER + '" recordStatus="Active" effectiveFrom="' + convert(varchar(10),@effectivefrom, 120) + '">
				<' + @dataNodeKey + '>'
				+ @xmlMessageBody +
				'</' + @dataNodeKey + '>
			</data>
			</' + @messagetype + '>' AS XMLCode
			, @parentID AS ParentID

	END
	ELSE
	BEGIN
		SELECT 'whoops - you ain''t configured this thing properly. Contact Harpenden QA on...'
	
	END

END

go
