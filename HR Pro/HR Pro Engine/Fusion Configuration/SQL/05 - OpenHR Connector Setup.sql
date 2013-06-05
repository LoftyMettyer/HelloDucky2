
----------------------------------------------------------------------------
-- OpenHR specifics
----------------------------------------------------------------------------

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spGetDataForMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[spGetDataForMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[makeXMLSafe]') AND xtype = 'FN')
		DROP FUNCTION [fusion].[makeXMLSafe]

GO


CREATE FUNCTION fusion.makeXMLSafe(@input varchar(MAX))
	RETURNS VARCHAR(MAX)
	BEGIN
	RETURN 
		Replace(Replace(Replace(Replace(Replace(@input,'&','&amp;'),'<', '&lt;'),'>', '&gt;'),'"', '&quot;'), '''', '&#39;')
	END

GO





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

	--DECLARE @xmldef TABLE (xmlmessageid smallint, xmlnodekey varchar(255), position tinyint
	--		, nilable bit, minoccurs bit
	--		, tableid integer, columnid integer
	--		, datatype tinyint, minsize integer, maxsize integer, value nvarchar(255))


	DECLARE @xmlcode TABLE (xmlmessageid smallint, xmlcode nvarchar(MAX))

	SET @linesepcode = ' + CHAR(13)+CHAR(10) + ''				'' + ';

	-- Get the creation date for this record
	--SELECT @effectivefrom = ISNULL(effectivefrom, GETDATE())
	--	FROM fusion.IdTranslation
	--	WHERE translationname = @messagetype AND localid = @ID;
	SELECT @effectivefrom = GETDATE();

	-- Last XML message
	SELECT TOP 1 @xmllastmessage = ISNULL(mt.LastGeneratedXml,'') FROM fusion.messagetracking mt
			INNER JOIN fusion.IdTranslation tr ON tr.LocalId = @ID AND tr.BusRef = mt.BusRef
		WHERE tr.TranslationName = @messagetype
		ORDER BY mt.LastProcessedDate DESC;

	--select columnid, ColumnName from ASRSysColumns where tableID = 1 order by columnname
	IF @messagetype = 'staffchange'
	BEGIN


--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (0, 1, 'title', 1, 1, 1, 13, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (1, 1, 'forenames', 0, 1, 1, 3, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 1, 'surname', 0, 1, 1, 2, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (3, 1, 'preferredName', 1, 0, 1, 20, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (4, 1, 'payrollNumber', 0, 0, 1, 2164, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (5, 1, 'DOB', 1, 0, 1, 12, 11)

--		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'employeeType', 1, 0, 1, 'Employee')
--		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'employmentStatus', 1, 0, 1, 'Active')

--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (8, 1, 'homePhoneNumber', 1, 0, 1, 29, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (9, 1, 'workMobile', 1, 0, 1, 1888, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (10, 1, 'personalMobile', 1, 0, 1, 1887, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (11, 1, 'email', 1, 0, 1, 531, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (12, 1, 'personalEmail', 1, 0, 1, 30, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (13, 1, 'addressLine1', 0, 1, 1, 23, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (14, 1, 'addressLine2', 0, 1, 1, 24, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (15, 1, 'addressLine3', 0, 1, 1, 25, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (16, 1, 'addressLine4', 0, 1, 1, 26, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (17, 1, 'addressLine5', 0, 1, 1, 27, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (18, 1, 'postCode', 0, 1, 1, 28, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (19, 1, 'gender', 0, 1, 1, 18, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (20, 1, 'startDate', 0, 1, 1, 14, 11)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (21, 1, 'leavingDate', 1, 0, 1, 15, 11)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (22, 1, 'leavingReason', 1, 0, 1, 17, 12)
----		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'companyName', 1, 0, 1, 'UNMAPPED FIELD')
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (24, 1, 'jobTitle', 0, 0, 1, 109, 12)
----		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'managerRef', 1, 0, 1, 'UNMAPPED FIELD')


--		-- Staff post change
--		--INSERT @xmldef (xmlmessageid, position, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 0, 'name', 1, 1, 1, 13, 12)
--		--INSERT @xmldef (xmlmessageid, position, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 0, 'department', 1, 1, 1, 13, 12)
--		--INSERT @xmldef (xmlmessageid, position, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 0, 'site', 1, 1, 1, 13, 12)
--		--INSERT @xmldef (xmlmessageid, position, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 0, 'siteManagerRef', 1, 1, 1, 13, 12)
--		--INSERT @xmldef (xmlmessageid, position, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 0, 'contractedHoursPerWeek', 1, 1, 1, 13, 12)
--		--INSERT @xmldef (xmlmessageid, position, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 0, 'maximumHoursPerWeek', 1, 1, 1, 13, 12)
		
				
		SELECT @ssql = @ssql + CASE LEN(@ssql) WHEN 0 THEN '' ELSE ' + ' END +
			CASE 
				WHEN x.value IS NOT NULL
					THEN @linesepcode + '''<' + x.xmlnodekey + '>' + x.value + '</' + x.xmlnodekey + '>''' 

				WHEN x.minoccurs = 0 AND x.nilable = 0 AND x.datatype = 11
					THEN 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN '''' ELSE '
						+ @linesepcode + '''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.minoccurs = 0 AND x.nilable = 0
					THEN 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN '''' ELSE '
						+ @linesepcode + '''<' + x.xmlnodekey + '>'' + fusion.makeXMLSafe([' + c.ColumnName + ']) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.nilable = 0 AND x.datatype = 11
					THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' +  x.xmlnodekey + '/>'' ELSE ''<' + x.xmlnodekey 
						+ '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.nilable = 0
					THEN + @linesepcode + '''<' + x.xmlnodekey + '>'' + ISNULL(fusion.makeXMLSafe([' + c.ColumnName + ']),'''') + ''</' + x.xmlnodekey + '>''' 

				WHEN x.nilable = 1 AND x.datatype = 11
					THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' + x.xmlnodekey 
						+ ' xsi:nil="true"/>'' ELSE ''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '],120) + ''</' + x.xmlnodekey + '>'' END' 

				WHEN x.nilable = 1
					THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' + x.xmlnodekey 
						+ ' xsi:nil="true"/>'' ELSE ''<' + x.xmlnodekey + '>'' + fusion.makeXMLSafe([' + c.ColumnName + ']) + ''</' + x.xmlnodekey + '>'' END' 
		 		
				ELSE '''UNKNOWN FIELD TYPE'''  END  
				-- + ' AS [column_' + convert(varchar(3), x.position) + ']'

			FROM [fusion].[MessageDefinition] x
			LEFT JOIN ASRSysColumns c ON c.columnID = x.columnid
			LEFT JOIN ASRSysTables t ON t.TableID = x.tableid

		--INSERT @xmlcode (xmlmessageid, xmlcode) VALUES (1, @ssql)
		--select TOP 1 @ssql2 = 'SELECT @xml = ' + xmlcode + 'FROM PERSONNEL_RECORDS WHERE ID = 17907'  from @xmlcode

		--print @ssql2
		
	--set @xmlwholemessage = ''
		select @ssql = N'SELECT @xml = ' + @ssql + ' FROM PERSONNEL_RECORDS WHERE ID = ' + convert(varchar(10),@ID)
		execute sp_executeSQL @ssql, N'@xml nvarchar(MAX) out', @xml = @xmlwholemessage output;

--print @xmlwholemessage


		--select TOP 1 @ssql = 'SELECT ' + @ssql + 'FROM PERSONNEL_RECORDS WHERE ID = 17907'
		--execute sp_executeSQL @ssql
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
