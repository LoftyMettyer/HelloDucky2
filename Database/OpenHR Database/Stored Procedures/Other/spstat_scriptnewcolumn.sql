CREATE PROCEDURE dbo.spstat_scriptnewcolumn (@columnid integer OUTPUT, @tableid integer, @columnname varchar(255)
	, @datatype integer, @description varchar(255), @size integer, @decimals integer, @islocked bit, @uniquekey varchar(37))
AS
BEGIN

	DECLARE @ssql nvarchar(MAX),
			@tablename varchar(255),
			@datasyntax	varchar(255);

	DECLARE @spinnerMinimum integer,
		@spinnerMaximum integer,
		@spinnerIncrement integer,
		@audit bit,
		@duplicate bit,
		@defaultvalue varchar(max),
		@columntype integer,
		@mandatory bit,
		@uniquecheck bit,
		@convertcase smallint,
		@mask varchar(MAX),
		@lookupTableID integer,
		@lookupColumnID integer,
		@controltype integer,
		@alphaonly bit,
		@blankIfZero bit,
		@multiline bit,
		@alignment smallint,
		@calcExprID integer,
		@gotFocusExprID integer,
		@lostFocusExprID integer,
		@calcTrigger smallint,
		@readOnly bit,
		@statusBarMessage varchar(255),
		@errorMessage varchar(255),
		@linkTableID integer, 
		@Afdenabled bit, 
		@Afdindividual integer,
		@Afdforename integer, 
		@Afdsurname integer,
		@Afdinitial integer, 
		@Afdtelephone integer, 
		@Afdaddress integer,
		@Afdproperty integer, 
		@Afdstreet integer, 
		@Afdlocality integer, 
		@Afdtown integer, 
		@Afdcounty integer,
		@dfltValueExprID integer, 
		@linkOrderID integer, 
		@OleOnServer bit, 
		@childUniqueCheck bit,
		@LinkViewID integer, 
		@DefaultDisplayWidth integer, 
		@UniqueCheckType integer,
		@Trimming integer, 
		@Use1000Separator bit,
		@LookupFilterColumnID integer, 
		@LookupFilterValueID integer, 
		@QAddressEnabled integer, 
		@QAIndividual integer, 
		@QAAddress integer, 
		@QAProperty integer, 
		@QAStreet integer,
		@QALocality integer, 
		@QATown integer, 
		@QACounty integer, 
		@LookupFilterOperator integer, 
		@Embedded bit, 
		@OLEType integer, 
		@MaxOLESizeEnabled bit, 
		@MaxOLESize integer,
		@AutoUpdateLookupValues bit, 
		@CalculateIfEmpty bit;

	-- Can we safely create this column?
	SELECT @columnid = ISNULL(columnid,0) FROM dbo.[ASRSysColumns] WHERE tableid = @tableid AND columnname = @columnname;
	IF @columnid > 0
	BEGIN
		RETURN;
	END

	SELECT @tablename = [tablename] FROM dbo.[ASRSysTables] WHERE tableid = @tableid;
	SELECT @columnid = MAX(columnid) + 1 FROM dbo.[ASRSysColumns];
	
	SET @defaultvalue = '';		
	SET @spinnerMinimum = 0;
	SET @spinnerMaximum = 0;
	SET @spinnerIncrement = 0;
	SET @audit = 0;
	SET @duplicate = 0;
	SET @columntype = 0;
	SET @mandatory = 0;
	SET @uniquecheck = 0;
	SET @convertcase = 0;
	SET @mask = '';
	SET @lookupTableID = 0;
	SET	@lookupColumnID = 0;
	SET	@controltype = 0;	
	SET @alphaonly = 0;
	SET @blankIfZero = 0;
	SET @multiline = 0;
	SET @alignment = 0;
	SET @calcExprID = 0;
	SET @gotFocusExprID = 0;
	SET @lostFocusExprID = 0;
	SET @calcTrigger = 0;
	SET @readOnly = 0;
	SET @statusBarMessage = '';
	SET @errorMessage = '';
	SET @linkTableID = 0; 
	SET @Afdenabled = 0; 
	SET @Afdindividual = 0;
	SET @Afdforename = 0; 
	SET @Afdsurname = 0;
	SET @Afdinitial = 0; 
	SET @Afdtelephone = 0; 
	SET @Afdaddress = 0;
	SET @Afdproperty = 0; 
	SET @Afdstreet = 0; 
	SET @Afdlocality = 0; 
	SET @Afdtown = 0; 
	SET @Afdcounty = 0;
	SET @dfltValueExprID = 0; 
	SET @linkOrderID = 0; 
	SET @OleOnServer = 0; 
	SET @childUniqueCheck = 0;
	SET @LinkViewID = 0; 
	SET @DefaultDisplayWidth = 0; 
	SET @UniqueCheckType = 0;
	SET @Trimming = 0;
	SET @Use1000Separator = 0;
	SET @LookupFilterColumnID = 0; 
	SET @LookupFilterValueID = 0; 
	SET @QAddressEnabled = 0; 
	SET @QAIndividual = 0; 
	SET @QAAddress = 0; 
	SET @QAProperty = 0; 
	SET @QAStreet = 0;
	SET @QALocality = 0; 
	SET @QATown = 0; 
	SET @QACounty = 0; 
	SET @LookupFilterOperator = 0; 
	SET @Embedded = 0; 
	SET @OLEType = 0; 
	SET @MaxOLESizeEnabled = 0; 
	SET @MaxOLESize = 0;
	SET @AutoUpdateLookupValues = 0; 
	SET @CalculateIfEmpty = 0;
	

	-- Logic
	IF @datatype = -7
	BEGIN
		SET @datasyntax = 'bit';
		SET @defaultvalue = 'FALSE';
		SET @controltype = 1;
	END

	-- OLE
	IF @datatype = -4
		SET @controltype = 1;

	-- Photo
	IF @datatype = -3
		SET @controltype = 1024;

	-- Link
	IF @datatype = -2
	BEGIN
		SET @datasyntax = 'varchar(255)';
		SET @controltype = 2048;
	END

	-- Working Pattern
	IF @datatype = -1
	BEGIN
		SET @datasyntax = 'varchar(14)';
		SET @controltype = 4096;
	END
	
	-- Numeric
	IF @datatype = 2
	BEGIN
		SET @datasyntax = 'numeric(' + convert(varchar(10),@size) + ',' + @decimals + ')';
		SET @defaultvalue = 0;	
		SET @controltype = 64;
	END

	-- Integers
	IF @datatype = 4
	BEGIN
		SET @datasyntax = 'integer';
		SET @controltype = 64;
	END
	
	-- Date
	IF @datatype = 11
	BEGIN
		SET @datasyntax = 'datetime';
		SET @controltype = 64;
	END

	-- Character
	IF @datatype = 12
	BEGIN
		SET @datasyntax = 'varchar(' + convert(varchar(10),@size) + ')';
		SET @controltype = 64;
	END

	-- System objects update
	INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
		SELECT @uniquekey, 2, @columnid, 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE', '01/01/1900',1,@islocked, GETDATE();

	-- Update base table								
	INSERT dbo.[tbsys_columns] ([columnID], [tableID], [columnType], [datatype], [defaultValue], [size], [decimals]
			, [lookupTableID], [lookupColumnID], [controltype], [spinnerMinimum], [spinnerMaximum], [spinnerIncrement], [audit]
			, [duplicate], [mandatory], [uniquecheck], [convertcase], [mask], [alphaonly], [blankIfZero], [multiline], [alignment]
			, [calcExprID], [gotFocusExprID], [lostFocusExprID], [calcTrigger], [readOnly], [statusBarMessage], [errorMessage]
			, [linkTableID], [Afdenabled], [Afdindividual], [Afdforename], [Afdsurname], [Afdinitial], [Afdtelephone], [Afdaddress]
			, [Afdproperty], [Afdstreet], [Afdlocality], [Afdtown], [Afdcounty], [dfltValueExprID], [linkOrderID], [OleOnServer]
			, [childUniqueCheck], [LinkViewID], [DefaultDisplayWidth], [ColumnName], [UniqueCheckType], [Trimming], [Use1000Separator]
			, [LookupFilterColumnID], [LookupFilterValueID], [QAddressEnabled], [QAIndividual], [QAAddress], [QAProperty], [QAStreet]
			, [QALocality], [QATown], [QACounty], [LookupFilterOperator], [Embedded], [OLEType], [MaxOLESizeEnabled], [MaxOLESize]
			, [AutoUpdateLookupValues], [CalculateIfEmpty]) 
		VALUES (@columnid, @tableid, @columntype, @datatype, @defaultvalue, @size, @decimals
			, @lookupTableID, @lookupColumnID, @controltype, @spinnerMinimum, @spinnerMaximum, @spinnerIncrement, @audit
			, @duplicate, @mandatory, @uniquecheck, @convertcase, @mask, @alphaonly, @blankIfZero, @multiline, @alignment
			, @calcExprID, @gotFocusExprID, @lostFocusExprID, @calcTrigger, @readOnly, @statusBarMessage, @errorMessage
			, @linkTableID, @Afdenabled, @Afdindividual, @Afdforename, @Afdsurname, @Afdinitial, @Afdtelephone, @Afdaddress
			, @Afdproperty, @Afdstreet, @Afdlocality, @Afdtown, @Afdcounty, @dfltValueExprID, @linkOrderID, @OleOnServer
			, @childUniqueCheck, @LinkViewID, @DefaultDisplayWidth, @ColumnName, @UniqueCheckType, @Trimming, @Use1000Separator
			, @LookupFilterColumnID, @LookupFilterValueID, @QAddressEnabled, @QAIndividual, @QAAddress, @QAProperty, @QAStreet
			, @QALocality, @QATown, @QACounty, @LookupFilterOperator, @Embedded, @OLEType, @MaxOLESizeEnabled, @MaxOLESize
			, @AutoUpdateLookupValues, @CalculateIfEmpty);

		-- Physically create this column (is regenerated by the System Manager save)	
		SET @ssql = N'ALTER TABLE dbo.tbuser_' + @tablename + ' ADD ' + @columnname + ' ' + @datasyntax;
		EXECUTE sp_executesql @ssql;

	RETURN;

END