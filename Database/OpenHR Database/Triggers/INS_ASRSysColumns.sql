CREATE TRIGGER [dbo].[INS_ASRSysColumns]
ON [dbo].[ASRSysColumns]
INSTEAD OF INSERT
AS
BEGIN

	SET NOCOUNT ON;

	-- Update objects table
	IF NOT EXISTS(SELECT [guid]
		FROM dbo.[tbsys_scriptedobjects] o
		INNER JOIN inserted i ON i.columnid = o.targetid AND o.objecttype = 2)
	BEGIN
		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT NEWID(), 2, [columnid], dbo.[udfsys_getownerid](), '01/01/1900',1,0, GETDATE()
				FROM inserted;
	END

	-- Update base table								
	INSERT dbo.[tbsys_columns] ([columnID], [tableID], [columnType], [datatype], [defaultValue], [size], [decimals], [lookupTableID], [lookupColumnID], [controltype], [spinnerMinimum], [spinnerMaximum], [spinnerIncrement], [audit], [duplicate], [mandatory], [uniquecheck], [convertcase], [mask], [alphaonly], [blankIfZero], [multiline], [alignment], [calcExprID], [gotFocusExprID], [lostFocusExprID], [calcTrigger], [readOnly], [statusBarMessage], [errorMessage], [linkTableID], [Afdenabled], [Afdindividual], [Afdforename], [Afdsurname], [Afdinitial], [Afdtelephone], [Afdaddress], [Afdproperty], [Afdstreet], [Afdlocality], [Afdtown], [Afdcounty], [dfltValueExprID], [linkOrderID], [OleOnServer], [childUniqueCheck], [LinkViewID], [DefaultDisplayWidth], [ColumnName], [UniqueCheckType], [Trimming], [Use1000Separator], [LookupFilterColumnID], [LookupFilterValueID], [QAddressEnabled], [QAIndividual], [QAAddress], [QAProperty], [QAStreet], [QALocality], [QATown], [QACounty], [LookupFilterOperator], [Embedded], [OLEType], [MaxOLESizeEnabled], [MaxOLESize], [AutoUpdateLookupValues], [CalculateIfEmpty]) 
		SELECT [columnID], [tableID], [columnType], [datatype], [defaultValue], [size], [decimals], [lookupTableID], [lookupColumnID], [controltype], [spinnerMinimum], [spinnerMaximum], [spinnerIncrement], [audit], [duplicate], [mandatory], [uniquecheck], [convertcase], [mask], [alphaonly], [blankIfZero], [multiline], [alignment], [calcExprID], [gotFocusExprID], [lostFocusExprID], [calcTrigger], [readOnly], [statusBarMessage], [errorMessage], [linkTableID], [Afdenabled], [Afdindividual], [Afdforename], [Afdsurname], [Afdinitial], [Afdtelephone], [Afdaddress], [Afdproperty], [Afdstreet], [Afdlocality], [Afdtown], [Afdcounty], [dfltValueExprID], [linkOrderID], [OleOnServer], [childUniqueCheck], [LinkViewID], [DefaultDisplayWidth], [ColumnName], [UniqueCheckType], [Trimming], [Use1000Separator], [LookupFilterColumnID], [LookupFilterValueID], [QAddressEnabled], [QAIndividual], [QAAddress], [QAProperty], [QAStreet], [QALocality], [QATown], [QACounty], [LookupFilterOperator], [Embedded], [OLEType], [MaxOLESizeEnabled], [MaxOLESize], [AutoUpdateLookupValues], [CalculateIfEmpty] FROM inserted;

END