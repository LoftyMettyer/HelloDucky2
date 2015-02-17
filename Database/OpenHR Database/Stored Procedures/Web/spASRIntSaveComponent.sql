CREATE PROCEDURE spASRIntSaveComponent(
	@componentID integer,
	@expressionID integer,
	@type tinyint,
	@calculationID integer = 0,
	@filterID integer = 0,
	@functionID integer = 0,
	@operatorID integer = 0,
	@valueType tinyint = 0,
	@valueCharacter varchar(255) = '',
	@valueNumeric float = 0,
	@valueLogic bit = 0,
	@valuedate datetime = null,
	@LookupTableID integer = 0,
	@LookupColumnID integer = 0,
	@fieldTableID integer = 0,
	@fieldColumnID integer = 0,
	@fieldPassBy tinyint = 0,
	@fieldSelectionRecord tinyint = 0,
	@fieldSelectionLine integer = 0,
	@fieldSelectionOrderID integer = 0,
	@fieldSelectionFilter integer = 0,
	@promptDescription varchar(255) = '',
	@promptSize smallint = 0,
	@promptDecimals smallint = 0,
	@promptMask varchar(255) = '',
	@promptDateType integer = 0)
AS
BEGIN 

	SET NOCOUNT ON;

	INSERT INTO ASRSysExprComponents (componentID, exprID, [type]
					, calculationID, filterID, FunctionID, operatorID
					, ValueType, ValueCharacter, ValueNumeric, ValueLogic, ValueDate
					, LookupTableID, LookupColumnID
					, fieldTableID, fieldColumnID, fieldPassBy, fieldSelectionRecord, fieldSelectionLine, fieldSelectionOrderID, fieldSelectionFilter
					, promptDescription, promptSize, promptDecimals, promptMask, promptDateType)
				VALUES (@componentID, @expressionID, @type
					, @calculationID, @filterID, @functionID, @operatorID
					, @valueType, @valueCharacter, @valueNumeric, @valueLogic, @valuedate
					, @LookupTableID, @LookupColumnID
					, @fieldTableID, @fieldColumnID, @fieldPassBy, @fieldSelectionRecord, @fieldSelectionLine, @fieldSelectionOrderID, @fieldSelectionFilter
					, @promptDescription,	@promptSize, @promptDecimals, @promptMask, @promptDateType);

END

