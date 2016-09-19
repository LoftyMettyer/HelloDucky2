CREATE PROCEDURE [dbo].[spASRIntGetSessionSettings]
AS
BEGIN

	SET NOCOUNT ON;

	-- Declarations and their default values.
	DECLARE @BlockSize				integer = 1000,
			@PrimaryStartMode		tinyint = 3,
			@HistoryStartMode		tinyint = 3,
			@LookupStartMode		tinyint = 2,
			@QuickAccessStartMode	tinyint = 1,
			@ExprColourMode			integer	= 2,
			@ExprNodeMode			tinyint	= 1;

	DECLARE @SupportTelNo			varchar(50) = '+44 (0)8451 609 999',
			@SupportFax				varchar(50) = '+44 (0)1582 714814',
			@SupportEmail			varchar(50) = 'ohrsupport@oneadvanced.com',
			@SupportWebpage			varchar(50)	= 'https://customers.oneadvanced.com',
			@DesktopColour			varchar(20) = 2147483660;



	SELECT @BlockSize = settingValue
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER AND section = 'IntranetFindWindow' AND settingKey = 'BlockSize';

	SELECT @PrimaryStartMode = settingValue
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER AND section = 'RecordEditing' AND settingKey = 'Primary';

	SELECT @HistoryStartMode = settingValue
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER AND section = 'RecordEditing' AND settingKey = 'History';

	SELECT @LookupStartMode = settingValue
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER AND section = 'RecordEditing' AND settingKey = 'LookUp';

	SELECT @QuickAccessStartMode = settingValue
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER AND section = 'RecordEditing' AND settingKey = 'QuickAccess';

	SELECT @ExprNodeMode = settingValue
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER AND section = 'ExpressionBuilder' AND settingKey = 'NodeSize';
	
	SELECT @SupportTelNo = settingValue
		FROM ASRSysSystemSettings
		WHERE section = 'Support' AND settingKey = 'Telephone No';

	SELECT @SupportFax = settingValue
		FROM ASRSysSystemSettings
		WHERE section = 'Support' AND settingKey = 'Fax';

	SELECT @SupportEmail = settingValue
		FROM ASRSysSystemSettings
		WHERE section = 'Support' AND settingKey = 'Email';
		
	SELECT @SupportWebpage = settingValue
		FROM ASRSysSystemSettings
		WHERE section = 'Support' AND settingKey = 'WebPage';

	SELECT @DesktopColour = settingValue
		FROM ASRSysSystemSettings
		WHERE section = 'DesktopSetting' AND settingKey = 'BackgroundColour';


	SELECT @BlockSize			AS [BlockSize]
		, @PrimaryStartMode		AS [PrimaryStartMode]
		, @HistoryStartMode		AS [HistoryStartMode]
		, @LookupStartMode		AS [LookupStartMode]
		, @QuickAccessStartMode	AS [QuickAccessStartMode]
		, @ExprColourMode		AS [ExprColourMode]
		, @ExprColourMode		AS [ExprColourMode]
		, @ExprNodeMode			AS [ExprNodeMode]
		, @SupportTelNo			AS [SupportTelNo]
		, @SupportFax			AS [SupportFax]
		, @SupportEmail			AS [SupportEmail]
		, @SupportWebpage		AS [SupportWebpage]
		, @DesktopColour		AS [DesktopColour];


END
