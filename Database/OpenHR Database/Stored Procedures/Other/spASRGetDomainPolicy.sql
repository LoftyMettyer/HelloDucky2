CREATE PROCEDURE [dbo].[spASRGetDomainPolicy]
	(@LockoutDuration int OUTPUT,
	 @lockoutThreshold int OUTPUT,
	 @lockoutObservationWindow int OUTPUT,
	 @maxPwdAge int OUTPUT, 
	 @minPwdAge int OUTPUT,
	 @minPwdLength int OUTPUT, 
	 @pwdHistoryLength int OUTPUT, 
	 @pwdProperties int OUTPUT)
AS
BEGIN

	SET NOCOUNT ON;

	-- Initialise the variables
	SET @LockoutDuration = 0;
	SET @lockoutThreshold  = 0;
	SET @lockoutObservationWindow  = 0;
	SET @maxPwdAge  = 0;
	SET @minPwdAge  = 0;
	SET @minPwdLength  = 0;
	SET @pwdHistoryLength  = 0;
	SET @pwdProperties  = 0;

	EXEC sp_executesql N'EXEC spASRGetDomainPolicyFromAssembly
			@lockoutDuration OUTPUT, @lockoutThreshold OUTPUT,
			@lockoutObservationWindow OUTPUT, @maxPwdAge OUTPUT,
			@minPwdAge OUTPUT, @minPwdLength OUTPUT,
			@pwdHistoryLength OUTPUT, @pwdProperties OUTPUT'
		, N'@lockoutDuration int OUT, @lockoutThreshold int OUT,
			@lockoutObservationWindow int OUT, @maxPwdAge int OUT,
			@minPwdAge int OUT,	@minPwdLength int OUT,
			@pwdHistoryLength int OUT, @pwdProperties int OUT'
		, @LockoutDuration OUT, @lockoutThreshold OUT
		, @lockoutObservationWindow OUT, @maxPwdAge OUT
		, @minPwdAge OUT, @minPwdLength OUT
		, @pwdHistoryLength OUT, @pwdProperties OUT;

END