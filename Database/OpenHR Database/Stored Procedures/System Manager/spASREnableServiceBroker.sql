CREATE PROCEDURE [dbo].[spASREnableServiceBroker]
AS
BEGIN
	DECLARE @sSQL nvarchar(MAX),
		@dbName	nvarchar(255) = DB_NAME(),
		@isBrokerEnabled bit = 0,
		@thisBrokerID uniqueidentifier,
		@uniqueBrokerCount integer;

	-- Is service broker enabled on this database?
	SELECT @isBrokerEnabled = is_broker_enabled, @thisBrokerID = service_broker_guid
		FROM sys.databases
		WHERE name = @dbName;

	-- Is it unique?
	SELECT @uniqueBrokerCount = COUNT(*)
		FROM sys.databases
		WHERE service_broker_guid = @thisBrokerID
		GROUP BY service_broker_guid;

	-- Enable if required
	IF @isBrokerEnabled = 0  OR (@isBrokerEnabled = 1 AND @uniqueBrokerCount > 1)
	BEGIN
		SET @sSQL = 'ALTER DATABASE [' + @dbName + '] SET NEW_BROKER WITH ROLLBACK IMMEDIATE;';
		EXEC sp_executeSQL @sSQL;
	END

END
