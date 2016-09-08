CREATE PROCEDURE [dbo].[spASRGetSQLMetadata](
	@sServerName nvarchar(128) OUTPUT,
	@sDBName nvarchar(128) OUTPUT)
	AS
	BEGIN
			SET @sServerName = CONVERT(nvarchar(128), SERVERPROPERTY('ServerName'));
			SET @sDBName = db_name();
	END