CREATE PROCEDURE [dbo].[spASRResizeColumn]
	(@sTableName	varchar(255)
	,@ColumnName	varchar(255)
	,@Size			varchar(4))
AS
BEGIN

	DECLARE @iRecCount integer;
	DECLARE @NVarCommand nvarchar(MAX);

	-- Modify the passed in column to a varchar(max)
	SELECT @iRecCount = COUNT(id) FROM syscolumns
		where [id] = (select id from sysobjects where name = @sTableName)
		and [name] = @ColumnName;

	if @iRecCount = 1
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].[' + @sTableName + '] ' +
				'ALTER COLUMN [' + @ColumnName + '] varchar(' + @Size + ');'
		EXEC sp_executesql @NVarCommand;
	END
END