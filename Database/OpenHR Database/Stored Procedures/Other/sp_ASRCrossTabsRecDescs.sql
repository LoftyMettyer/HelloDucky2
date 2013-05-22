CREATE PROCEDURE [dbo].[sp_ASRCrossTabsRecDescs]
	(@tablename varchar(8000), @recordDescid int)
AS
BEGIN

	DECLARE @sSQL nvarchar(MAX);

	IF EXISTS (SELECT * FROM sysobjects WHERE type = 'P' AND name = 'sp_ASRExpr_' + convert(varchar,@RecordDescID))
	BEGIN
		SET @sSQL = '
			declare @tableid int;
			declare @recordid int;
			declare @recorddesc varchar(MAX);

			DECLARE table_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ID FROM '+ convert(nvarchar(MAX), @tablename) +'; 

			OPEN table_cursor;
			FETCH NEXT FROM table_cursor INTO @recordid;

			WHILE (@@fetch_status = 0)
			BEGIN
				exec sp_ASRExpr_' + convert(nvarchar(128),@RecordDescID) + ' @RecordDesc OUTPUT, @Recordid
				UPDATE ' + convert(nvarchar(128), @tablename) + ' SET RecDesc = @recordDesc WHERE id = @Recordid; 
				FETCH NEXT FROM table_cursor INTO @recordid
			END
			CLOSE table_cursor
			DEALLOCATE table_cursor';
		EXEC sp_executesql @ssql

	END

END
