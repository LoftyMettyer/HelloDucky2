CREATE PROCEDURE [dbo].[sp_ASRInsertNewUtility]
(
    @piNewRecordID	integer OUTPUT,   /* Output variable to hold the new record ID. */
    @psInsertString nvarchar(MAX),    /* SQL Insert string to insert the new record. */
    @psTableName	varchar(255),		 /* Table Name you want to retrieve */
    @psIDColumnName varchar(30)      /* Name of the ID column  */
)
AS
BEGIN
    DECLARE @sCommand		nvarchar(MAX),
		@sParamDefinition 	nvarchar(MAX);

    BEGIN TRANSACTION;

    /* Run the given SQL INSERT string. */
    EXECUTE sp_ExecuteSQL @psInsertString;

    /* Get the ID of the inserted record.
    NB. We do not use @@IDENTITY as the insertion that we have just performed may have triggered
    other insertions (eg. into the Audit Trail table. The @@IDENTITY variable would then be the last IDENTITY value
    entered in the Audit Trail table.*/
    SET @sCommand = 'SELECT @recordID = MAX(' + @psIDColumnName + ') FROM ' + @psTableName + ';';

    SET @sParamDefinition = N'@recordID integer OUTPUT';
    EXEC sp_executesql @sCommand,  @sParamDefinition, @piNewRecordID OUTPUT;

    COMMIT TRANSACTION;

END