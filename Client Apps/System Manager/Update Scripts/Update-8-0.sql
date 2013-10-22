/* --------------------------------------------------- */
/* Update the database from version 7.0 to version 8.0 */
/* Stub file as this version has been skipped		   */
/* --------------------------------------------------- */

	EXEC spsys_setsystemsetting 'database', 'version', '8.0';
	EXEC spsys_setsystemsetting 'intranet', 'version', '8.0.16';
	EXEC spsys_setsystemsetting 'ssintranet', 'version', '8.0.16';


	-- TODO - all of it

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.0 Of OpenHR'


/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE @sObject nvarchar(255)
DECLARE @sObjectType nvarchar(255)
DECLARE @NVarCommand nvarchar(MAX)
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
		 INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p' OR sysobjects.xtype = 'pc') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%' OR sysobjects.name LIKE 'spadmin%'))
		OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
		OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
		AND (sysusers.name = 'dbo')
--IF (@@ERROR <> 0) goto QuitWithRollback
 
OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
	IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'PC' OR rtrim(@sObjectType) = 'FN'
	BEGIN
		IF @sObject LIKE 'sp_ASRExpr_%' OR @sObject LIKE 'sp_ASRDfltExpr_%' OR @sObject LIKE 'spASREmail_%' OR @sObject LIKE 'spASRUpdateOLEField_%'
			SET @NVarCommand = 'REVOKE EXECUTE ON [' + @sObject + '] TO [ASRSysGroup]'
		ELSE              
			SET @NVarCommand = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
		END
	ELSE
	BEGIN
		SET @NVarCommand = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
	END
 
	EXECUTE sp_executeSQL @NVarCommand
	
	FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects
 

/* For the reset password functionality */
GRANT EXEC ON spadmin_commitresetpassword TO [openhr2iis]