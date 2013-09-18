/* --------------------------------------------------- */
/* Update the database from version 7.0 to version 8.0 */
/* Stub file as this version has been skipped		   */
/* --------------------------------------------------- */

	EXEC spsys_setsystemsetting 'database', 'version', '8.0';
	EXEC spsys_setsystemsetting 'intranet', 'version', '8.0.12';


	-- TODO - all of it

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.0 Of OpenHR'
