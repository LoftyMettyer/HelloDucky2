/* --------------------------------------------------- */
/* Update the database from version 5.2 to version 6.0 */
/* Stub file as this version has been skipped		   */
/* --------------------------------------------------- */

	EXEC spsys_setsystemsetting 'database', 'version', '6.0';

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v6.0 Of OpenHR'
