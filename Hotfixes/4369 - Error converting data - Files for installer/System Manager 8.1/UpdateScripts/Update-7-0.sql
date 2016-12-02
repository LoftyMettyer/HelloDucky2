/* --------------------------------------------------- */
/* Update the database from version 6.0 to version 7.0 */
/* Stub file as this version has been skipped		   */
/* --------------------------------------------------- */

	EXEC spsys_setsystemsetting 'database', 'version', '7.0';

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v7.0 Of OpenHR'
