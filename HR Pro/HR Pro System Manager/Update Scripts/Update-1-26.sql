
/* -------------------------------------------------- */
/* Update the database from version 26 to version 27. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@sDBVersion varchar(10),
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
        @GroupName varchar(8000),
        @AuditCommand nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */
SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

if @sDBVersion = ''
BEGIN
  SELECT @sDBVersion = SystemManagerVersion FROM ASRSysConfig
END


/* Exit if the database is not version 25 or 26. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.25') and (@sDBVersion <> '1.26')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 5 - Amending Import Definition Table'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'FilterID'
	AND sysobjects.name = 'ASRSysImportName'

IF @iRecCount = 0 
BEGIN
	ALTER TABLE [dbo].[ASRSysImportName] ADD [FilterID] int
END


/* -------------------------------------------- */

PRINT 'Step 2 of 5 - Adding new security'

/* Update CMG picture */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 20

IF @iRecCount = 1 
BEGIN
	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 20

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x424DF604000000000000360000002800000019000000100000000100180000000000C0040000130B0000130B00000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFECECFF1818FF8C8CFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0FF101000FFFFFFFFFFFFFF9898FF1010FF0000FF1414FF4040FFACACFF2424FF9090FFF0F0FF3030FF6868FFFFFFFFFFFFFF8888FF1818FF6262FF4040FF5454FF8C8CFFB8B8FF9898FF0404FF2C2C00FFFFFFFFFFFFFF0C0CFFB4B4FFFFFFFFF0F0FFB8B8FF8C8CFF2020FF4848FF5454FF0000FF0000FFACACFFFFFFFF3434FFC8C8FF2828FF0000FF0000FF0000FF0000FF0000FF0C0CFFD4D400FFFFFFFFFFFFFF0000FFF4F4FFFFFFFFFFFFFFFFFFFFFFFFFF2020FF0000FF0000FF8080FFA8A8FF1010FF8484FF1C1CFFFFFFFF8888FF2828FFECECFFC4C4FF4C4CFF0000FF4C4CFFFCFC00FFFFFFFFFFFFFF5858FF3434FFDCDCFFE0E0FFFFFFFFFFFFFF2020FF0000FF1010FFF8F8FFFFFFFF3030FF0000FF7070FFFFFFFFFFFFFF5454FF6868FFFFFFFFF4F4FF2828FF0000FF080800FFFFFFFFFFFFFFFCFCFF4040FF0000FF0000FFE4E4FFFFFFFF3C3CFF0000FF7878FFFFFFFFFFFFFFACACFF1818FFE0E0FFFFFFFFFFFFFFFCFCFF4C4CFF7070FFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFF4F4FF3838FF2C2CFFF8F8FFFFFFFFF0F0FFACACFFFCFCFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4F4FF3C3CFF7070FFFCFCFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4F4FF3C3CFF9494FFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00
END


/* Insert calculation category */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 21

IF @iRecCount = 0 
BEGIN
	SET IDENTITY_INSERT ASRSysPermissionCategories ON

	/* The record doesn't exist, so create it. */
	INSERT INTO ASRSysPermissionCategories
		(categoryID, description, picture, listOrder, categoryKey)
		VALUES(21,'Calculations','',10,'CALCULATIONS')

	SET IDENTITY_INSERT ASRSysPermissionCategories OFF

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 21

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000000000000680300001600000028000000100000002000000001001800000000004003000000000000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000808080808080808080808080808080808080808080808080808080808080808080808080808080808080808080000000808080C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0808080000000808080C0C0C0808080000000C0C0C0808080000000C0C0C0808080000000C0C0C0808080000000C0C0C0808080000000808080C0C0C0FFFFFF000000C0C0C0FFFFFF000000C0C0C0FFFFFF000000C0C0C0FFFFFF000000C0C0C0808080000000808080C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0808080000000808080C0C0C0808080000000C0C0C0808080000000C0C0C0808080000000C0C0C0808080000000C0C0C0808080000000808080C0C0C0FFFFFF000000C0C0C0FFFFFF000000C0C0C0FFFFFF000000C0C0C0FFFFFF000000C0C0C0808080000000808080C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0808080000000808080C0C0C0000000000000000000000000000000000000000000000000000000C0C0C0C0C0C0C0C0C0808080000000808080C0C0C0000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000C0C0C0C0C0C0C0C0C0808080000000808080C0C0C0000000000000000000000000000000000000000000000000000000C0C0C0C0C0C0C0C0C0808080000000808080C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0808080000000808080808080808080808080808080808080808080808080808080808080808080808080808080808080808080808080FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF
END


/* Insert configuration category */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 22

IF @iRecCount = 0 
BEGIN
	SET IDENTITY_INSERT ASRSysPermissionCategories ON

	/* The record doesn't exist, so create it. */
	INSERT INTO ASRSysPermissionCategories
		(categoryID, description, picture, listOrder, categoryKey)
		VALUES(22,'Configuration','',10,'CONFIGURATION')

	SET IDENTITY_INSERT ASRSysPermissionCategories OFF

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 22

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000000000000000000100000007000000001000007000000000010007000000000000107000000000000007000000000000007010000000000007000100003000000000001000330000000000003300000000000003330000000000000333000000000003300000000000000000000000000000000000000FFFF00009FCF00008F8F0000C71F0000E23F0000F07F0000F8FF0000F07B0000023100000700000067810000E7830000C7030000FC070000FE1F0000FFFF0000
END


/* Insert email option into the Event Log category */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionItems
WHERE ItemID = 88

IF @iRecCount = 0 
BEGIN
	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (88, 'Email', 40, 17,'EMAIL')
	INSERT INTO ASRSysGroupPermissions Select distinct 88,GroupName,Permitted from ASRSysGroupPermissions Where ItemID = 78
END


/* Insert options into the Calulation category, take the default security from the filters category */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionItems
WHERE ItemID = 89

IF @iRecCount = 0 
BEGIN
	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (89, 'New', 10, 21,'NEW')
	INSERT INTO ASRSysGroupPermissions Select distinct 89,GroupName,Permitted from ASRSysGroupPermissions Where ItemID = 52
	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (90, 'Edit', 20, 21,'EDIT')
	INSERT INTO ASRSysGroupPermissions Select distinct 90,GroupName,Permitted from ASRSysGroupPermissions Where ItemID = 53
	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (91, 'View', 30, 21,'VIEW')
	INSERT INTO ASRSysGroupPermissions Select distinct 91,GroupName,Permitted from ASRSysGroupPermissions Where ItemID = 72
	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (92, 'Delete', 40, 21,'DELETE')
	INSERT INTO ASRSysGroupPermissions Select distinct 92,GroupName,Permitted from ASRSysGroupPermissions Where ItemID = 54
END

/* Insert options into the Configuration category, give access to these to all security groups */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionItems
WHERE ItemID = 93

IF @iRecCount = 0 
BEGIN
	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (93, 'User', 10, 22,'USER')
	INSERT INTO ASRSysGroupPermissions Select distinct 93,GroupName,1 from ASRSysGroupPermissions

	INSERT INTO ASRSysPermissionItems (ItemID, Description, ListOrder, CategoryID, ItemKey) VALUES (94, 'PC', 10, 22,'PC')
	INSERT INTO ASRSysGroupPermissions Select distinct 94,GroupName,1 from ASRSysGroupPermissions

END


/* -------------------------------------------- */

PRINT 'Step 3 of 5 - Updating Messaging Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetMessages]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetMessages]

exec('CREATE PROCEDURE sp_ASRGetMessages AS
BEGIN
	DECLARE @iDBID		integer,
		@iID		integer,
		@dtLoginTime	datetime,
		@sLoginName	varchar(256),
		@iCount		integer

	/* Get the current user''s process information. */
	SELECT @iDBID = dbID,
		@dtLoginTime = login_time,
		@sLoginName = loginame
	FROM master..sysprocesses
	WHERE spid = @@spid

	/* Return the recordset of messages. */
	SELECT ''Message from user '''''' + ltrim(rtrim(messageFrom)) + 
		'''''' using '' + ltrim(rtrim(messageSource)) + 
		'' ('' + convert(varchar(100), messageTime, 100) +'')'' + 
		char(10) + char(10) + message
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND spid = @@spid
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime

	/* Remove any messages that have just been picked up. */
	DELETE
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND spid = @@spid
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime

	/* Remove any orphaned messages. */
	/* NB. This is done via a cursor to avoid any possible collation conflict between ASRSysMessages.loginName and sysprocesses.loginame. */
	DECLARE messages_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT id,
		loginName, 
		dbID, 
		loginTime 
	FROM ASRSysMessages
	OPEN messages_cursor
	FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM master..sysprocesses
		WHERE loginame =  @sLoginName
			AND dbID = @iDBID
			AND login_time = @dtLoginTime

		IF @iCount = 0
		BEGIN
			DELETE FROM ASRSysMessages 
			WHERE id = @iID
		END
			
		FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime
	END
	CLOSE messages_cursor 
	DEALLOCATE messages_cursor 

END')

/* -------------------------------------------- */

PRINT 'Step 4 of 5 - Amending Settings Table'

ALTER TABLE ASRSysSystemSettings ALTER COLUMN SettingValue varchar(200)


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 5 of 5 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.26')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.26')

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
/*
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)
*/

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.26 Of HR Pro'
