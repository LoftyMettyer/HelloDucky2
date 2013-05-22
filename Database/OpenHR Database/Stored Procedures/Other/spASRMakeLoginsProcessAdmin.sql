CREATE PROCEDURE spASRMakeLoginsProcessAdmin
AS
BEGIN

	SET NOCOUNT OFF

	DECLARE @cursLogins cursor
	DECLARE @Mode smallint
	DECLARE @sName nvarchar(200)
	DECLARE @tmp_role_member_ids TABLE(id int not null, role_id int null, sub_role_id int null, generation int null)
	DECLARE @generation int

	SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'ProcessAccount' AND [SettingKey] = 'Mode'
	IF @@ROWCOUNT = 0 SET @Mode = 0

	SET @generation = 0

	INSERT INTO @tmp_role_member_ids (id)
		SELECT CAST(rl.uid AS int) AS [ID]
		FROM dbo.sysusers AS rl
		WHERE (rl.issqlrole = 1)and(rl.name=N'ASRSysGroup')

	UPDATE @tmp_role_member_ids SET role_id = id, sub_role_id = id, generation=@generation
	WHILE ( 1=1 )
	BEGIN
		INSERT INTO @tmp_role_member_ids (id, role_id, sub_role_id, generation)
			SELECT a.memberuid, b.role_id, a.groupuid, @generation + 1
				FROM sysmembers AS a INNER JOIN @tmp_role_member_ids AS b
				ON a.groupuid = b.id
				WHERE b.generation = @generation
		IF @@ROWCOUNT <= 0
			BREAK
		SET @generation = @generation + 1
	END

	DELETE @tmp_role_member_ids
	WHERE id in ( SELECT CAST(rl.uid AS int) AS [ID]
		FROM dbo.sysusers AS rl
		WHERE (rl.issqlrole = 1)and(rl.name=N'ASRSysGroup'))

	UPDATE @tmp_role_member_ids SET generation = 0;

	INSERT INTO @tmp_role_member_ids (id, role_id, generation) 
		SELECT distinct id, role_id, 1 FROM @tmp_role_member_ids

	DELETE @tmp_role_member_ids WHERE generation = 0

	SET @cursLogins = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR 
		SELECT u.name
			FROM dbo.sysusers AS rl
			INNER JOIN @tmp_role_member_ids AS m ON m.role_id=CAST(rl.uid AS int)
			INNER JOIN dbo.sysusers AS u ON u.uid = m.id
			WHERE (rl.issqlrole = 1)and(rl.name=N'ASRSysGroup')
    OPEN @cursLogins
    FETCH NEXT FROM @cursLogins INTO @sName

    WHILE (@@fetch_status = 0)
    BEGIN
		FETCH NEXT FROM @cursLogins INTO @sName

		PRINT '--' + @sName

		IF (@Mode = 3)
		BEGIN
			EXEC master..sp_addsrvrolemember @loginame = @sName, @rolename = N'processadmin'
		END
		ELSE
		BEGIN
			EXEC master..sp_dropsrvrolemember @loginame = @sName, @rolename = N'processadmin'
		END

	END
	
END		



